# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Cement-тесты remediation-волны 02.11.2026 (независимое ревью REVIEW-DOCX-C).

Каждый тест — воспроизведение живого вектора из отчёта, а не теоретический кейс.
"""
from __future__ import annotations

import json
import os
import tempfile
import zipfile

import pytest

from docx_cli import inject_scan
from docx_cli.md_engine import _Parser, MdEngineError
from docx_cli.md_court import build_md_document
from docx_cli.spec import DocSpec
from docx_cli.builder import build_document


# --- helpers ---------------------------------------------------------------

def _make_docx(body_inner: str, styles_xml: str = "") -> str:
    tmp = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
    with zipfile.ZipFile(tmp.name, "w") as zf:
        zf.writestr(
            "word/document.xml",
            '<?xml version="1.0"?>'
            '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
            f"<w:body>{body_inner}</w:body></w:document>")
        if styles_xml:
            zf.writestr("word/styles.xml", styles_xml)
    return tmp.name


def _finding_types(path: str) -> list[str]:
    return [f["type"] for f in inject_scan.scan_all(path).get("findings", [])]


# --- S1-C1: пороги сканера --------------------------------------------------

def test_position_10pt_halfpt_caught():
    p = _make_docx('<w:r><w:rPr><w:position w:val="20"/></w:rPr><w:t>x</w:t></w:r>')
    assert "offset_text" in _finding_types(p)
    os.unlink(p)


def test_fractional_position_caught():
    p = _make_docx('<w:r><w:rPr><w:position w:val="36.0"/></w:rPr><w:t>x</w:t></w:r>')
    assert "offset_text" in _finding_types(p)
    os.unlink(p)


def test_font_5pt_caught():
    p = _make_docx('<w:r><w:rPr><w:sz w:val="10"/></w:rPr><w:t>x</w:t></w:r>')
    assert "microfont" in _finding_types(p)
    os.unlink(p)


def test_deep_basedon_chain_microfont_caught():
    styles = (
        '<?xml version="1.0"?>'
        '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:style w:type="character" w:styleId="s1"><w:basedOn w:val="s2"/></w:style>'
        '<w:style w:type="character" w:styleId="s2"><w:basedOn w:val="s3"/></w:style>'
        '<w:style w:type="character" w:styleId="s3"><w:basedOn w:val="s4"/></w:style>'
        '<w:style w:type="character" w:styleId="s4"><w:basedOn w:val="s5"/></w:style>'
        '<w:style w:type="character" w:styleId="s5"><w:rPr><w:sz w:val="8"/></w:rPr></w:style>'
        '</w:styles>')
    p = _make_docx('<w:r><w:rPr><w:rStyle w:val="s1"/></w:rPr><w:t>x</w:t></w:r>', styles)
    assert "microfont" in _finding_types(p)
    os.unlink(p)


def test_legit_superscript_not_flagged():
    p = _make_docx(
        '<w:r><w:rPr><w:position w:val="12"/>'
        '<w:vertAlign w:val="superscript"/></w:rPr><w:t>x</w:t></w:r>')
    assert _finding_types(p) == []
    os.unlink(p)


# --- S1-C2: split instrText -------------------------------------------------

def test_split_instrtext_macrobutton_caught():
    p = _make_docx(
        '<w:p><w:r><w:instrText xml:space="preserve"> MACRO</w:instrText></w:r>'
        '<w:r><w:instrText xml:space="preserve">BUTTON DoEvil </w:instrText></w:r></w:p>')
    assert "macrobutton_field" in _finding_types(p)
    os.unlink(p)


def test_split_instrtext_ddeauto_caught():
    p = _make_docx(
        '<w:p><w:r><w:instrText> DDE</w:instrText></w:r>'
        '<w:r><w:instrText>AUTO calc </w:instrText></w:r></w:p>')
    assert "dde_auto" in _finding_types(p)
    os.unlink(p)


def test_legit_page_numpages_fields_clean():
    for instr in (" PAGE ", " NUMPAGES ", r" REF sec-1 \h"):
        p = _make_docx(f'<w:p><w:fldSimple w:instr="{instr}">'
                       '<w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>')
        assert _finding_types(p) == [], f"legit field {instr!r} flagged"
        os.unlink(p)


# --- S1-H2: реальный io_error-фикс ------------------------------------------

def test_scanner_no_re_shadowing():
    """Регрессия: вложенный `import re` в _check_field/_domain_from_text
    делал re локальной переменной и сканер падал с UnboundLocalError."""
    p = _make_docx('<w:p><w:r><w:t>plain</w:t></w:r></w:p>')
    result = inject_scan.scan_all(p)
    assert not any(f["type"] == "io_error" for f in result.get("findings", [])), \
        "scanner io_error (re-shadowing regression)"
    os.unlink(p)


# --- S2-M1: маркеры сносок в heading/шапке ----------------------------------

def test_heading_footnote_marker_fail_closed():
    with pytest.raises(MdEngineError, match="footnote/endnote marker"):
        _Parser("## Раздел[^1]\n", "gost").parse()


def test_table_header_footnote_marker_fail_closed():
    md = "| Колонка[^1] | B |\n|---|---|\n| a | b |\n"
    with pytest.raises(MdEngineError, match="footnote/endnote marker"):
        _Parser(md, "gost").parse()


def test_heading_anchor_still_works():
    p = _Parser("## Раздел {#sec-1}\n\nтекст [см.](#sec-1)\n", "gost")
    p.parse()
    heading = [b for b in p.blocks if b.type == "heading"][0]
    assert "{#sec-1}" in heading.text


# --- S2-H1: GFM-выравнивание таблиц -----------------------------------------

def test_unescaped_pipe_row_not_killed():
    """GFM-reference: лишние ячейки отсекаются, неэкранированный '|'
    в тексте не гасит всю сборку."""
    p = _Parser("| A | B |\n|---|---|\n| ст. 12|ч. 2 | b |\n", "gost")
    p.parse()
    table = [b for b in p.blocks if b.type == "table"][0]
    assert len(table.rows[0]) == 2


def test_short_row_padded():
    p = _Parser("| A | B | C |\n|---|---|---|\n| a | b |\n", "gost")
    p.parse()
    table = [b for b in p.blocks if b.type == "table"][0]
    assert len(table.rows[0]) == 3
    assert table.rows[0][2].text == ""


def test_escaped_pipe_preserved():
    p = _Parser("| A | B |\n|---|---|\n| ст. 12\\|ч. 2 | b |\n", "gost")
    p.parse()
    table = [b for b in p.blocks if b.type == "table"][0]
    assert table.rows[0][0].text == "ст. 12|ч. 2"


# --- S2-M6: cross-ref на несуществующий якорь — warn ------------------------

def _minimal_spec(body_blocks):
    d = {
        "profile": "gost",
        "output": {"file_name": "/tmp/x.docx", "author": "t"},
        "header": {"align": "right", "lines": ["t"]},
        "title": {"main": "T"},
        "checks": {"no_yo": True, "no_en_dash": True, "require_norms_refs": False},
        "body": [{"type": "paragraph", "text": "x", "inline_markdown": False}],
    }
    spec = DocSpec.from_dict(d)
    spec.body = body_blocks
    return spec


def test_orphan_cross_ref_warns_not_ok():
    p = _Parser("[см.](#no-such-anchor)\n", "gost")
    p.parse()
    spec = _minimal_spec(p.blocks)
    out = "/tmp/cement-s2m6.docx"
    if os.path.exists(out):
        os.unlink(out)
    result = build_document(spec, out)
    assert result["status"] == "degraded", \
        f"orphan cross-ref должен давать degraded, got {result['status']}"
    assert any("cross_ref_target_missing" in w for w in result["metadata"]["warnings"])


def test_valid_cross_ref_stays_ok():
    p = _Parser("Раздел {#sec-1}\n\n[см.](#sec-1)\n", "gost")
    p.parse()
    spec = _minimal_spec(p.blocks)
    out = "/tmp/cement-s2m6-valid.docx"
    if os.path.exists(out):
        os.unlink(out)
    result = build_document(spec, out)
    assert result["status"] == "ok"
    assert not any("cross_ref_target_missing" in w for w in result["metadata"]["warnings"])


# --- S2-H5: дата во вложенном list-run --------------------------------------

def _fli(p):
    v = p.paragraph_format.first_line_indent
    return round(v.cm, 3) if v is not None else None


def test_date_after_list_breaks_run():
    d = build_md_document(
        "# ЗАЯВЛЕНИЕ\n1. Первое.\n2. Второе.\n1. января 2020 года договор.", "t")
    date_p = [p for p in d.paragraphs if "января" in p.text][0]
    v = _fli(date_p)
    assert v is None or v >= 0


def test_date_inside_nested_run_not_hanging():
    d = build_md_document(
        "# ЗАЯВЛЕНИЕ\n1. Первое.\n   1. января 2020 года договор.\n2. Второе.", "t")
    date_p = [p for p in d.paragraphs if "января" in p.text][0]
    v = _fli(date_p)
    assert v is None or v >= 0


def test_real_list_after_date_still_formatted():
    d = build_md_document(
        "# ЗАЯВЛЕНИЕ\n1. января 2020 года договор.\n1. Первое обстоятельство.", "t")
    item = [p for p in d.paragraphs if "Первое обстоятельство" in p.text][0]
    v = _fli(item)
    assert v is not None and v < 0
