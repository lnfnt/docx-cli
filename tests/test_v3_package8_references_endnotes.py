# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 8 (TASK_PACKAGE8_REFERENCES.md): концевые сноски (word/endnotes.xml),
закладки (w:bookmarkStart / w:bookmarkEnd), перекрестные ссылки (w:fldSimple REF / PAGEREF).
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile

import pytest
from docx import Document

from docx_cli.bookmarks import add_bookmark, add_cross_reference
from docx_cli.endnotes import add_endnote
from docx_cli.md_engine import MdEngineError, md_build_v2
from docx_cli.spec import DocSpec, OutputSpec, ParagraphBlock, Profile, TitleSpec, HeaderSpec

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
ENDNOTES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"
ENDNOTES_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"


def _zip_part(path, name: str) -> str:
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read(name).decode("utf-8")


def _names(path) -> set[str]:
    with zipfile.ZipFile(str(path)) as zf:
        return set(zf.namelist())


# ------------------------------------------------------------------------------
# 1. Тесты концевых сносок (Endnotes)
# ------------------------------------------------------------------------------

MD_ENDNOTES = """# Документ с концевыми сносками

Введение в теорию графов [^!graph] и примечание автора [^!note].

[^!graph]: См. монографию Оре О. Теория графов.
[^!note]: Замечание о нотации в главе 2.
"""

MD_ENDNOTES_PREFIX = """# Концевые сноски префикс e

Ссылка на источник [^e:src1] и дополнительный факт [^e:src2].

[^e:src1]: Источник 1.
[^e:src2]: Источник 2.
"""


def test_endnotes_package_structure(tmp_path):
    out = tmp_path / "endnotes.docx"
    res = md_build_v2(MD_ENDNOTES, str(out), "pytest")
    assert res["ok"] is True, res

    # 1. Часть word/endnotes.xml существует в пакете
    assert "word/endnotes.xml" in _names(out)

    # 2. Связь на endnotes.xml в word/_rels/document.xml.rels
    rels = _zip_part(out, "word/_rels/document.xml.rels")
    assert "endnotes.xml" in rels
    assert ENDNOTES_REL in rels

    # 3. Content-Type зарегистрирован в [Content_Types].xml
    assert ENDNOTES_CT in _zip_part(out, "[Content_Types].xml")

    # 4. Ссылки на концевые сноски в document.xml (нативный w:endnoteReference со стилем EndnoteReference)
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:endnoteReference w:id="1"/>' in doc_xml
    assert '<w:endnoteReference w:id="2"/>' in doc_xml
    assert 'w:val="EndnoteReference"' in doc_xml

    # 5. Текст сносок физически находится в word/endnotes.xml, вместе со служебными -1 и 0
    en_root = ET.fromstring(_zip_part(out, "word/endnotes.xml"))
    assert en_root.tag == f"{{{W_NS}}}endnotes"
    endnotes = en_root.findall(f"{{{W_NS}}}endnote")
    by_id = {e.get(f"{{{W_NS}}}id"): e for e in endnotes}
    assert by_id["-1"].get(f"{{{W_NS}}}type") == "separator"
    assert by_id["0"].get(f"{{{W_NS}}}type") == "continuationSeparator"

    all_text = "\n".join("".join(e.itertext()) for e in endnotes)
    assert "См. монографию Оре О. Теория графов." in all_text
    assert "Замечание о нотации в главе 2." in all_text

    # 6. В теле документа маркеры и текст сносок отсутствуют
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert not any("[^!" in t for t in texts)
    assert not any("См. монографию" in t for t in texts)


def test_endnotes_prefix_e_syntax(tmp_path):
    out = tmp_path / "endnotes_e.docx"
    res = md_build_v2(MD_ENDNOTES_PREFIX, str(out), "pytest")
    assert res["ok"] is True, res
    en_xml = _zip_part(out, "word/endnotes.xml")
    assert "Источник 1." in en_xml
    assert "Источник 2." in en_xml


def test_endnotes_fail_closed(tmp_path):
    # Концевая сноска без определения
    with pytest.raises(MdEngineError):
        md_build_v2("# T\n\nТекст [^!ref].\n", str(tmp_path / "e1.docx"), "pytest")
    # Определение концевой сноски без ссылки
    with pytest.raises(MdEngineError):
        md_build_v2("# T\n\nТекст.\n\n[^!ref]: Примечание.\n", str(tmp_path / "e2.docx"), "pytest")
    # Концевая сноска в ячейке таблицы — fail-closed
    with pytest.raises(MdEngineError):
        md_build_v2("# T\n\n| A | B |\n|---|---|\n| факт [^!ref] | 2 |\n\n[^!ref]: Примечание.\n",
                    str(tmp_path / "e3.docx"), "pytest")


# ------------------------------------------------------------------------------
# 2. Тесты закладок и перекрестных ссылок (Bookmarks & Cross-References)
# ------------------------------------------------------------------------------

MD_BOOKMARKS = """# Документ со ссылками

## Раздел определений {#bm_terms}

Здесь приводятся основные термины и определения.

## Аналитическая часть

Как указано в [разделе определений](#bm_terms), термины зафиксированы.
Номер страницы раздела: [стр.](#bm_terms) или на странице [@page:#bm_terms].
"""


def test_bookmarks_and_cross_references_md(tmp_path):
    out = tmp_path / "bookmarks.docx"
    res = md_build_v2(MD_BOOKMARKS, str(out), "pytest")
    assert res["ok"] is True, res

    doc_xml = _zip_part(out, "word/document.xml")

    # 1. Наличие w:bookmarkStart и w:bookmarkEnd с именем bm_terms
    assert '<w:bookmarkStart' in doc_xml
    assert 'w:name="bm_terms"' in doc_xml
    assert '<w:bookmarkEnd' in doc_xml

    # 2. Наличие w:fldSimple с инструкцией REF bm_terms \h и PAGEREF bm_terms \h
    assert 'w:instr="REF bm_terms \\h"' in doc_xml
    assert 'w:instr="PAGEREF bm_terms \\h"' in doc_xml


def test_bookmarks_api_direct(tmp_path):
    doc = Document()
    p1 = doc.add_paragraph()
    bm_id = add_bookmark(p1, "target_section", "Текст внутри закладки")
    assert bm_id == 1

    p2 = doc.add_paragraph("Ссылка: ")
    add_cross_reference(p2, "target_section", display_text="Перейти к разделу", page_ref=False)
    p2.add_run(", страница: ")
    add_cross_reference(p2, "target_section", page_ref=True)

    out = tmp_path / "direct_bm.docx"
    doc.save(str(out))

    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:bookmarkStart w:id="1" w:name="target_section"/>' in doc_xml
    assert '<w:bookmarkEnd w:id="1"/>' in doc_xml
    assert 'w:instr="REF target_section \\h"' in doc_xml
    assert 'w:instr="PAGEREF target_section \\h"' in doc_xml
    assert 'Перейти к разделу' in doc_xml


def test_both_footnotes_and_endnotes_together(tmp_path):
    md_both = """# Документ с обоими типами сносок

Текст с постраничной сноской [^f1] и концевой сноской [^!e1].

[^f1]: Постраничная сноска внизу.
[^!e1]: Концевая сноска в конце документа.
"""
    out = tmp_path / "both.docx"
    res = md_build_v2(md_both, str(out), "pytest")
    assert res["ok"] is True, res

    names = _names(out)
    assert "word/footnotes.xml" in names
    assert "word/endnotes.xml" in names

    fn_xml = _zip_part(out, "word/footnotes.xml")
    en_xml = _zip_part(out, "word/endnotes.xml")
    assert "Постраничная сноска внизу." in fn_xml
    assert "Концевая сноска в конце документа." in en_xml

    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:footnoteReference w:id="1"/>' in doc_xml
    assert '<w:endnoteReference w:id="1"/>' in doc_xml
