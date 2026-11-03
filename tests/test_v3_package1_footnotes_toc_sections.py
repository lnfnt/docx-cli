# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 1 (TASK_PACKAGE1_NATIVE_WORD.md): нативные сноски OpenXML,
автособираемое поле TOC (w:fldSimple + w:updateFields), разделы с реальной
сменой ориентации (landscape/portrait). Проверки структуры внутри ZIP-пакета.
"""
from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET

import pytest
from docx import Document
from docx.enum.section import WD_ORIENTATION

from docx_cli.md_engine import MdEngineError, md_build_v2

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FOOTNOTES_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
)


def _zip_part(path, name: str) -> str:
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read(name).decode("utf-8")


def _names(path) -> set[str]:
    with zipfile.ZipFile(str(path)) as zf:
        return set(zf.namelist())


# ------------------------------------------------------------------------------
# А. Нативные сноски OpenXML
# ------------------------------------------------------------------------------

MD_FOOTNOTES = """# Документ со сносками

Факт подтвержден [^1] и уточнен [^2].

[^1]: Источник данных: квартальный отчет.

[^2]: Методика расчета KPI.
"""


def test_footnotes_native_package_structure(tmp_path):
    out = tmp_path / "footnotes.docx"
    res = md_build_v2(MD_FOOTNOTES, str(out), "pytest")
    assert res["ok"] is True, res

    # 1. Часть word/footnotes.xml существует в пакете
    assert "word/footnotes.xml" in _names(out)

    # 2. Связь на footnotes.xml в word/_rels/document.xml.rels
    rels = _zip_part(out, "word/_rels/document.xml.rels")
    assert "footnotes.xml" in rels
    assert (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
        in rels
    )

    # 3. Content-Type зарегистрирован в [Content_Types].xml
    assert FOOTNOTES_CT in _zip_part(out, "[Content_Types].xml")

    # 4. Ссылки на сноски в document.xml (нативный w:footnoteReference)
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:footnoteReference w:id="1"/>' in doc_xml
    assert '<w:footnoteReference w:id="2"/>' in doc_xml

    # 5. Текст сносок физически находится в word/footnotes.xml, вместе
    #    с обязательными служебными сносками separator/continuationSeparator
    fn_root = ET.fromstring(_zip_part(out, "word/footnotes.xml"))
    assert fn_root.tag == f"{{{W_NS}}}footnotes"
    footnotes = fn_root.findall(f"{{{W_NS}}}footnote")
    by_id = {f.get(f"{{{W_NS}}}id"): f for f in footnotes}
    assert by_id["-1"].get(f"{{{W_NS}}}type") == "separator"
    assert by_id["0"].get(f"{{{W_NS}}}type") == "continuationSeparator"
    all_text = "\n".join("".join(f.itertext()) for f in footnotes)
    assert "Источник данных: квартальный отчет." in all_text
    assert "Методика расчета KPI." in all_text

    # 6. Никаких фейков в теле документа: маркеров и списка сносок нет
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert not any("[^" in t for t in texts)
    assert not any("Источник данных" in t for t in texts)


def test_footnotes_reference_superscript(tmp_path):
    out = tmp_path / "footnotes_run.docx"
    assert md_build_v2(MD_FOOTNOTES, str(out), "pytest")["ok"]
    doc = Document(str(out))
    target = next(p for p in doc.paragraphs if "подтвержден" in p.text)
    superscripts = target._p.findall(
        f"{{{W_NS}}}r/{{{W_NS}}}rPr/{{{W_NS}}}vertAlign")
    ref_runs = target._p.findall(f"{{{W_NS}}}r/{{{W_NS}}}footnoteReference")
    assert len(superscripts) == 2
    assert len(ref_runs) == 2


def test_footnotes_fail_closed(tmp_path):
    # ссылка без определения
    with pytest.raises(MdEngineError):
        md_build_v2("# Т\n\nТекст [^1].\n", str(tmp_path / "a.docx"), "pytest")
    # определение без ссылки
    with pytest.raises(MdEngineError):
        md_build_v2("# Т\n\nТекст.\n\n[^1]: нет ссылки.\n",
                    str(tmp_path / "b.docx"), "pytest")
    # маркер сноски вне абзаца (в ячейке таблицы) — fail-closed
    with pytest.raises(MdEngineError):
        md_build_v2(
            "# Т\n\n| A | B |\n|---|---|\n| факт [^1] | 2 |\n\n[^1]: с.\n",
            str(tmp_path / "c.docx"), "pytest")


# ------------------------------------------------------------------------------
# Б. Нативное оглавление TOC
# ------------------------------------------------------------------------------

MD_TOC = """# Отчет о работе

[TOC]

## Раздел первый

Текст первого раздела.

### Подраздел

Детали.

## Раздел второй

Текст второго раздела.
"""


def test_toc_field_and_update_fields(tmp_path):
    out = tmp_path / "toc.docx"
    res = md_build_v2(MD_TOC, str(out), "pytest")
    assert res["ok"] is True, res

    # w:fldSimple с инструкцией TOC \o "1-3" \h \z \u в document.xml
    doc_root = ET.fromstring(_zip_part(out, "word/document.xml"))
    instrs = [
        f.get(f"{{{W_NS}}}instr") or ""
        for f in doc_root.iter(f"{{{W_NS}}}fldSimple")
    ]
    assert any(
        i.startswith("TOC") and "1-3" in i and "\\h" in i and "\\z" in i and "\\u" in i
        for i in instrs
    ), instrs

    # автообновление полей при открытии: w:updateFields в word/settings.xml
    assert '<w:updateFields w:val="true"/>' in _zip_part(out, "word/settings.xml")


def test_toc_alias_and_title(tmp_path):
    out = tmp_path / "toc_alias.docx"
    res = md_build_v2("# Т\n\n[[toc]]\n\n## Глава\n\nТекст.\n",
                      str(out), "pytest")
    assert res["ok"] is True, res
    assert "TOC" in _zip_part(out, "word/document.xml")
    doc = Document(str(out))
    assert any(p.text.strip() == "Содержание" for p in doc.paragraphs)


def test_toc_update_fields_with_page_numbers(tmp_path):
    # поля страниц (PAGE/NUMPAGES) тоже требуют w:updateFields
    out = tmp_path / "pages.docx"
    res = md_build_v2("# Т\n\nОбычный текст.\n", str(out), "pytest")
    assert res["ok"] is True, res
    assert '<w:updateFields w:val="true"/>' in _zip_part(out, "word/settings.xml")


# ------------------------------------------------------------------------------
# В. Разделы и смена ориентации
# ------------------------------------------------------------------------------

MD_SECTIONS = """# Широкий отчет

Вводный текст первой секции (портрет).

<!-- section:landscape -->

| Показатель | Значение |
|:-----------|---------:|
| Выручка    | 100      |
| Затраты    | 80       |

<!-- section:portrait -->

Финальный текст третьей секции (портрет).
"""


def test_sections_orientation_landscape(tmp_path):
    out = tmp_path / "sections.docx"
    res = md_build_v2(MD_SECTIONS, str(out), "pytest")
    assert res["ok"] is True, res

    doc = Document(str(out))
    assert len(doc.sections) == 3
    sec1, sec2, sec3 = doc.sections
    # Секция 1 — портрет: высота больше ширины
    assert sec1.orientation == WD_ORIENTATION.PORTRAIT
    assert sec1.page_height > sec1.page_width
    # Секция 2 — ландшафт: ориентация LANDSCAPE и реальный обмен W/H
    assert sec2.orientation == WD_ORIENTATION.LANDSCAPE
    assert sec2.page_width > sec2.page_height
    assert abs(sec2.page_width.cm - 29.7) < 0.1
    assert abs(sec2.page_height.cm - 21.0) < 0.1
    # Секция 3 — возврат в портрет (обмен назад)
    assert sec3.orientation == WD_ORIENTATION.PORTRAIT
    assert sec3.page_height > sec3.page_width
    assert abs(sec3.page_width.cm - 21.0) < 0.1
    assert abs(sec3.page_height.cm - 29.7) < 0.1


def test_sections_table_in_landscape(tmp_path):
    out = tmp_path / "sections_table.docx"
    assert md_build_v2(MD_SECTIONS, str(out), "pytest")["ok"]
    doc = Document(str(out))
    assert len(doc.tables) == 1
    assert doc.tables[0].rows[0].cells[0].text == "Показатель"


def test_sections_custom_margins(tmp_path):
    out = tmp_path / "sections_margins.docx"
    res = md_build_v2(
        "# Т\n\nСтарт.\n\n<!-- section:landscape -->\n\nШирокий текст.\n",
        str(out), "pytest")
    assert res["ok"] is True, res
    doc = Document(str(out))
    assert len(doc.sections) == 2
    # не заданные поля наследуются из профиля (corporate: 2.0 см)
    assert abs(doc.sections[1].left_margin.cm - 2.0) < 0.05
