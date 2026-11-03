# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Фазы 2/3: md engine v2 (таблицы, H1-H4, выноски, сноски), профили
corporate/report, команда template list/show/build."""
from __future__ import annotations

import json
import zipfile
import xml.etree.ElementTree as ET

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from docx_cli.cli import main
from docx_cli.md_engine import MdEngineError, md_build_v2

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _xml(path) -> str:
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read("word/document.xml").decode("utf-8")


MD_TABLE = """# Отчет по продажам

## Показатели

| Показатель | Цель | Факт |
|:-----------|-----:|:----:|
| Выручка    | 100  | 112.5|
| Затраты    | 80   | 74.2 |
"""


def test_md_table_build(tmp_path):
    out = tmp_path / "table.docx"
    res = md_build_v2(MD_TABLE, str(out), "pytest")
    assert res["ok"] is True, res
    doc = Document(str(out))
    assert len(doc.tables) == 1
    t = doc.tables[0]
    assert len(t.rows) == 3
    assert t.rows[0].cells[0].text == "Показатель"
    assert t.rows[1].cells[1].text == "100"
    xml = _xml(out)
    # заливка шапки #F2F2F2
    assert 'w:fill="F2F2F2"' in xml
    # выравнивание по двоеточиям: колонка ы right, колонка ц center
    cell = t.rows[1].cells[1].paragraphs[0]
    assert cell.alignment == WD_ALIGN_PARAGRAPH.RIGHT
    head = t.rows[0].cells[0].paragraphs[0]
    # шапка центрируется (контракт v1: None -> шапка center, тело = по колонке)
    assert head.alignment == WD_ALIGN_PARAGRAPH.CENTER


def test_md_table_header_only_for_report_profile(tmp_path):
    out = tmp_path / "table_report.docx"
    md_build_v2(MD_TABLE, str(out), "pytest", profile="report")
    xml = _xml(out)
    # только линия под шапкой: нет внутрилайн границ Table Grid
    assert "TableGrid" not in xml
    assert 'w:bottom w:val="single"' in xml or "<w:tblBorders" not in xml


MD_HEADINGS = """# Титул документа

## Раздел

Текст раздела.

### Параграф

Текст параграфа.

#### Подпункт

Текст подпункта.
"""


def test_md_headings_h1_h4(tmp_path):
    out = tmp_path / "headings.docx"
    assert md_build_v2(MD_HEADINGS, str(out), "pytest")["ok"]
    doc = Document(str(out))
    texts = {p.text.strip(): p for p in doc.paragraphs}
    # H1 -> титул: bold 18pt
    title = texts["Титул документа"]
    assert all(r.bold and r.font.size.pt == 18 for r in title.runs if r.text)
    # H2-H4 -> стиль Заголовок N, размеры 16/14/12, bold, keep_with_next
    for level, text, size in ((2, "Раздел", 16), (3, "Параграф", 14),
                              (4, "Подпункт", 12)):
        p = texts[text]
        assert p.style.name == f"Heading {level}"
        assert all(r.bold and r.font.size.pt == size for r in p.runs if r.text)
        assert p.paragraph_format.keep_with_next is True


def test_md_no_leaked_markers(tmp_path):
    out = tmp_path / "markers.docx"
    md_build_v2(MD_TABLE, str(out), "pytest")
    doc = Document(str(out))
    texts = "".join(p.text for p in doc.paragraphs)
    assert "##" not in texts and "###" not in texts
    assert "| " not in texts and "|:--" not in texts


MD_CALLOUT = """# Предложение

> [!NOTE]
> Важная информация для клиента.

> [!WARNING]
> Специальные условия действуют до конца квартала.

> Простая цитата без маркера.

Обычный текст.
"""


def test_md_callout_note(tmp_path):
    out = tmp_path / "callout.docx"
    assert md_build_v2(MD_CALLOUT, str(out), "pytest")["ok"]
    doc = Document(str(out))
    texts = {p.text.strip(): p for p in doc.paragraphs}
    # [!NOTE] АКцент 0969DA, левой полосе и фону
    note = texts["Важная информация для клиента."]
    ppr = note._p.find(f"{{{W}}}pPr")
    assert ppr is not None
    ppr_xml = ET.tostring(ppr, encoding="unicode")
    assert '"0969DA"' in ppr_xml
    assert '<ns0:left ns0:val="single"' in ppr_xml
    assert '"F9F9F9"' in ppr_xml
    warn = texts["Специальные условия действуют до конца квартала."]
    warn_xml = ET.tostring(warn._p.find(f"{{{W}}}pPr"), encoding="unicode")
    assert '"9A6700"' in warn_xml
    quote = texts["Простая цитата без маркера."]
    quote_xml = ET.tostring(quote._p.find(f"{{{W}}}pPr"), encoding="unicode")
    assert '"2B579A"' in quote_xml
    assert "!!! NOTE" not in "".join(texts)


def test_md_bold_inline(tmp_path):
    out = tmp_path / "inline.docx"
    md_build_v2("# Т\n\nТекст с **жирным** и *курсивом*.\n", str(out), "pytest")
    doc = Document(str(out))
    runs = [r for p in doc.paragraphs for r in p.runs if r.text]
    by_text = {r.text: r for r in runs}
    assert by_text["жирным"].bold is True
    assert by_text["курсивом"].italic is True


MD_FOOTNOTES = """# Документ со сносками

Факт подтвержден [^1] и уточнен [^2].

[^1]: Источник данных: квартальный отчет.

[^2]: Методика расчета KPI.
"""


def test_md_footnotes(tmp_path):
    # ПАКЕТ 1: сноски — нативные OpenXML (footnotes.xml), а не список в конце.
    out = tmp_path / "footnotes.docx"
    res = md_build_v2(MD_FOOTNOTES, str(out), "pytest")
    assert res["ok"] is True, res
    with zipfile.ZipFile(str(out)) as zf:
        assert "word/footnotes.xml" in zf.namelist()
        fn_xml = zf.read("word/footnotes.xml").decode()
    assert "Источник данных: квартальный отчет." in fn_xml
    assert "Методика расчета KPI." in fn_xml
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    # маркеры [^N] убраны из текста
    assert not any("[^" in t for t in texts)
    # прежний суррогат («Сноски» списком в конце) больше не создаётся
    assert not any(t.strip() == "Сноски" for t in texts)


def test_md_footnotes_fail_closed(tmp_path):
    # ссылка без определения / определение без ссылки — ошибка сборки
    with pytest.raises(MdEngineError):
        md_build_v2("# Т\n\nТекст [^1].\n", str(tmp_path / "a.docx"), "pytest")
    with pytest.raises(MdEngineError):
        md_build_v2("# Т\n\nТекст.\n\n[^1]: нет ссылки.\n",
                    str(tmp_path / "b.docx"), "pytest")


def test_md_fail_closed_no_h1(tmp_path):
    with pytest.raises(MdEngineError):
        md_build_v2("Текст без титула\n", str(tmp_path / "x.docx"), "pytest")


def test_md_list_and_hr(tmp_path):
    out = tmp_path / "lists.docx"
    md_build_v2("# Д\n\n- пункт один\n- пункт два\n\n1. первый\n2. второй\n",
                str(out), "pytest")
    doc = Document(str(out))
    texts = [p.text.strip() for p in doc.paragraphs]
    styles = {p.text.strip(): p.style.name for p in doc.paragraphs}
    assert "пункт один" in texts
    assert styles.get("пункт один") == "List Bullet"
    assert "1. первый" in texts
    assert "2. второй" in texts


# --- профили corporate/report ------------------------------------------------

MD_SIMPLE = """# Корпоративный документ

Пояснительный текст к документу.
"""


def test_profile_corporate(tmp_path):
    out = tmp_path / "corp.docx"
    res = md_build_v2(MD_SIMPLE, str(out), "pytest")
    assert res["data"]["profile"] == "corporate"
    doc = Document(str(out))
    sec = doc.sections[0]
    assert abs(sec.left_margin.cm - 2.0) < 0.05
    assert abs(sec.top_margin.cm - 2.0) < 0.05
    names = {r.font.name for p in doc.paragraphs for r in p.runs if r.text.strip()}
    assert names == {"Arial"}
    # цвет заголовков профиля #1F4E79 (титул H1)
    assert "1F4E79" in _xml(out)


def test_profile_report(tmp_path):
    out = tmp_path / "report.docx"
    res = md_build_v2(MD_SIMPLE, str(out), "pytest", profile="report")
    assert res["data"]["profile"] == "report"
    names = {r.font.name for p in Document(str(out)).paragraphs
             for r in p.runs if r.text.strip()}
    assert names == {"Calibri"}


def test_profile_corporate_validation_passes(tmp_path):
    out = tmp_path / "corp_val.docx"
    res = md_build_v2(MD_TABLE, str(out), "pytest")
    # встроенная валидация builder'а (с спецификацией, корпоративный профиль)
    assert res["metadata"]["errors"] == [], res["metadata"]["errors"]
    assert res["metadata"]["warnings"] == [], res["metadata"]["warnings"]


# --- команда template --------------------------------------------------------

TEMPLATES = {"analytic_report", "commercial_offer", "service_contract",
             "meeting_minutes"}


def test_template_list():
    code = main(["template", "list"])
    assert code == 0


def test_template_list_content(capsys):
    main(["template", "list"])
    payload = json.loads(capsys.readouterr().out)
    names = {t["name"] for t in payload["data"]["templates"]}
    assert TEMPLATES <= names
    for t in payload["data"]["templates"]:
        assert t["description"] and t["profile"]


def test_template_show(capsys):
    code = main(["template", "show", "analytic_report"])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    data = payload["data"]
    assert data["name"] == "analytic_report"
    assert data["profile"] == "report"
    assert "# " in data["markdown"]


def test_template_show_unknown():
    assert main(["template", "show", "no_such_template"]) == 1


def test_template_build(tmp_path, capsys):
    out = tmp_path / "offer.docx"
    code = main(["template", "build", "commercial_offer", "-o", str(out)])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["ok"] is True
    assert payload["status"] == "ok"
    assert out.exists()
    doc = Document(str(out))
    texts = "".join(p.text for p in doc.paragraphs)
    assert "Коммерческое предложение" in texts
    fonts = {r.font.name for p in doc.paragraphs for r in p.runs if r.text.strip()}
    assert fonts == {"Arial"}


def test_template_build_all(tmp_path):
    for name in sorted(TEMPLATES):
        out = tmp_path / f"{name}.docx"
        assert main(["template", "build", name, "-o", str(out)]) == 0
        assert out.exists()


def test_template_build_profile_override(tmp_path, capsys):
    out = tmp_path / "contract.docx"
    code = main(["template", "build", "service_contract", "-o", str(out),
                 "--profile", "report", "--author", "agent-a"])
    assert code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["data"]["profile"] == "report"
    fonts = {r.font.name for p in Document(str(out)).paragraphs
             for r in p.runs if r.text.strip()}
    assert fonts == {"Calibri"}


def test_template_build_unknown(tmp_path):
    code = main(["template", "build", "no_such", "-o", str(tmp_path / "x.docx")])
    assert code == 1
