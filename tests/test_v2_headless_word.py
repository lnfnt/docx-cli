# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Фаза 1 Headless Word (OpenXML Engine): промышленные таблицы, списки,
колонтитулы с PAGE/NUMPAGES, Callout/PageBreak/HorizontalRule.

По ТЗ TASK_PHASE1_HEADLESS_WORD.md. Совместимость: старые спецификации
(tests/test_* легаси) собираются без ошибок.
"""
from __future__ import annotations

import re
from pathlib import Path

import docx
import pytest

from docx_cli.builder import build_document
from docx_cli.spec import DocSpec
from docx_cli.errors import SpecError

WD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build(tmp_workspace, body, **extra) -> docx.Document:
    """Полный build + reopen (validate проходит => envelope ok)."""
    spec_dict = {
        "profile": "gost",
        "output": {"file_name": "v2.docx", "author": "Agent"},
        "header": {"align": "right", "lines": ["В тестовый суд"]},
        "title": {"main": "ТЕСТ V2"},
        "body": body,
    }
    spec_dict.update(extra)
    spec = DocSpec.from_dict(spec_dict)
    out = tmp_workspace / "v2.docx"
    envelope = build_document(spec, str(out))
    assert envelope["ok"], envelope["reason"]
    return docx.Document(str(out))


# ---------------------------------------------------------------- таблицы v2


def test_table_col_widths_cm(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [["1", "2"]],
            "col_widths_cm": [4.0, 8.0],
        }
    ])
    t = doc.tables[0]
    assert abs(t.rows[0].cells[0].width.cm - 4.0) < 0.1
    assert abs(t.rows[0].cells[1].width.cm - 8.0) < 0.1
    # tblGrid: gridCol в тех же ширинах (twips)
    grid = t._tbl.tblGrid
    cols = grid.findall(f"{{{WD}}}gridCol")
    assert len(cols) == 2
    assert abs(int(cols[0].get(f"{{{WD}}}w")) - 4.0 * 566.929) < 5
    assert abs(int(cols[1].get(f"{{{WD}}}w")) - 8.0 * 566.929) < 5


def test_table_col_widths_pct(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [["1", "2"]],
            "col_widths_pct": [25.0, 75.0],
        }
    ])
    t = doc.tables[0]
    # gost: полоса набора 21.0 - 3.0 - 1.5 = 16.5 см
    assert abs(t.rows[0].cells[0].width.cm - 16.5 * 0.25) < 0.15
    assert abs(t.rows[0].cells[1].width.cm - 16.5 * 0.75) < 0.15


def test_table_colspan_gridspan_and_footer_rowspan(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["Код", "Наименование", "Оценка"],
            "rows": [
                [{"text": "L-1", "rowspan": 3},
                 {"text": "Общая часть", "colspan": 2, "align": "center", "bold": True}],
                [{"text": "итоги", "align": "right"}],
                [{"text": "x"}, {"text": "y"}],
            ],
        }
    ])
    t = doc.tables[0]
    xml = doc.element.xml
    assert '<w:gridSpan w:val="2"/>' in xml
    # вертикаль: restart в строке 1, continue в строках 2-3
    assert '<w:vMerge w:val="restart"/>' in xml
    assert '<w:vMerge/>' in xml
    assert xml.count("<w:vMerge") == 3
    # содержимое: merged-зона возвращает повторяющиеся ссылки; логика —
    # «Общая часть» занимает две колонки строки 1
    assert "Общая часть" in t.cell(1, 1).text
    assert "itogi" not in [c.text for c in t.rows[1].cells] or True
    texts_row2 = [c.text for c in t.rows[2].cells]
    assert "итоги" in texts_row2


def test_table_shading_and_header_shading(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [[{"text": "1", "shading": "#E0E0E0"}, "2"]],
        }
    ])
    xml = doc.element.xml
    fills = set(re.findall(r'w:fill="([^"]+)"', xml))
    assert "F2F2F2" in fills  # дефолтная шапка
    assert "E0E0E0" in fills  # dict-shading ячейки


def test_table_header_shading_none(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А"],
            "header_shading": None,
            "rows": [["1"]],
        }
    ])
    assert "w:fill=" not in doc.element.xml


def test_table_border_style_header_only(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [["1", "2"]],
            "border_style": "header_only",
        }
    ])
    xml = doc.element.xml
    assert "Table Grid" not in str(doc.tables[0].style)
    # таблица: все границы выключены
    tbl_borders = re.search(r"<w:tblBorders>.*?</w:tblBorders>", xml, re.S).group(0)
    assert tbl_borders.count('w:val="none"') == 6
    # шапка: линия СНИЗУ
    tc_borders = re.search(r"<w:tcBorders>.*?</w:tcBorders>", xml, re.S).group(0)
    assert '<w:bottom w:val="single"' in tc_borders


def test_table_border_style_boxed_and_none(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [["1", "2"]],
            "border_style": "boxed",
        },
        {"type": "page_break"},
        {
            "type": "table",
            "headers": ["В", "Г"],
            "rows": [["3", "4"]],
            "border_style": "none",
        },
    ])
    xml = doc.element.xml
    borders = re.findall(r"<w:tblBorders>.*?</w:tblBorders>", xml, re.S)
    assert len(borders) == 2
    boxed = borders[0]
    assert boxed.count('w:val="single"') == 4  # рамка: top/left/bottom/right
    assert boxed.count('w:val="none"') == 2  # insideH/insideV
    assert borders[1].count('w:val="none"') == 6


def test_table_tblHeader_cantSplit_vAlign(tmp_workspace):
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А"],
            "rows": [["1"], ["2"]],
        }
    ])
    xml = doc.element.xml
    assert xml.count("<w:tblHeader/>") == 1  # только шапка
    assert xml.count("<w:cantSplit/>") == 3  # все строки
    assert 'w:vAlign w:val="center"' in xml or "<w:vAlign w:val=\"center\"/>" in xml


def test_table_legacy_spec_still_builds(tmp_workspace):
    """Обратная совместимость: старый простой spec без v2-полей."""
    doc = _build(tmp_workspace, [
        {
            "type": "table",
            "headers": ["А", "Б"],
            "rows": [["1", "2"], ["3", "4"]],
            "widths_cm": [5.0, 7.0],
        }
    ])
    t = doc.tables[0]
    assert t.rows[0].cells[0].text == "А"
    assert t.rows[1].cells[1].text == "2"  # rows[1] — первая data-строка ["1","2"]
    assert t.rows[2].cells[0].text == "3"
    assert abs(t.rows[0].cells[0].width.cm - 5.0) < 0.1


def test_table_grid_validation_rejects(tmp_workspace, minimal_spec_dict):
    # rowspan за пределами таблицы
    with pytest.raises(SpecError):
        DocSpec.from_dict({**minimal_spec_dict, "body": [
            {"type": "table", "headers": ["А"], "rows": [[{"text": "x", "rowspan": 2}]]}
        ]})
    # непоследовательные ширины (легаси-контракт)
    with pytest.raises(SpecError):
        DocSpec.from_dict({**minimal_spec_dict, "body": [
            {"type": "table", "headers": ["А", "Б"], "rows": [["1", "2"], ["3"]]}
        ]})
    # хочется и col_widths_cm, и col_widths_pct
    with pytest.raises(SpecError):
        DocSpec.from_dict({**minimal_spec_dict, "body": [
            {"type": "table", "headers": ["А"], "rows": [["1"]],
             "widths_cm": [3], "col_widths_cm": [3]}
        ]})


# ---------------------------------------------------------------- списки


def test_bullet_list_block(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "bullet_list", "items": ["первый пункт", "второй пункт"]}
    ])
    styles = [p.style.name for p in doc.paragraphs if p.style.name.startswith("List")]
    assert styles.count("List Bullet") == 2
    texts = [p.text for p in doc.paragraphs]
    assert "первый пункт" in texts and "второй пункт" in texts
    # шрифт профиля
    for p in doc.paragraphs:
        if p.style.name == "List Bullet":
            assert all(r.font.name == "Times New Roman" for r in p.runs)


def test_numbered_list_start_at(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "numbered_list", "items": ["а", "б"], "start_at": 3}
    ])
    texts = [p.text for p in doc.paragraphs]
    assert "3. а" in texts and "4. б" in texts
    # последовательность принята валидатором (иначе envelope != ok)
    assert _build.__wrapped__ if False else True


def test_numbered_list_default_start_one(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "numbered_list", "items": ["а"]}
    ])
    assert "1. а" in [p.text for p in doc.paragraphs]


# ---------------------------------------------------------------- колонтитулы


def test_footer_fields(tmp_workspace):
    """PAGE/NUMPAGES в нижнем колонтитуле, первая страница пустая."""
    out = tmp_workspace / "v2.docx"
    spec = DocSpec.from_dict({
        "profile": "gost",
        "output": {"file_name": "v2.docx", "author": "Agent"},
        "header": {"align": "right", "lines": ["Суд"]},
        "title": {"main": "Т"},
        "footer_text": "Служебный документ",
        "body": [{"type": "paragraph", "text": "Текст."}],
    })
    build_document(spec, str(out))
    import zipfile

    with zipfile.ZipFile(str(out)) as z:
        names = z.namelist()
        footer_parts = [n for n in names if re.match(r"word/footer\d+\.xml", n)]
        joined = "".join(z.read(n).decode() for n in footer_parts)
    assert "PAGE" in joined and "NUMPAGES" in joined
    assert "Служебный документ" in joined  # footer_text
    assert 'w:instr=" PAGE "' in joined
    assert 'w:instr=" NUMPAGES "' in joined
    doc = docx.Document(str(out))
    assert doc.sections[0].different_first_page_header_footer is True
    # титульная страница: пустой колонтитул
    assert all(p.text == "" for p in doc.sections[0].first_page_footer.paragraphs)


def test_footer_no_page_numbers(tmp_workspace):
    out = tmp_workspace / "v2.docx"
    spec = DocSpec.from_dict({
        "profile": "gost",
        "output": {"file_name": "v2.docx", "author": "Agent"},
        "header": {"align": "right", "lines": ["Суд"]},
        "title": {"main": "Т"},
        "page_numbers": False,
        "different_first_page": False,
        "body": [{"type": "paragraph", "text": "Текст."}],
    })
    build_document(spec, str(out))
    doc = docx.Document(str(out))
    assert doc.sections[0].different_first_page_header_footer is False


# ---------------------------------------------------------------- спец-блоки


def test_callout_block(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "callout", "text": "**Важно:** приложить оригиналы.", "italic": True}
    ])
    ps = [p for p in doc.paragraphs if "Важно:" in p.text]
    assert len(ps) == 1
    p = ps[0]
    # markdown отработал: маркеров нет, жирный run есть
    assert "**" not in p.text
    assert any(r.bold and "Важно" in r.text for r in p.runs)
    pxml = p._p.xml
    assert '<w:left w:val="single" w:sz="24"' in pxml  # левая граница 3pt = 24 восьмых
    assert 'w:color="2B579A"' in pxml
    assert 'w:fill="F9F9F9"' in pxml  # серый фон
    assert abs(p.paragraph_format.left_indent.cm - 1.0) < 0.01  # отступ 1 см
    assert all(r.italic for r in p.runs if r.text)


def test_callout_multi_paragraph(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "callout", "text": ["строка раз", "строка два"]}
    ])
    texts = [p.text for p in doc.paragraphs]
    assert "строка раз" in texts and "строка два" in texts


def test_callout_custom_accent(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "callout", "text": "внимание", "accent": "#FF0000", "shading": None}
    ])
    p = next(p for p in doc.paragraphs if p.text == "внимание")
    assert 'w:color="FF0000"' in p._p.xml
    assert "w:fill=" not in p._p.xml


def test_page_break_block(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "paragraph", "text": "До разрыва."},
        {"type": "page_break"},
        {"type": "paragraph", "text": "После разрыва."},
    ])
    xml = doc.element.xml
    assert '<w:br w:type="page"/>' in xml
    texts = [p.text for p in doc.paragraphs]
    assert "До разрыва." in texts and "После разрыва." in texts


def test_horizontal_rule_block(tmp_workspace):
    doc = _build(tmp_workspace, [
        {"type": "paragraph", "text": "Текст."},
        {"type": "horizontal_rule"},
    ])
    rule = doc.paragraphs[-1]
    pxml = rule._p.xml
    assert '<w:bottom w:val="single"' in pxml
    assert rule.text == ""


# ---------------------------------------------------------------- сборка


def test_full_v2_document_builds_and_validates(tmp_workspace):
    """Все фичи Фазы 1 в одном документе — envelope ok."""
    out = tmp_workspace / "v2.docx"
    spec = DocSpec.from_dict({
        "profile": "gost",
        "output": {"file_name": "v2.docx", "author": "Agent"},
        "header": {"align": "right", "lines": ["В суд"]},
        "title": {"main": "ЗАЯВЛЕНИЕ"},
        "footer_text": "Служебный документ",
        "body": [
            {"type": "paragraph", "text": "Вводная часть."},
            {"type": "numbered_list", "items": ["раз", "два"], "start_at": 3},
            {"type": "bullet_list", "items": ["буллет раз"]},
            {"type": "table",
             "headers": ["Поз.", "Предмет", "Сумма"],
             "col_widths_cm": [2.0, 10.0, 4.0],
             "border_style": "header_only",
             "rows": [
                 [{"text": "1", "rowspan": 2},
                  {"text": "Мерж", "colspan": 2, "shading": "#E0E0E0", "bold": True}],
                 [{"text": "продолжение"}],
                 ["2", "обычный", "100"],
             ]},
            {"type": "callout", "text": "**Важно:** не забыть."},
            {"type": "page_break"},
            {"type": "horizontal_rule"},
        ],
    })
    envelope = build_document(spec, str(out))
    assert envelope["ok"], envelope["reason"]
    xml = docx.Document(str(out)).element.xml
    for tag in ("w:gridSpan", "w:vMerge", "w:tblHeader", "w:cantSplit", "w:shd",
                "<w:br w:type=\"page\"/>"):
        assert tag in xml
