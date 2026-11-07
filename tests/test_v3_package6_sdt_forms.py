# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 6: интерактивные формы (SDT / Content Controls).

По ТЗ TASK_PACKAGE6_SDT_FORMS.md:
  - w:sdt / w:sdtPr / w14:checkbox (val=1/0, ☒/☐) в word/document.xml;
  - w:dropDownList + w:listItem;
  - w:date (dateFormat, fullDate);
  - md: `- [x] Пункт А` / `- [ ] Пункт Б`, `{{dropdown: [...]}}`, `{{date}}`.
"""
from __future__ import annotations

import zipfile

import docx

from docx_cli.builder import build_document
from docx_cli.md_engine import md_build_v2
from docx_cli.spec import DocSpec

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build(tmp_workspace, body) -> docx.Document:
    spec = DocSpec.from_dict({
        "profile": "corporate",
        "output": {"file_name": "forms.docx", "author": "Agent"},
        "header": {"align": "right", "lines": ["Подготовлено: Agent"]},
        "title": {"main": "Тест форм (SDT)"},
        "body": body,
    })
    out = tmp_workspace / "forms.docx"
    envelope = build_document(spec, str(out))
    assert envelope["ok"], envelope["reason"]
    return out


def _document_xml(out) -> str:
    with zipfile.ZipFile(out) as z:
        return z.read("word/document.xml").decode("utf-8")


def _sdts(out) -> list:
    doc = docx.Document(str(out))
    return doc.element.body.findall(f".//{{{W}}}sdt")


# --------------------------------------------------------------- checkbox SDT

def test_checkbox_sdt_structure(tmp_workspace):
    out = _build(tmp_workspace, [
        {"type": "checkbox", "label": "Согласовано юристом", "checked": True},
        {"type": "checkbox", "label": "Ожидает оплаты", "checked": False},
    ])
    xml = _document_xml(out)
    assert "<w:sdt>" in xml and "<w:sdtPr>" in xml and "<w:sdtContent>" in xml
    assert "<w14:checkbox>" in xml
    # отмеченный чекбокс идёт первым, неотмеченный — вторым
    i1 = xml.index('<w14:checked w14:val="1"/>')
    i0 = xml.index('<w14:checked w14:val="0"/>')
    assert i1 < i0
    # состояния Word: 2612 checked / 2610 unchecked, шрифт MS Gothic
    assert 'w14:val="2612"' in xml and 'w14:val="2610"' in xml
    assert 'w14:font="MS Gothic"' in xml
    # символы Unicode в sdtContent
    assert "☒" in xml and "☐" in xml
    assert "Согласовано юристом" in xml
    # структурно: sdt -> sdtPr -> w14:checkbox, sdtContent содержит символ
    sdts = _sdts(out)
    assert len(sdts) == 2
    first = sdts[0]
    assert first.find(f"{{{W}}}sdtPr/{{{W14}}}checkbox") is not None
    content_run = first.find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t")
    assert content_run is not None and content_run.text == "\u2612"


def test_checkbox_ids_unique(tmp_workspace):
    out = _build(tmp_workspace, [
        {"type": "checkbox", "label": "Пункт 1", "checked": False},
        {"type": "checkbox", "label": "Пункт 2", "checked": True},
    ])
    ids = [s.find(f"{{{W}}}sdtPr/{{{W}}}id").get(f"{{{W}}}val") for s in _sdts(out)]
    assert len(set(ids)) == len(ids) == 2


# -------------------------------------------------------------- dropdown SDT

def test_dropdown_sdt_structure(tmp_workspace):
    out = _build(tmp_workspace, [
        {
            "type": "dropdown",
            "items": ["Российский рубль (RUB)", "Доллар США (USD)", "Евро (EUR)"],
            "selected": "Российский рубль (RUB)",
            "label": "Валюта договора: ",
        },
    ])
    xml = _document_xml(out)
    assert "<w:dropDownList>" in xml
    assert xml.count("<w:listItem ") == 3
    assert '<w:listItem w:value="Российский рубль (RUB)" w:displayText="Российский рубль (RUB)"/>' in xml
    assert '<w:listItem w:value="Доллар США (USD)" w:displayText="Доллар США (USD)"/>' in xml
    # предвыбранный элемент отображается в sdtContent
    sdt = _sdts(out)[0]
    text = sdt.find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t").text
    assert text == "Российский рубль (RUB)"


def test_dropdown_value_display_pairs(tmp_workspace):
    out = _build(tmp_workspace, [
        {
            "type": "dropdown",
            "items": [
                {"value": "RUB", "displayText": "Российский рубль (RUB)"},
                {"value": "USD", "displayText": "Доллар США (USD)"},
            ],
            "selected": "RUB",
        },
    ])
    xml = _document_xml(out)
    assert 'w:value="RUB"' in xml and 'w:value="USD"' in xml
    text = _sdts(out)[0].find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t").text
    assert text == "Российский рубль (RUB)"


def test_dropdown_no_selected_shows_first(tmp_workspace):
    out = _build(tmp_workspace, [{"type": "dropdown", "items": ["Да", "Нет"]}])
    text = _sdts(out)[0].find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t").text
    assert text == "Да"


# ------------------------------------------------------------- date и text

def test_date_picker_and_text_input(tmp_workspace):
    from docx_cli.sdt import add_date_picker, add_text_input

    doc = docx.Document()
    p = doc.add_paragraph()
    add_date_picker(p, "2026-10-28")
    add_text_input(p, "Нажмите для ввода текста")
    sdt1 = doc.element.body.findall(f".//{{{W}}}sdt")[0]
    date_el = sdt1.find(f"{{{W}}}sdtPr/{{{W}}}date")
    assert date_el is not None
    assert date_el.find(f"{{{W}}}dateFormat").get(f"{{{W}}}val") == "dd.MM.yyyy"
    assert date_el.find(f"{{{W}}}fullDate").get(f"{{{W}}}val") == "2026-10-28"
    assert sdt1.find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t").text == "28.10.2026"
    # текстовое поле: w:text в sdtPr, placeholder в sdtContent
    sdt2 = doc.element.body.findall(f".//{{{W}}}sdt")[1]
    assert sdt2.find(f"{{{W}}}sdtPr/{{{W}}}text") is not None
    assert sdt2.find(f"{{{W}}}sdtContent/{{{W}}}r/{{{W}}}t").text == "Нажмите для ввода текста"


# ---------------------------------------------------- markdown -> SDT-чекбоксы

def test_md_checkboxes_to_sdt(tmp_workspace):
    out = tmp_workspace / "md.docx"
    res = md_build_v2(
        "# Договор\n\n- [x] Пункт А\n- [ ] Пункт Б\n",
        str(out), "pytest")
    assert res["ok"], res["reason"]
    xml = _document_xml(out)
    assert xml.count("<w14:checkbox>") == 2
    assert xml.count('<w14:checked w14:val="1"/>') == 1
    assert xml.count('<w14:checked w14:val="0"/>') == 1
    assert "Пункт А" in xml and "Пункт Б" in xml
    assert "☒" in xml and "☐" in xml
    # маркеры `- [x]` не просочились в текст
    assert "[x]" not in xml and "[ ]" not in xml


def test_md_dropdown_and_date(tmp_workspace):
    out = tmp_workspace / "md2.docx"
    res = md_build_v2(
        '# Акты\n\n{{dropdown: ["Да", "Нет", "Не применимо"]}}\n\n{{date}}\n',
        str(out), "pytest")
    assert res["ok"], res["reason"]
    xml = _document_xml(out)
    assert "<w:dropDownList>" in xml
    assert xml.count("<w:listItem ") == 3
    assert '<w:listItem w:value="Не применимо" w:displayText="Не применимо"/>' in xml
    assert "<w:date>" in xml
    assert "{{" not in xml and "}}" not in xml


def test_md_dropdown_bad_items_fail_closed(tmp_workspace):
    import pytest
    from docx_cli.md_engine import MdEngineError

    with pytest.raises(MdEngineError):
        md_build_v2('# А\n\n{{dropdown: []}}\n', str(tmp_workspace / "fail.docx"), "pytest")
    with pytest.raises(MdEngineError):
        md_build_v2('# А\n\n{{dropdown: [1, 2]}}\n', str(tmp_workspace / "fail2.docx"), "pytest")
    with pytest.raises(MdEngineError):
        md_build_v2('# А\n\n{{dropdown: "Да"}}\n', str(tmp_workspace / "fail3.docx"), "pytest")
