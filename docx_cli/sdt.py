# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 6: нативные Structured Document Tags (w:sdt) — интерактивные формы.

ECMA-376 домен 9: Content Controls Word:
  - w14:checkbox   — интерактивный чекбокс (☒/☐, MS Gothic);
  - w:dropDownList — выпадающий список с w:listItem и предвыбором;
  - w:date         — селектор даты (Date Picker) с форматом YYYY-MM-DD /
                     DD.MM.YYYY;
  - w:text         — текстовое поле формы (plain text placeholder).

Все элементы создаются внутри существующего абзаца как блок w:sdt
(run-level content control). Функции add_* — низкоуровневый API,
render_* — рендереры блоков DocSpec для builder.BLOCK_RENDERERS.
"""
from __future__ import annotations

import datetime
import itertools
from xml.sax.saxutils import escape

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

from docx_cli.blocks import _body_paragraph_format, _set_run_font
from docx_cli.errors import BuildError
from docx_cli.spec import (
    CheckboxBlock,
    DatePickerBlock,
    DropdownBlock,
    Profile,
)

W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Символы чекбокса по ТЗ (checkedState/uncheckedState val="2612"/"2610").
CHECKED_SYMBOL = "\u2612"
UNCHECKED_SYMBOL = "\u2610"
CHECKBOX_FONT = "MS Gothic"

# w:id должен быть уникален в документе; простой счётчик по процессу.
_sdt_ids = itertools.count(1)


def _next_sdt_id() -> int:
    return next(_sdt_ids)


def _label_run_xml(text: str, profile: Profile | None) -> str:
    """Run sdtContent с профильным шрифтом (валидируется как body)."""
    if profile is None:
        rpr = ""
    else:
        f = escape(profile.font_name)
        sz = profile.body_font_size_pt * 2
        rpr = f'<w:rPr><w:rFonts w:ascii="{f}" w:hAnsi="{f}" w:cs="{f}"/>'
        rpr += f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    return f'<w:r>{rpr}<w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


# -------------------------------------------------------------------------------------
# Низкоуровневые конструкторы w:sdt (по сигнатурам ТЗ §2А)
# -------------------------------------------------------------------------------------

def _wrap_sdt(sdt: str, content_runs: str) -> str:
    return (
        f'<w:sdt {nsdecls("w", "w14")}>{sdt}<w:sdtContent>{content_runs}'
        f"</w:sdtContent></w:sdt>"
    )


def _append(paragraph, xml: str) -> None:
    paragraph._p.append(parse_xml(xml))


def add_checkbox(paragraph, checked: bool = False, label: str = "",
                 profile: Profile | None = None) -> None:
    """Интерактивный чекбокс: w14:checkbox + ☒/☐ (MS Gothic) и текст label."""
    sym = CHECKED_SYMBOL if checked else UNCHECKED_SYMBOL
    sdtpr = (
        f'<w:sdtPr><w:id w:val="{_next_sdt_id()}"/>'
        f"<w14:checkbox>"
        f'<w14:checked w14:val="{1 if checked else 0}"/>'
        f'<w14:checkedState w14:val="2612" w14:font="MS Gothic"/>'
        f'<w14:uncheckedState w14:val="2610" w14:font="MS Gothic"/>'
        f"</w14:checkbox></w:sdtPr>"
    )
    symbol_run = (
        '<w:r><w:rPr><w:rFonts w:ascii="MS Gothic" w:hAnsi="MS Gothic"/></w:rPr>'
        f"<w:t>{sym}</w:t></w:r>"
    )
    _append(paragraph, _wrap_sdt(sdtpr, symbol_run))
    if label:
        run = paragraph.add_run(label)
        if profile is not None:
            _set_run_font(run, profile, profile.body_font_size_pt)


def add_dropdown(paragraph, items: list, selected: str = "",
                 label: str | None = None, profile: Profile | None = None) -> None:
    """Выпадающий список: w:dropDownList + w:listItem; selected — предвыбор.

    items: list[str] (value == displayText) или list[dict] с ключами
    value/displayText. Текст sdtContent — displayText выбранного (или
    первого) элемента.
    """
    if not items:
        raise BuildError("dropdown needs at least one item")
    entries: list[tuple[str, str]] = []
    for raw in items:
        if isinstance(raw, dict):
            value = str(raw.get("value", raw.get("displayText", "")))
            display = str(raw.get("displayText", value))
        else:
            value = display = str(raw)
        entries.append((value, display))
    display = next((d for v, d in entries if v == selected), entries[0][1])
    sdtpr = (
        f'<w:sdtPr><w:id w:val="{_next_sdt_id()}"/>'
        "<w:dropDownList>"
        + "".join(
            f'<w:listItem w:value="{escape(v)}" w:displayText="{escape(d)}"/>'
            for v, d in entries
        )
        + "</w:dropDownList></w:sdtPr>"
    )
    if label:
        run = paragraph.add_run(label)
        if profile is not None:
            _set_run_font(run, profile, profile.body_font_size_pt)
    _append(paragraph, _wrap_sdt(sdtpr, _label_run_xml(display, profile)))


def _date_parts(initial_date: str) -> tuple[str, str]:
    """(дисплей dd.mm.yyyy, fullDate yyyy-mm-dd) по валидированной строке."""
    for fmt, sort in (("%Y-%m-%d", 1), ("%d.%m.%Y", 0)):
        try:
            d = datetime.datetime.strptime(initial_date, fmt)
        except ValueError:
            continue
        return d.strftime("%d.%m.%Y"), d.strftime("%Y-%m-%d")
    raise BuildError(f"unparseable date: {initial_date!r}")


def add_date_picker(paragraph, initial_date: str = "", date_format: str = "dd.MM.yyyy",
                    profile: Profile | None = None) -> None:
    """Селектор даты: w:date с w:dateFormat/w:lid и начальным значением."""
    sdtpr = (
        f'<w:sdtPr><w:id w:val="{_next_sdt_id()}"/>'
        "<w:date>"
        f'<w:dateFormat w:val="{escape(date_format)}"/>'
        '<w:lid w:val="ru-RU"/>'
    )
    display = ""
    if initial_date:
        display, full = _date_parts(initial_date)
        sdtpr += f'<w:fullDate w:val="{full}"/>'
    sdtpr += "</w:date></w:sdtPr>"
    _append(paragraph, _wrap_sdt(sdtpr, _label_run_xml(display, profile)))


def add_text_input(paragraph, placeholder: str = "Введите текст...",
                   profile: Profile | None = None) -> None:
    """Текстовое поле формы: w:text в sdtPr + placeholder-текст в sdtContent."""
    sdtpr = f'<w:sdtPr><w:id w:val="{_next_sdt_id()}"/><w:text/></w:sdtPr>'
    _append(paragraph, _wrap_sdt(sdtpr, _label_run_xml(placeholder, profile)))


# -------------------------------------------------------------------------------------
# Рендереры блоков DocSpec (регистрируются в builder.BLOCK_RENDERERS)
# -------------------------------------------------------------------------------------

def _form_paragraph(doc: Document, profile: Profile):
    p = doc.add_paragraph()
    _body_paragraph_format(p, profile)
    return p


def render_checkbox(doc: Document, block: CheckboxBlock, profile: Profile) -> None:
    add_checkbox(_form_paragraph(doc, profile), checked=block.checked,
                 label=block.label, profile=profile)


def render_dropdown(doc: Document, block: DropdownBlock, profile: Profile) -> None:
    from docx_cli.spec import DropdownItem

    items = [
        {"value": it.value, "displayText": it.display}
        if isinstance(it, DropdownItem) else it
        for it in block.items
    ]
    add_dropdown(_form_paragraph(doc, profile), items=items,
                 selected=block.selected or "", label=block.label, profile=profile)


def render_date_picker(doc: Document, block: DatePickerBlock, profile: Profile) -> None:
    add_date_picker(_form_paragraph(doc, profile), initial_date=block.initial_date,
                    date_format=block.date_format, profile=profile)


__all__ = [
    "add_checkbox", "add_dropdown", "add_date_picker", "add_text_input",
    "render_checkbox", "render_dropdown", "render_date_picker",
]
