# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

r"""Нативные закладки (Bookmarks) и перекрестные ссылки (Cross-References).

ECMA-376 / ISO/IEC 29500 Домен 8:
- add_bookmark: оборачивает текст в пару:
    <w:bookmarkStart w:id="{id}" w:name="{name}"/>
    ... (run с текстом, если передан)
    <w:bookmarkEnd w:id="{id}"/>
  Либо ставит пустой якорь:
    <w:bookmarkStart w:id="{id}" w:name="{name}"/>
    <w:bookmarkEnd w:id="{id}"/>

- add_cross_reference: создает поле OpenXML:
    <w:fldSimple w:instr="REF {bookmark_name} \h">
      <w:r><w:rPr>...</w:rPr><w:t>{display_text}</w:t></w:r>
    </w:fldSimple>
  или при page_ref=True:
    <w:fldSimple w:instr="PAGEREF {bookmark_name} \h">
      <w:r><w:rPr>...</w:rPr><w:t>{display_text or '1'}</w:t></w:r>
    </w:fldSimple>
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape as _xml_escape
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

NS = nsdecls("w")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Regex для поиска якорей вида {#anchor-name} или <a id="anchor-name"></a>
ANCHOR_RE = re.compile(r"(?:\{#([A-Za-z0-9_\-\.]+)\}|<a\s+(?:id|name)=[\"']([A-Za-z0-9_\-\.]+)[\"']\s*>\s*</a>)")
# Regex для перекрестных ссылок:
# 1) [display text](#anchor-name) -> REF anchor-name \h
# 2) [display text](@page:#anchor-name) -> PAGEREF anchor-name \h
# 3) [@page:#anchor-name] или [@page:anchor-name] -> PAGEREF anchor-name \h
CROSS_REF_RE = re.compile(
    r"(?:\[([^\]]+)\]\((@page:)?#([A-Za-z0-9_\-\.]+)\)|\[@page:#?([A-Za-z0-9_\-\.]+)\])"
)


def _next_bookmark_id(paragraph) -> int:
    """Генерирует свободный целочисленный id для закладки в документе."""
    doc_root = paragraph._p.getroottree().getroot()
    existing_ids = []
    for bm in doc_root.iter(qn("w:bookmarkStart")):
        val = bm.get(qn("w:id"))
        if val is not None and val.isdigit():
            existing_ids.append(int(val))
    return max([0, *existing_ids]) + 1


def add_bookmark(paragraph, name: str, text: str = "") -> int:
    """Добавляет закладку в paragraph.

    Если text задан — создает run с этим текстом между bookmarkStart и bookmarkEnd.
    Если text пустой — создает пустой якорь bookmarkStart/End.
    Возвращает числовой id закладки.
    """
    bm_id = _next_bookmark_id(paragraph)
    bm_start = parse_xml(f'<w:bookmarkStart {NS} w:id="{bm_id}" w:name="{_xml_escape(name)}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {NS} w:id="{bm_id}"/>')

    paragraph._p.append(bm_start)
    if text:
        run = paragraph.add_run(text)
    paragraph._p.append(bm_end)
    return bm_id


def add_cross_reference(paragraph, bookmark_name: str, display_text: str = "",
                        page_ref: bool = False, profile = None, size_pt: int | None = None) -> None:
    """Добавляет перекрестную ссылку в paragraph через w:fldSimple.

    page_ref=False -> REF {bookmark_name} \\h
    page_ref=True  -> PAGEREF {bookmark_name} \\h
    """
    instr_type = "PAGEREF" if page_ref else "REF"
    instr = f"{instr_type} {_xml_escape(bookmark_name)} \\h"

    if not display_text:
        display_text = "1" if page_ref else bookmark_name

    rpr_xml = ""
    if profile is not None:
        font = profile.font_name
        sz = str((size_pt or profile.body_font_size_pt) * 2)
        rpr_xml = (
            f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" '
            f'w:eastAsia="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
        )

    fld_xml = (
        f'<w:fldSimple {NS} w:instr="{instr}">'
        f'<w:r>{rpr_xml}<w:t>{_xml_escape(display_text)}</w:t></w:r>'
        f'</w:fldSimple>'
    )
    paragraph._p.append(parse_xml(fld_xml))
