# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 3: настоящие многоуровневые списки OpenXML (word/numbering.xml).

Word нумерует списки не текстом "1.", а связкой
<w:numPr><w:ilvl w:val="N"/><w:numId w:val="M"/></w:numPr> -> w:num_NUM ->
w:abstractNum_ID -> w:lvl (lvlText "%1.", "%1.%2.", "%1.%2.%3.").
Модуль добавляет в numbering_part документа (шаблон python-docx всегда
содержит word/numbering.xml) по одному w:abstractNum с тремя уровнями и
w:num-экземпляру на каждый блок списка; ID выделяются детерминированно
как max+1 поверх существующих.
"""
from __future__ import annotations

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

NS = nsdecls("w")

# Уровень -> lvlText. decimal: сквозная иерархия 1. / 1.1. / 1.1.1.
_LEVEL_TEXTS = {
    "decimal": ("%1.", "%1.%2.", "%1.%2.%3."),
    "bullet": ("•", "•", "•"),
}
# Отступы уровней: 0.63 / 1.27 / 1.90 см (twips).
_LEVEL_INDENT = (357, 720, 1080)


def _next_id(element, tag: str, attr: str) -> int:
    vals = [int(e.get(qn(attr))) for e in element.findall(qn(tag))]
    return max(vals, default=0) + 1


def _levels_xml(fmt: str, start: int) -> str:
    parts = []
    for ilvl, (text, indent) in enumerate(zip(_LEVEL_TEXTS[fmt], _LEVEL_INDENT)):
        lvl_start = start if ilvl == 0 else 1
        parts.append(
            f'<w:lvl {NS} w:ilvl="{ilvl}">'
            f'<w:start w:val="{lvl_start}"/>'
            f'<w:numFmt w:val="{fmt}"/>'
            f'<w:lvlText w:val="{text}"/>'
            f'<w:lvlJc w:val="left"/>'
            f'<w:pPr><w:ind w:left="{indent}" w:hanging="{indent}"/></w:pPr>'
            f'</w:lvl>'
        )
    return "".join(parts)


def add_multilevel_list(doc: Document, kind: str = "decimal", start: int = 1) -> int:
    """Добавить w:abstractNum (уровни 0-2) + w:num в numbering part; -> numId."""
    numbering = doc.part.numbering_part.element
    fmt = "decimal" if kind == "decimal" else "bullet"
    abstract_id = _next_id(numbering, "w:abstractNum", "w:abstractNumId")
    abstract = parse_xml(
        f'<w:abstractNum {NS} w:abstractNumId="{abstract_id}">'
        f'<w:multiLevelType w:val="multilevel"/>'
        f'{_levels_xml(fmt, start)}'
        f'</w:abstractNum>'
    )
    first_num = numbering.find(qn("w:num"))
    if first_num is not None:
        first_num.addprevious(abstract)  # по CT_Numbering abstractNum перед num
    else:
        numbering.append(abstract)
    num_id = _next_id(numbering, "w:num", "w:numId")
    numbering.append(parse_xml(
        f'<w:num {NS} w:numId="{num_id}">'
        f'<w:abstractNumId w:val="{abstract_id}"/>'
        f'</w:num>'
    ))
    return num_id
