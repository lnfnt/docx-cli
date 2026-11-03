# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Нативные концевые сноски OpenXML: word/endnotes.xml + w:endnoteReference.

Аналог docx_cli.footnotes, но для концевых сносок (выносятся в конец документа):
  - ensure_endnotes_part: создаёт word/endnotes.xml с обязательными
    служебными сносками separator (w:id="-1") и continuationSeparator
    (w:id="0"), связь типа .../relationships/endnotes в
    word/_rels/document.xml.rels и Content-Type
    application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml
    в [Content_Types].xml;
  - add_endnote: новая <w:endnote w:id="{n}"> с абзацем и текстом сноски,
    а в целевом параграфе run (необязательный run_text) и superscript-run
    со стилем EndnoteReference и <w:endnoteReference w:id="{n}"/>.
"""
from __future__ import annotations

import re
from xml.sax.saxutils import escape as _xml_escape

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

NS = nsdecls("w")

_ENDNOTES_PARTNAME = PackURI("/word/endnotes.xml")

# Служебный каркас части: две обязательные концевые сноски-разделителя.
_ENDNOTES_SKELETON = (
    f'<w:endnotes {NS}>'
    '<w:endnote w:type="separator" w:id="-1">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:separator/></w:r></w:p></w:endnote>"
    '<w:endnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:continuationSeparator/></w:r></w:p></w:endnote>"
    "</w:endnotes>"
)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DEFAULT_FONT = "Times New Roman"

# Синтаксис маркера концевой сноски: [^!id] или [^e:id] или [^e1]
ENDNOTE_MARKER_RE = re.compile(r"\[\^(?:!|e:)(\w+)\]")


def ensure_endnotes_part(doc: Document) -> XmlPart:
    """Часть word/endnotes.xml: существующая или созданная вместе с rels/content-type."""
    try:
        return doc.part.part_related_by(RT.ENDNOTES)
    except KeyError:
        pass
    element = parse_xml(_ENDNOTES_SKELETON)
    endnotes_part = XmlPart(_ENDNOTES_PARTNAME, CT.WML_ENDNOTES, element, doc.part.package)
    doc.part.relate_to(endnotes_part, RT.ENDNOTES)
    return endnotes_part


def next_endnote_id(endnotes_part: XmlPart) -> int:
    """Свободный числовой id (служебные -1/0 заняты; нумерация ссылок — 1, 2, ...)."""
    ids = [int(f.get(qn("w:id"))) for f in endnotes_part.element.findall(qn("w:endnote"))]
    return max([0, *ids]) + 1


def split_endnote_markers(text: str) -> list[tuple[str, str | None]]:
    """Текст -> [(chunk, None | endnote_id)]: маркеры [^!id] или [^e:id] — отдельные сегменты."""
    segments: list[tuple[str, str | None]] = []
    pos = 0
    for m in ENDNOTE_MARKER_RE.finditer(text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], None))
        segments.append(("", m.group(1)))
        pos = m.end()
    if pos < len(text) or not segments:
        segments.append((text[pos:], None))
    return segments


def add_endnote(doc: Document, paragraph, run_text: str, endnote_text: str) -> int:
    """Нативная концевая сноска: body в word/endnotes.xml + ссылка в параграфе.

    В параграфе создается superscript run со стилем EndnoteReference и
    <w:endnoteReference w:id="{n}"/>. Возвращает id сноски.
    """
    font = doc.styles["Normal"].font.name or _DEFAULT_FONT
    endnotes_part = ensure_endnotes_part(doc)
    eid = next_endnote_id(endnotes_part)

    endnotes_part.element.append(parse_xml(
        f'<w:endnote {NS} w:id="{eid}">'
        '<w:p><w:pPr><w:pStyle w:val="EndnoteText"/>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="{_xml_escape(font)}" w:hAnsi="{_xml_escape(font)}"'
        f' w:cs="{_xml_escape(font)}"/></w:rPr><w:t xml:space="preserve">'
        f'{_xml_escape(endnote_text)}</w:t></w:r></w:p></w:endnote>'
    ))

    if run_text:
        paragraph.add_run(run_text)
    ref_run = paragraph.add_run()
    ref_run._element.append(parse_xml(
        f'<w:rPr xmlns:w="{_W_NS}">'
        f'<w:rStyle w:val="EndnoteReference"/>'
        f'<w:vertAlign w:val="superscript"/>'
        f'</w:rPr>'
    ))
    ref_run._element.append(parse_xml(
        f'<w:endnoteReference xmlns:w="{_W_NS}" w:id="{eid}"/>'
    ))
    return eid
