# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Нативные сноски OpenXML: word/footnotes.xml + w:footnoteReference.

python-docx не даёт высокоуровневого API для сносок, поэтому здесь —
низкоуровневая работа с OPC-упаковкой:
  - ensure_footnotes_part: создаёт word/footnotes.xml с обязательными
    служебными сносками separator (w:id="-1") и continuationSeparator
    (w:id="0"), связь типа .../relationships/footnotes в
    word/_rels/document.xml.rels и Content-Type
    application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml
    в [Content_Types].xml;
  - add_footnote: новая <w:footnote w:id="{n}"> с абзацем и текстом сноски,
    а в целевом параграфе run (необязательный run_text) и superscript-run
    с <w:footnoteReference w:id="{n}"/>.
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

_FOOTNOTES_PARTNAME = PackURI("/word/footnotes.xml")

# Служебный каркас части: две обязательные сноски-разделителя.
_FOOTNOTES_SKELETON = (
    f'<w:footnotes {NS}>'
    '<w:footnote w:type="separator" w:id="-1">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:separator/></w:r></w:p></w:footnote>"
    '<w:footnote w:type="continuationSeparator" w:id="0">'
    '<w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
    "<w:r><w:continuationSeparator/></w:r></w:p></w:footnote>"
    "</w:footnotes>"
)

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_DEFAULT_FONT = "Times New Roman"

FOOTNOTE_MARKER_RE = re.compile(r"\[\^(\w+)\]")


def ensure_footnotes_part(doc: Document) -> XmlPart:
    """Часть word/footnotes.xml: существующая или созданная вместе с rels/content-type."""
    try:
        return doc.part.part_related_by(RT.FOOTNOTES)
    except KeyError:
        pass
    element = parse_xml(_FOOTNOTES_SKELETON)
    footnotes_part = XmlPart(_FOOTNOTES_PARTNAME, CT.WML_FOOTNOTES, element, doc.part.package)
    doc.part.relate_to(footnotes_part, RT.FOOTNOTES)
    return footnotes_part


def next_footnote_id(footnotes_part: XmlPart) -> int:
    """Свободный числовой id (служебные -1/0 заняты; нумерация ссылок — 1, 2, ...)."""
    ids = [int(f.get(qn("w:id"))) for f in footnotes_part.element.findall(qn("w:footnote"))]
    return max([0, *ids]) + 1


def add_footnote(doc: Document, paragraph, run_text: str, footnote_text: str) -> int:
    """Нативная сноска: body в word/footnotes.xml + ссылка в параграфе.

    Если run_text не пуст — сначала run с этим текстом, затем superscript-run
    с <w:footnoteReference w:id="{n}"/>. Возвращает id сноски.
    """
    font = doc.styles["Normal"].font.name or _DEFAULT_FONT
    footnotes_part = ensure_footnotes_part(doc)
    fid = next_footnote_id(footnotes_part)

    footnotes_part.element.append(parse_xml(
        f'<w:footnote {NS} w:id="{fid}">'
        '<w:p><w:pPr><w:pStyle w:val="FootnoteText"/>'
        '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
        f'<w:r><w:rPr><w:rFonts w:ascii="{_xml_escape(font)}" w:hAnsi="{_xml_escape(font)}"'
        f' w:cs="{_xml_escape(font)}"/></w:rPr><w:t xml:space="preserve">'
        f'{_xml_escape(footnote_text)}</w:t></w:r></w:p></w:footnote>'
    ))

    if run_text:
        paragraph.add_run(run_text)
    ref_run = paragraph.add_run()
    ref_run._element.append(parse_xml(
        f'<w:rPr xmlns:w="{_W_NS}"><w:vertAlign w:val="superscript"/></w:rPr>'
    ))
    ref_run._element.append(parse_xml(
        f'<w:footnoteReference xmlns:w="{_W_NS}" w:id="{fid}"/>'
    ))
    return fid


def split_footnote_markers(text: str) -> list[tuple[str, str | None]]:
    """Текст -> [(chunk, None | footnote_id)]: маркеры [^id] — отдельные сегменты."""
    segments: list[tuple[str, str | None]] = []
    pos = 0
    for m in FOOTNOTE_MARKER_RE.finditer(text):
        if m.start() > pos:
            segments.append((text[pos:m.start()], None))
        segments.append(("", m.group(1)))
        pos = m.end()
    if pos < len(text) or not segments:
        segments.append((text[pos:], None))
    return segments
