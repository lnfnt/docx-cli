# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Нативные комментарии OpenXML: word/comments.xml + w:commentRangeStart/End
+ w:commentReference.

python-docx не даёт API для комментариев (аналогично сноскам), поэтому здесь
низкоуровневая работа с OPC-упаковкой:
  - ensure_comments_part: создаёт word/comments.xml с корнем w:comments,
    связь типа .../relationships/comments в word/_rels/document.xml.rels и
    Content-Type ...comments+xml в [Content_Types].xml;
  - add_comment: находит run'ы с target_text (при необходимости режет
    пограничные run'ы), оборачивает их в w:commentRangeStart/End, ставит run
    с w:commentReference и пишет тело комментария в word/comments.xml;
  - add_comment_to_runs: обёртка уже готовых run'ов (рендерер md знает
    границы сам — поиск по тексту не нужен).
"""
from __future__ import annotations

import copy
import re
from datetime import datetime, timezone
from xml.sax.saxutils import escape as _xml_escape

from docx import Document
from docx.opc.constants import CONTENT_TYPE as CT
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.packuri import PackURI
from docx.opc.part import XmlPart
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn

NS = nsdecls("w")
_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_COMMENTS_PARTNAME = PackURI("/word/comments.xml")


class CommentError(Exception):
    """Fail-closed ошибка комментария (цель не найдена в параграфе)."""


def ensure_comments_part(doc: Document) -> XmlPart:
    """Часть word/comments.xml: существующая или созданная вместе с rels/content-type."""
    try:
        return doc.part.part_related_by(RT.COMMENTS)
    except KeyError:
        pass
    element = parse_xml(f"<w:comments {NS}/>")
    comments_part = XmlPart(_COMMENTS_PARTNAME, CT.WML_COMMENTS, element, doc.part.package)
    doc.part.relate_to(comments_part, RT.COMMENTS)
    return comments_part


def next_comment_id(comments_part: XmlPart) -> int:
    """Свободный числовой id (1, 2, ...)."""
    ids = [int(c.get(qn("w:id"))) for c in comments_part.element.findall(qn("w:comment"))]
    return max([0, *ids]) + 1


def _stamp(date: str | None) -> str:
    # S2-L: микросекунды в w:date — недетерминизм пакета; отрезаем.
    return date or datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def add_comment_body(comments_part: XmlPart, comment_text: str,
                     author: str = "colleague", date: str | None = None) -> int:
    """Тело комментария <w:comment w:id w:author w:date><w:p><w:r><w:t/></w:r></w:p>. Возвращает id."""
    cid = next_comment_id(comments_part)
    comments_part.element.append(parse_xml(
        f'<w:comment {NS} w:id="{cid}" w:author="{_xml_escape(author)}"'
        f' w:date="{_xml_escape(_stamp(date))}">'
        f'<w:p><w:r><w:t xml:space="preserve">{_xml_escape(comment_text)}</w:t></w:r></w:p>'
        f"</w:comment>"
    ))
    return cid


def add_comment_to_runs(doc: Document, paragraph, runs: list,
                        comment_text: str, author: str = "colleague",
                        date: str | None = None) -> int:
    """Комментарий на готовые run'ы: RangeStart перед первым, RangeEnd после
    последнего, run с w:commentReference после RangeEnd, тело в comments.xml."""
    if not runs:
        raise CommentError("comment requires at least one run (fail-closed)")
    if not comment_text.strip():
        raise CommentError("empty comment text (fail-closed)")
    comments_part = ensure_comments_part(doc)
    cid = add_comment_body(comments_part, comment_text, author, date)

    first = runs[0]._element if hasattr(runs[0], "_element") else runs[0]
    last = runs[-1]._element if hasattr(runs[-1], "_element") else runs[-1]
    first.addprevious(parse_xml(f'<w:commentRangeStart {NS} w:id="{cid}"/>'))
    end = parse_xml(f'<w:commentRangeEnd {NS} w:id="{cid}"/>')
    last.addnext(end)
    end.addnext(parse_xml(
        f'<w:r {NS}><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr>'
        f'<w:commentReference w:id="{cid}"/></w:r>'
    ))
    return cid


def _set_run_text(r_element, text: str) -> None:
    for t in r_element.findall(qn("w:t")):
        r_element.remove(t)
    r_element.append(parse_xml(
        f'<w:t xmlns:w="{_W_NS}" xml:space="preserve">{_xml_escape(text)}</w:t>'
    ))


def _split_run(r_element, offset: int):
    """Разрезать run на два по символьному смещению: r_element := левая часть,
    возвращается клон с правой частью (форматирование копируется).

    ponytail: предполагает текст в w:t (без w:tab/w:br внутри run'а) — в
    собранных нами документах это всегда так; для чужих файлов с дробными
    текстовыми узлами текст склеивается и перезаписывается одним w:t.
    """
    text = "".join(t.text or "" for t in r_element.findall(qn("w:t")))
    # S2-M1: run с не-текстовыми узлами (табуляции/разрывы) не режем —
    # иначе склейка w:t молча уничтожает w:tab/w:br чужого файла.
    if r_element.findall(qn("w:tab")) or r_element.findall(qn("w:br")):
        raise CommentError(
            "cannot split run containing w:tab/w:br (formatting preservation)")
    clone = copy.deepcopy(r_element)
    r_element.addnext(clone)
    _set_run_text(r_element, text[:offset])
    _set_run_text(clone, text[offset:])
    return clone


def add_comment(doc: Document, paragraph, target_text: str,
                comment_text: str, author: str = "colleague",
                date: str | None = None, occurrence: int = 1) -> int:
    """Комментарий на фрагмент target_text внутри параграфа.

    Находит run'ы, покрывающие target_text (пограничные режет по символам),
    оборачивает их в w:commentRangeStart/End и добавляет w:commentReference.
    """
    if not target_text:
        raise CommentError("empty comment target (fail-closed)")
    joined = "".join(r.text for r in paragraph.runs)
    # S2-H6: occurrence — какое по счёту вхождение внутри параграфа якорить
    # (1-based). Молчаливое «первое совпадение» при повторном target —
    # источник ошибок адресации замечаний; выход за диапазон — fail-closed.
    starts = [m.start() for m in re.finditer(re.escape(target_text), joined)]
    occurrence = max(1, int(occurrence))
    if occurrence > len(starts):
        raise CommentError(
            f"comment target {target_text!r} occurrence {occurrence} "
            f"out of range ({len(starts)} found)")
    start = starts[occurrence - 1]
    stop = start + len(target_text)

    covered: list = []
    pos = 0
    for run in list(paragraph.runs):
        rs, re_ = pos, pos + len(run.text)
        pos = re_
        if re_ <= start or rs >= stop:
            continue
        el = run._element
        cur_rs = rs
        if rs < start:  # левое плечо: el остаётся до цели, клон — внутри
            el = _split_run(el, start - rs)
            cur_rs = start
        if re_ > stop:  # правое плечо: el (левая часть) внутри, клон — после
            _split_run(el, stop - cur_rs)
        covered.append(el)
    if not covered:
        raise CommentError(f"comment target {target_text!r} covers no runs")

    return add_comment_to_runs(doc, paragraph, covered, comment_text, author, date)
