# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Track Changes OpenXML: <w:ins>, <w:del>/<w:delText> и флаг
<w:trackRevisions/> в word/settings.xml.

add_revision: добавленный текст оборачивается в <w:ins> (run внутри),
удалённый — в <w:del> с <w:delText> (по ECMA-376 удалённый текст обязан
жить в delText, а не w:t). Латентно-неизменённый текст остаётся обычным
run'ом документа.

ensure_track_revisions: постоянный режим записи правок в Word.
"""
from __future__ import annotations

import itertools
from datetime import datetime, timezone
from xml.sax.saxutils import escape as _xml_escape

from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import qn

_W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Уникальные w:id правок: глобальный счётчик, ids не пересекаются в документе.
_REVISION_IDS = itertools.count(1)

# Позиция w:trackRevisions в CT_Settings: элементы схемы ПОСЛЕ него.
# w:documentProtection..w:defaultTabStop идут следом за trackRevisions,
# далее всё то же, что после w:updateFields.
_SETTINGS_AFTER_TRACK_REVISIONS = (
    "w:documentProtection", "w:autoFormatOverride", "w:styleLockTheme",
    "w:styleLockQFSet", "w:defaultTabStop", "w:autoHyphenation",
    "w:consecutiveHyphenLimit", "w:hyphenationZone", "w:doNotHyphenateCaps",
    "w:showEnvelope", "w:summaryLength", "w:clickAndTypeStyle",
    "w:defaultTableStyle", "w:evenAndOddHeaders", "w:bookFoldRevPrinting",
    "w:bookFoldPrinting", "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing", "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery", "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin", "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin", "w:doNotShadeFormData", "w:noPunctuationKerning",
    "w:characterSpacingControl", "w:printTwoOnOne", "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter", "w:noLineBreaksBefore", "w:savePreviewPicture",
    "w:doNotValidateAgainstSchema", "w:saveInvalidXml", "w:ignoreMixedContent",
    "w:alwaysShowPlaceholderText", "w:doNotDemarcateInvalidXml",
    "w:saveXmlDataOnly", "w:useXSLTWhenSaving", "w:saveThroughXslt",
    "w:showXMLTags", "w:alwaysMergeEmptyNamespace", "w:updateFields",
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)


def _stamp(date: str | None) -> str:
    return date or datetime.now(timezone.utc).isoformat()


def _run_props(font: str | None, size_pt: int | None) -> str:
    """run-properties для run'а внутри w:ins/w:del (шрифт тела документа)."""
    if not font:
        return ""
    sz = str(int((size_pt or 12) * 2))
    return (
        f'<w:rPr><w:rFonts w:ascii="{_xml_escape(font)}" w:hAnsi="{_xml_escape(font)}"'
        f' w:cs="{_xml_escape(font)}" w:eastAsia="{_xml_escape(font)}"/>'
        f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    )


def add_revision(paragraph, inserted_text: str | None = None,
                 deleted_text: str | None = None, author: str = "colleague",
                 date: str | None = None, font: str | None = None,
                 size_pt: int | None = None) -> None:
    """Правка Track Changes в конце параграфа: w:ins (добавлено) и/или
    w:del с w:delText (удалено) с атрибутами w:id/w:author/w:date."""
    if inserted_text is None and deleted_text is None:
        raise ValueError("add_revision requires inserted_text or deleted_text")
    stamp = _stamp(date)
    if inserted_text is not None:
        rid = next(_REVISION_IDS)
        paragraph._p.append(parse_xml(
            f'<w:ins xmlns:w="{_W_NS}" w:id="{rid}" w:author="{_xml_escape(author)}"'
            f' w:date="{_xml_escape(stamp)}">'
            f'<w:r>{_run_props(font, size_pt)}'
            f'<w:t xml:space="preserve">{_xml_escape(inserted_text)}</w:t></w:r></w:ins>'
        ))
    if deleted_text is not None:
        rid = next(_REVISION_IDS)
        paragraph._p.append(parse_xml(
            f'<w:del xmlns:w="{_W_NS}" w:id="{rid}" w:author="{_xml_escape(author)}"'
            f' w:date="{_xml_escape(stamp)}">'
            f'<w:r>{_run_props(font, size_pt)}'
            f'<w:delText xml:space="preserve">{_xml_escape(deleted_text)}</w:delText></w:r></w:del>'
        ))


def ensure_track_revisions(doc: Document) -> None:
    """<w:trackRevisions/> в word/settings.xml — постоянный режим записи правок."""
    settings = doc.settings.element
    if settings.find(qn("w:trackRevisions")) is not None:
        return
    settings.insert_element_before(
        parse_xml(f'<w:trackRevisions xmlns:w="{_W_NS}"/>'),
        *_SETTINGS_AFTER_TRACK_REVISIONS,
    )
