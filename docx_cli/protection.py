# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 9: Защита документа от редактирования (Document Protection — Домен 11).

Режимы:
- readOnly: только для чтения
- comments: разрешено оставлять только комментарии
- trackedChanges: любые изменения пишутся через Track Changes
- forms: разрешено только заполнение полей форм (SDT)
"""

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os

from docx_cli.errors import BuildError

# S3-L1: одна каноническая таблица значений + явный слой алиасов —
# иначе две параллельные таблицы расходятся при добавлении режимов.
_PROTECTION_CANONICAL = {
    "readOnly": "readOnly",
    "comments": "comments",
    "trackedChanges": "trackedChanges",
    "forms": "forms",
}
_PROTECTION_ALIASES = {
    "read_only": "readOnly",
    "readonly": "readOnly",
    "comments_only": "comments",
    "tracked_changes": "trackedChanges",
    "forms_only": "forms",
}
PROTECTION_MODES = {
    **_PROTECTION_CANONICAL,
    **_PROTECTION_ALIASES,
}

# Порядок следования дочерних элементов в w:settings по схеме ECMA-376:
_SETTINGS_CHILD_ORDER = [
    "w:writeProtection",
    "w:view",
    "w:zoom",
    "w:removePersonalInformation",
    "w:removeDateAndTime",
    "w:doNotDisplayPageBoundaries",
    "w:displayBackgroundShape",
    "w:printPostScriptOverText",
    "w:printFractionalCharacterWidth",
    "w:printFormsData",
    "w:embedTrueTypeFonts",
    "w:embedSystemFonts",
    "w:saveSubsetFonts",
    "w:saveFormsData",
    "w:mirrorMargins",
    "w:alignBordersAndEdges",
    "w:bordersDoNotSurroundHeader",
    "w:bordersDoNotSurroundFooter",
    "w:gutterAtTop",
    "w:hideSpellingErrors",
    "w:hideGrammaticalErrors",
    "w:activeWritingStyle",
    "w:proofState",
    "w:formsDesign",
    "w:attachedTemplate",
    "w:linkStyles",
    "w:stylePaneFormatFilter",
    "w:stylePaneSortMethod",
    "w:documentType",
    "w:mailMerge",
    "w:revisionView",
    "w:trackRevisions",
    "w:doNotTrackMoves",
    "w:doNotTrackFormatting",
    "w:documentProtection",
    "w:autoFormatOverride",
    "w:styleLockTheme",
    "w:styleLockQFSet",
    "w:defaultTabStop",
    "w:autoHyphenation",
    "w:consecutiveHyphenLimit",
    "w:hyphenationZone",
    "w:doNotHyphenateCaps",
    "w:showEnvelope",
    "w:summaryLength",
    "w:clickAndTypeStyle",
    "w:defaultTableStyle",
    "w:evenAndOddHeaders",
    "w:bookFoldRevPrinting",
    "w:bookFoldPrinting",
    "w:bookFoldPrintingSheets",
    "w:drawingGridHorizontalSpacing",
    "w:drawingGridVerticalSpacing",
    "w:displayHorizontalDrawingGridEvery",
    "w:displayVerticalDrawingGridEvery",
    "w:doNotUseMarginsForDrawingGridOrigin",
    "w:drawingGridHorizontalOrigin",
    "w:drawingGridVerticalOrigin",
    "w:doNotShadeFormData",
    "w:noPunctuationKerning",
    "w:characterSpacingControl",
    "w:printTwoOnOne",
    "w:strictFirstAndLastChars",
    "w:noLineBreaksAfter",
    "w:noLineBreaksBefore",
    "w:saveXmlDataOnly",
    "w:saveThroughXslt",
    "w:showXMLTags",
    "w:alwaysMergeEmptyNamespace",
    "w:updateFields",
    "w:footnotePr",
    "w:endnotePr",
    "w:compat",
    "w:docVars",
]


def set_document_protection(doc, edit_mode: str = "readOnly",
                            enforcement: bool = True,
                            password: str | None = None) -> None:
    """Устанавливает защиту документа от изменений в word/settings.xml.

    S2-H3: enforcement без пароля запрещён — такой тег Word трактует как
    защиту с пустым паролем (снимается в один клик, защита-театр).
    С паролем пишем salted SHA-1 атрибуты (MS-OI29500): hashValue/salt/
    algorithmName — без них enforcement бессмысленен.
    """
    val = PROTECTION_MODES.get(edit_mode)
    if not val:
        valid = ", ".join(sorted(PROTECTION_MODES.keys()))
        raise BuildError(f"Unknown protection mode '{edit_mode}'. Valid: {valid}")
    if enforcement and not password:
        raise BuildError(
            "document protection without password is forbidden: "
            "enforcement without hash is equivalent to empty password")

    settings = doc.settings.element
    w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

    # Удаляем старый тег защиты, если был
    for old in settings.findall(f"{{{w_ns}}}documentProtection"):
        settings.remove(old)

    enf_str = "1" if enforcement else "0"
    extra = ""
    if password:
        import base64
        import hashlib
        salt = os.urandom(16)
        digest = hashlib.sha1(salt + password.encode("utf-16le")).digest()
        extra = (
            f' w:hashValue="{base64.b64encode(digest).decode()}"'
            f' w:salt="{base64.b64encode(salt).decode()}"'
            f' w:algorithmName="SHA-1"'
        )
    tag_xml = (
        f'<w:documentProtection {nsdecls("w")} '
        f'w:edit="{val}" w:enforcement="{enf_str}"{extra}/>'
    )
    prot_elem = parse_xml(tag_xml)

    # Вставляем по порядку схемы
    prot_idx = _SETTINGS_CHILD_ORDER.index("w:documentProtection")
    inserted = False
    for child in list(settings):
        tag_name = child.tag
        if tag_name.startswith(f"{{{w_ns}}}"):
            q_name = f"w:{tag_name[len(w_ns) + 2:]}"
            if q_name in _SETTINGS_CHILD_ORDER:
                if _SETTINGS_CHILD_ORDER.index(q_name) > prot_idx:
                    child.addprevious(prot_elem)
                    inserted = True
                    break

    if not inserted:
        settings.append(prot_elem)
