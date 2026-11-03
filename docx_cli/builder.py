# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Core document builder."""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from docx import Document
from docx.shared import Cm

from docx_cli import blocks, omml, sdt, validate
from docx_cli.comments import add_comment
from docx_cli.errors import BuildError, ValidationError
from docx_cli.spec import (
    AttachmentsListBlock,
    DocSpec,
    NumberedListBlock,
    ParagraphBlock,
    Profile,
    SignatureBlock,
    TableBlock,
    now_iso,
)


BLOCK_RENDERERS = {
    "paragraph": blocks.render_paragraph,
    "numbered_list": blocks.render_numbered_list,
    "bullet_list": blocks.render_bullet_list,
    "heading": blocks.render_heading,
    "table": blocks.render_table,
    "signature_block": blocks.render_signature_block,
    "attachments_list": blocks.render_attachments_list,
    "pleading_marker": blocks.render_pleading_marker,
    "callout": blocks.render_callout,
    "page_break": blocks.render_page_break,
    "horizontal_rule": blocks.render_horizontal_rule,
    "toc": blocks.render_toc,
    "section_break": blocks.render_section_break,
    "image": blocks.render_image,
    # ПАКЕТ 6: интерактивные формы (w:sdt content controls).
    "checkbox": sdt.render_checkbox,
    "dropdown": sdt.render_dropdown,
    "date_picker": sdt.render_date_picker,
    # ПАКЕТ 7: математические выражения (Office Math / OMML).
    "math": omml.render_math,
}


def _configure_section(doc: Document, profile: Profile) -> None:
    section = doc.sections[0]
    section.page_width = Cm(profile.paper_width_cm)
    section.page_height = Cm(profile.paper_height_cm)
    section.left_margin = Cm(profile.margin_left_cm)
    section.right_margin = Cm(profile.margin_right_cm)
    section.top_margin = Cm(profile.margin_top_cm)
    section.bottom_margin = Cm(profile.margin_bottom_cm)


def build_document(spec: DocSpec, output_path: str) -> dict[str, Any]:
    """Generate a .docx and run built-in validation.

    Returns the standard envelope. On validation error or any unexpected
    failure the partial output file is deleted and an error envelope is
    returned.
    """
    profile = Profile(spec.profile, line_spacing=spec.line_spacing)
    # ПАКЕТ 1: определения markdown-сносок (`[^id]: текст`) — рендерер абзацев
    # связывает маркеры [^id] в тексте с нативными w:footnote через этот словарь.
    profile.footnotes = dict(spec.footnotes)
    # ПАКЕТ 8: определения markdown концевых сносок (`[^!id]: текст`)
    profile.endnotes = dict(spec.endnotes)
    doc = Document()
    _configure_section(doc, profile)

    doc.core_properties.author = spec.output.author
    doc.core_properties.title = spec.title.main

    try:
        blocks.render_header(doc, spec.header, profile)
        blocks.render_title(doc, spec.title, profile)

        for block in spec.body:
            renderer = BLOCK_RENDERERS[block.type]
            renderer(doc, block, profile)

        # ПАКЕТ 4: комментарии из spec — фрагмент target ищется в собранных
        # параграфах и оборачивается в w:commentRangeStart/End (fail-closed,
        # если фрагмент нигде не найден). S2-H6: occurrence — N-е вхождение в
        # порядке документа; перепрыгивать на первое совпадение нельзя.
        for comment in spec.comments:
            remaining = comment.occurrence
            target = None
            local_occ = 1
            for p in doc.paragraphs:
                count = p.text.count(comment.target)
                if count == 0:
                    continue
                if remaining <= count:
                    target = p
                    local_occ = remaining
                    break
                remaining -= count
            if target is None:
                raise BuildError(
                    f"comment target not found in document: {comment.target!r} "
                    f"(occurrence {comment.occurrence})")
            add_comment(doc, target, comment.target, comment.text,
                        author=comment.author, date=comment.date,
                        occurrence=local_occ)

        # Фаза 1 В: колонтитулы и нумерация страниц (PAGE/NUMPAGES).
        if spec.footer_text or spec.page_numbers:
            blocks.render_footer(doc, spec, profile)

        # ПАКЕТ 1: TOC/поля страниц требуют автообновления при открытии в Word.
        if any(block.type == "toc" for block in spec.body) or spec.page_numbers:
            blocks.ensure_update_fields(doc)

        # ПАКЕТ 4: Режим отслеживания правок (Track Changes) в settings.xml
        from docx_cli.revisions import ensure_track_revisions
        if getattr(spec, "track_revisions", False) or any(
            "<w:ins" in (p._p.xml or "") or "<w:del" in (p._p.xml or "") for p in doc.paragraphs
        ):
            ensure_track_revisions(doc)

        # ПАКЕТ 9: Защита документа (readOnly, comments, trackedChanges, forms)
        if spec.protection:
            from docx_cli.protection import set_document_protection
            set_document_protection(doc, edit_mode=spec.protection, enforcement=True,
                                    password=spec.protection_password)
    except Exception as exc:  # noqa: BLE001
        # F-07 (аудит w2): BuildError рендереров (image file not found, comment
        # target not found, …) вылетал мимо конверта — вызывающие системы
        # получали исключение вместо JSON-envelope. Единый ответ: конверт.
        return {
            "ok": False,
            "status": "error",
            "reason": f"build_failed: {exc}",
            "parsed_at": now_iso(),
            "data": {"output_path": str(Path(output_path).resolve())},
            "metadata": {},
        }

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    # S3-M3: атомарность — сохраняем и валидируем временную копию,
    # по место вывода переносим только полностью собранный документ.
    # Падение валидации не стирает прежний хороший файл.
    import os as _os
    tmp_out = out.with_name(out.name + f".tmp-{_os.getpid()}")
    saved = False
    validation: dict[str, Any] | None = None
    try:
        doc.save(str(tmp_out))
        saved = True
        validation = validate.validate_document(str(tmp_out), spec)
    except Exception as exc:
        if saved:
            try:
                os.remove(str(tmp_out))
            except FileNotFoundError:
                pass
        return {
            "ok": False,
            "status": "error",
            "reason": f"build_failed: {exc}",
            "parsed_at": now_iso(),
            "data": {"output_path": str(out.resolve())},
            "metadata": {},
        }

    result: dict[str, Any] = {
        "ok": True,
        "status": "ok",
        "reason": None,
        "parsed_at": now_iso(),
        "data": {
            "output_path": str(out.resolve()),
            "profile": profile.name,
            "blocks_rendered": len(spec.body),
        },
        "metadata": validation,
    }

    if validation["errors"]:
        # S3-M3: fail-closed до переноса — прежний файл остаётся нетронутым
        try:
            os.remove(str(tmp_out))
        except FileNotFoundError:
            pass
        result["ok"] = False
        result["status"] = "error"
        result["reason"] = "validation_failed: " + "; ".join(validation["errors"])
        return result

    # Полностью собран и провалидирован — атомарно переносим на место вывода.
    _os.replace(str(tmp_out), str(out))

    if validation["warnings"]:
        result["status"] = "degraded"

    return result
