"""Core document builder."""
from __future__ import annotations

import os
from pathlib import Path
from typing import Any

from docx import Document
from docx.shared import Cm

from docx_cli import blocks, validate
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
    "table": blocks.render_table,
    "signature_block": blocks.render_signature_block,
    "attachments_list": blocks.render_attachments_list,
    "pleading_marker": blocks.render_pleading_marker,
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
    doc = Document()
    _configure_section(doc, profile)

    doc.core_properties.author = spec.output.author
    doc.core_properties.title = spec.title.main

    blocks.render_header(doc, spec.header, profile)
    blocks.render_title(doc, spec.title, profile)

    for block in spec.body:
        renderer = BLOCK_RENDERERS[block.type]
        renderer(doc, block, profile)

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    saved = False
    validation: dict[str, Any] | None = None
    try:
        doc.save(str(out))
        saved = True
        validation = validate.validate_document(str(out), spec)
    except Exception as exc:
        if saved:
            try:
                os.remove(str(out))
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
        try:
            os.remove(str(out))
        except FileNotFoundError:
            pass
        result["ok"] = False
        result["status"] = "error"
        result["reason"] = "validation_failed: " + "; ".join(validation["errors"])
        return result

    if validation["warnings"]:
        result["status"] = "degraded"

    return result
