# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""docx-cli template - библиотека бизнес-шаблонов builtin_templates.

Каждый шаблон: JSON (name/description/profile/author) + соседний .md.
build -> md_engine (v2) -> builder (полный профильный конвейер + валидация).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

from docx_cli.errors import SpecError
from docx_cli.md_engine import MdEngineError

TEMPLATES_DIR = Path(__file__).parent / "builtin_templates"

_TEMPLATE_NAME_RE = re.compile(r"^[a-z0-9_-]+$")


def _template_meta(name: str) -> dict:
    if not isinstance(name, str) or _TEMPLATE_NAME_RE.fullmatch(name) is None:
        raise SpecError("unknown template: invalid name")
    path = (TEMPLATES_DIR / f"{name}.json").resolve()
    if not path.is_file() or TEMPLATES_DIR.resolve() not in path.parents:
        raise SpecError(f"unknown template: {name}")
    return json.loads(path.read_text(encoding="utf-8-sig"))


def list_templates() -> list[dict]:
    out = []
    for path in sorted(TEMPLATES_DIR.glob("*.json")):
        meta = json.loads(path.read_text(encoding="utf-8-sig"))
        out.append({
            "name": meta.get("name"),
            "description": meta.get("description"),
            "profile": meta.get("profile"),
            "md": path.with_suffix(".md").name,
        })
    return out


def load_template(name: str) -> dict:
    meta = _template_meta(name)
    md_path = TEMPLATES_DIR / f"{name}.md"
    if not md_path.is_file():
        raise SpecError(f"template markdown missing for: {name}")
    meta["markdown"] = md_path.read_text(encoding="utf-8-sig")
    return meta


def build_template(name: str, output_docx: str, profile: str | None = None,
                   author: str | None = None) -> dict:
    """md-движок v2 -> DOCX; возвращает конверт builder'a."""
    from docx_cli.md_engine import md_build_v2

    meta = _template_meta(name)
    md_path = TEMPLATES_DIR / f"{name}.md"
    if not md_path.is_file():
        raise SpecError(f"template markdown missing for: {name}")
    return md_build_v2(
        md_path.read_text(encoding="utf-8-sig"), output_docx,
        author or meta.get("author", "docx-cli"),
        profile or meta.get("profile", "corporate"))


__all__ = ["TEMPLATES_DIR", "list_templates", "load_template", "build_template",
           "MdEngineError"]
