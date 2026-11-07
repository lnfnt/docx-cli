# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-H3: защита документа без пароля запрещена.

Тег w:documentProtection с enforcement="1" и без hash-атрибутов Word
трактует как защиту с пустым паролем (снимается в один клик) — раньше
docx-cli писал именно такой тег. Теперь: enforcement без пароля — ошибка
сборки; с паролем — salted SHA-1 (hashValue/salt/algorithmName).
"""
from __future__ import annotations


def _base_spec(protection: str, password=None) -> dict:
    spec = {
        "profile": "gost",
        "output": {"file_name": "p.docx", "author": "Т"},
        "header": {"align": "right", "lines": ["В суд"]},
        "title": {"main": "Т", "subtitle": None},
        "body": [{"type": "paragraph", "text": "Текст.", "inline_markdown": False}],
        "checks": {"no_yo": False, "no_en_dash": False,
                   "require_norms_refs": False},
        "protection": protection,
    }
    if password is not None:
        spec["protection_password"] = password
    return spec


def test_protection_without_password_fails(tmp_path):
    import pytest
    from docx_cli.builder import build_document
    from docx_cli.errors import BuildError
    from docx_cli.spec import DocSpec
    out = tmp_path / "p.docx"
    # F-07 (аудит w2): конверт error вместо исключения; fail-closed сохранён.
    res = build_document(
        DocSpec.from_dict(_base_spec("read_only")), str(out))
    assert res["ok"] is False and res["status"] == "error"
    assert "password" in res["reason"].lower()
    assert not out.exists(), "документ с защитой-театром создан"


def test_protection_with_password_has_hash(tmp_path):
    import zipfile
    from docx_cli.builder import build_document
    from docx_cli.spec import DocSpec
    out = tmp_path / "p.docx"
    res = build_document(
        DocSpec.from_dict(_base_spec("comments_only", "pw123")), str(out))
    assert res["ok"], res
    with zipfile.ZipFile(out) as z:
        xml = z.read("word/settings.xml").decode()
    assert 'w:edit="comments"' in xml
    assert 'w:hashValue="' in xml and 'w:salt="' in xml
