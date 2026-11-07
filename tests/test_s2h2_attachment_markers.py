# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-H2: приложения получают явный маркер «Приложение N».

Без маркера ссылки «см. Приложение 1» в теле документа висят, а перечень
превращается в безымянный список (осиротевшие приложения).
"""
from __future__ import annotations


def _spec_with_attachments(items):
    return {
        "profile": "gost",
        "output": {"file_name": "a.docx", "author": "Т"},
        "header": {"align": "right", "lines": ["В суд"]},
        "title": {"main": "Т", "subtitle": None},
        "body": [{"type": "attachments_list", "items": items}],
        "checks": {"no_yo": False, "no_en_dash": False,
                   "require_norms_refs": False},
    }


def test_attachment_markers_present(tmp_path):
    import docx
    from docx_cli.builder import build_document
    from docx_cli.spec import DocSpec
    out = tmp_path / "a.docx"
    build_document(DocSpec.from_dict(
        _spec_with_attachments(["Копия договора", "Выписка"])), str(out))
    d = docx.Document(str(out))
    texts = [p.text.strip() for p in d.paragraphs]
    assert "Приложение 1" in texts, texts
    assert "Приложение 2" in texts, texts
    assert "Копия договора" in texts
