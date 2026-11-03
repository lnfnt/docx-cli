# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-H6: якорение комментариев при повторном target.

Раньше замечание молча цеплялось к первому вхождению target в документе —
при повторах это ошибка адресации. Теперь: явный occurrence (1-based, по
порядку документа); выход за диапазон — fail-closed.
"""
from __future__ import annotations


def _minimal_spec(extra: dict) -> dict:
    spec = {
        "profile": "gost",
        "output": {"file_name": "c.docx", "author": "Тест"},
        "header": {"align": "right", "lines": ["В тестовый суд"]},
        "title": {"main": "ТЕСТ", "subtitle": None},
        "body": [
            {"type": "paragraph",
             "text": "Основания указаны ниже. Ещё раз: Основания.",
             "inline_markdown": False},
            {"type": "paragraph",
             "text": "Здесь тоже Основания есть.", "inline_markdown": False},
        ],
        "checks": {"no_yo": False, "no_en_dash": False,
                   "require_norms_refs": False},
    }
    spec.update(extra)
    return spec


def test_comment_occurrence_second_across_paragraphs(tmp_path):
    import docx
    from docx_cli.builder import build_document
    from docx_cli.spec import DocSpec
    spec = _minimal_spec({"comments": [
        {"target": "Основания", "text": "замечание ко второму",
         "occurrence": 2}]})
    out = tmp_path / "c.docx"
    build_document(DocSpec.from_dict(spec), str(out))
    d = docx.Document(str(out))
    # комментарий должен быть на втором вхождении (в первом body-абзаце,
    # после «Ещё раз: »), а не на первом и не в другом абзаце
    anchored = [p for p in d.paragraphs
                if p._p.xml.count("commentRangeStart") == 1]
    assert len(anchored) == 1
    xml = anchored[0]._p.xml
    import re
    m = re.search(r"commentRangeStart.*?<w:t[^>]*>([^<]*)", xml, re.DOTALL)
    assert m and m.group(1).strip().startswith("Основания"), \
        m.group(1) if m else None
    assert "Ещё раз" in xml.split("commentRangeStart")[0][-200:], \
        "якорь не на втором вхождении"


def test_comment_occurrence_out_of_range_fails(tmp_path):
    import pytest
    from docx_cli.builder import build_document
    from docx_cli.errors import BuildError
    from docx_cli.spec import DocSpec
    spec = _minimal_spec({"comments": [
        {"target": "Основания", "text": "x", "occurrence": 9}]})
    # F-07 (аудит w2): конверт error вместо исключения.
    res = build_document(DocSpec.from_dict(spec), str(tmp_path / "c.docx"))
    assert res["ok"] is False and res["status"] == "error"
    assert "comment target not found" in res["reason"]
