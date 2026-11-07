# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S1-H1: внешние relationship (TargetMode="External") — не unresolved.

Гиперссылки на внешние URL — валидная часть OOXML; валидатор бракует
собственный же md-конвейер (exit 1 на здоровом документе).
"""
from __future__ import annotations

import io

import docx
import pytest
from docx.oxml.ns import qn

from docx_cli import validate


def _doc_with_external_link(tmp_path) -> str:
    d = docx.Document()
    p = d.add_paragraph("текст ")
    part = d.part
    r_id = part.relate_to(
        "https://example.com/some/page",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    run = p.add_run("ссылка")
    hle = p._p.makeelement(qn("w:hyperlink"), {qn("r:id"): r_id})
    p._p.append(hle)
    hle.append(run._r)
    path = str(tmp_path / "ext.docx")
    d.save(path)
    return path


def test_external_hyperlink_not_unresolved(tmp_path):
    path = _doc_with_external_link(tmp_path)
    errors = validate._check_package_integrity(path)
    assert not any("unresolved_relationship" in e for e in errors), errors


def test_internal_missing_target_still_flagged(tmp_path):
    path = _doc_with_external_link(tmp_path)
    # подменяем TargetMode на Internal с несуществующей целью
    import zipfile, shutil
    from docx_cli.validate import DOC_RELS
    src = zipfile.ZipFile(path)
    out = str(tmp_path / "broken.docx")
    with zipfile.ZipFile(out, "w") as dst:
        for item in src.infolist():
            data = src.read(item.filename)
            if item.filename == DOC_RELS:
                data = data.replace(b'TargetMode="External"', b'')
            dst.writestr(item, data)
    src.close()
    errors = validate._check_package_integrity(out)
    assert any("unresolved_relationship" in e for e in errors), errors
