# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S3-H4: плейсхолдеры внутри w:hyperlink видны рендеру.

python-docx p.runs гиперссылки не возвращает — недозаменённый {{var}}
внутри ссылки молча проскакивал со status ok и пустым missing.
"""
from __future__ import annotations

import json
import subprocess
import sys


def _make_hyperlink_template(tmp_path):
    import docx
    from docx.oxml.ns import qn
    d = docx.Document()
    part = d.part
    rid = part.relate_to(
        "https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True)
    p = d.add_paragraph()
    hle = p._p.makeelement(qn("w:hyperlink"), {qn("r:id"): rid})
    p._p.append(hle)
    run = p.add_run("{{secret}}")
    r_el = run._r
    p._p.remove(r_el)
    hle.append(r_el)
    tpl = tmp_path / "tpl_hl.docx"
    d.save(str(tpl))
    return tpl


def test_hyperlink_placeholder_reported(tmp_path):
    tpl = _make_hyperlink_template(tmp_path)
    out = tmp_path / "out.docx"
    r = subprocess.run(
        [sys.executable, "-m", "docx_cli", "render", str(tpl), "-",
         "-o", str(out)],
        input=json.dumps({}), capture_output=True, text=True)
    payload = json.loads(r.stdout)
    assert payload["status"] == "degraded", payload
    assert "secret" in payload.get("missing", []), payload


def test_hyperlink_placeholder_replaced(tmp_path):
    tpl = _make_hyperlink_template(tmp_path)
    out = tmp_path / "out.docx"
    r = subprocess.run(
        [sys.executable, "-m", "docx_cli", "render", str(tpl), "-",
         "-o", str(out)],
        input=json.dumps({"secret": "VALUE"}), capture_output=True, text=True)
    payload = json.loads(r.stdout)
    assert payload["status"] == "ok", payload
    import docx
    d = docx.Document(str(out))
    assert "VALUE" in d.paragraphs[0].text
