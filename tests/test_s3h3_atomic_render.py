# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S3-H3: рендер шаблона атомарен — падение после doc.save не оставляет
полуфайл и не портит прежний результат.

Раньше: doc.save(output) -> _patch_notes(...) — исключение в _patch_notes
оставляло полностью отрендеренный output.docx при error-конверте, а обрыв
перезаписи порч файл.
"""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path


def _make_template(tmp_path: Path) -> Path:
    import docx
    d = docx.Document()
    d.add_paragraph("Договор №{number} от {date}")
    p = tmp_path / "tpl.docx"
    d.save(str(p))
    return p


def test_render_atomic_on_patch_failure(tmp_path, monkeypatch):
    from docx_cli import templating
    tpl = _make_template(tmp_path)
    out = tmp_path / "out.docx"
    # подменяем _patch_notes на падение
    monkeypatch.setattr(templating, "_patch_notes",
                        lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom")))
    try:
        templating.render_template(str(tpl), {"number": "1", "date": "01.01.2026"},
                                   str(out))
        raised = False
    except Exception:
        raised = True
    assert raised, "исключение не прорвалось"
    assert not out.exists(), "полуфайл остался на диске при error-конверте"


def test_render_replaces_only_on_success(tmp_path):
    from docx_cli import templating
    tpl = _make_template(tmp_path)
    out = tmp_path / "out.docx"
    r1 = templating.render_template(str(tpl), {"number": "1", "date": "d1"}, str(out))
    assert r1["ok"] and out.exists()
    before = out.read_bytes()
    # повторный рендер с падением patch — прежний файл не должен измениться
    import docx_cli.templating as t
    orig = t._patch_notes
    t._patch_notes = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("boom"))
    try:
        t.render_template(str(tpl), {"number": "2", "date": "d2"}, str(out))
    except Exception:
        pass
    finally:
        t._patch_notes = orig
    assert out.read_bytes() == before, "прежний хороший файл испорчен падением"
