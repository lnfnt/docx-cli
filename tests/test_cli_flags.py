# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""TASK_FLAG_FIX: поддержка -o / --output для build и md-build."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

PYTHON = sys.executable


def run_cli(args, cwd=None):
    cmd = [str(PYTHON), "-m", "docx_cli"] + args
    result = subprocess.run(
        cmd, capture_output=True, text=True, cwd=str(cwd or Path(__file__).parent.parent)
    )
    try:
        payload = json.loads(result.stdout.splitlines()[0]) if result.stdout else {}
    except json.JSONDecodeError:
        payload = {}
    return result.returncode, payload, result


class TestBuildOutputFlags:
    def test_build_short_o(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "o.docx"
        code, payload, _ = run_cli(["build", str(example_spec_path), "-o", str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert out.exists()

    def test_build_long_output(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "o.docx"
        code, payload, _ = run_cli(["build", str(example_spec_path), "--output", str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert out.exists()

    def test_build_positional_backward_compatible(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "o.docx"
        code, payload, _ = run_cli(["build", str(example_spec_path), str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert out.exists()

    def test_build_missing_output_errors(self, example_spec_path):
        code, payload, _ = run_cli(["build", str(example_spec_path)])
        assert code == 2
        assert payload["ok"] is False
        assert "missing output file" in payload["reason"]


class TestMdBuildOutputFlags:
    def _md(self, tmp_workspace: Path) -> str:
        md = tmp_workspace / "input.md"
        md.write_text("# ИСКОВОЕ ЗАЯВЛЕНИЕ\n\nПросто абзац текста.\n", encoding="utf-8")
        return str(md)

    def test_md_build_short_o(self, tmp_workspace):
        src = self._md(tmp_workspace)
        out = tmp_workspace / "o.docx"
        code, payload, _ = run_cli(["md-build", src, "-o", str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert out.exists()

    def test_md_build_long_output(self, tmp_workspace):
        src = self._md(tmp_workspace)
        out = tmp_workspace / "o.docx"
        code, payload, _ = run_cli(["md-build", src, "--output", str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert out.exists()
