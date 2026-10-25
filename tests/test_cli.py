"""Tests for CLI commands and envelope shape."""
from __future__ import annotations

import json
import subprocess
import sys
from pathlib import Path

import pytest

from docx_cli.cli import main

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


class TestBuildCli:
    def test_build_exit_0(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "cli_out.docx"
        code, payload, _ = run_cli(["build", str(example_spec_path), str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert payload["status"] == "ok"
        assert out.exists()

    def test_build_envelope_shape(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "cli_out2.docx"
        code, payload, _ = run_cli(["build", str(example_spec_path), str(out)])
        assert code == 0
        assert set(payload.keys()) == {"ok", "status", "reason", "parsed_at", "data", "metadata"}

    def test_build_profile_override(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "cli_compact.docx"
        code, payload, _ = run_cli(["build", "--profile", "compact", str(example_spec_path), str(out)])
        assert code == 0
        assert payload["data"]["profile"] == "compact"
        assert out.exists()

    def test_build_bad_spec_exit_1(self, tmp_workspace):
        bad = tmp_workspace / "bad.json"
        bad.write_text('{"unknown": 1}', encoding="utf-8")
        out = tmp_workspace / "out.docx"
        code, payload, _ = run_cli(["build", str(bad), str(out)])
        assert code == 1
        assert payload["ok"] is False
        assert payload["status"] == "error"


class TestValidateCli:
    def test_validate_exit_0(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "val_out.docx"
        run_cli(["build", str(example_spec_path), str(out)])
        code, payload, _ = run_cli(["validate", str(out)])
        assert code == 0
        assert payload["ok"] is True
        assert payload["status"] == "ok"

    def test_validate_strict_exit_2_on_warnings(self, tmp_workspace):
        spec = {
            "profile": "gost",
            "output": {"file_name": "x.docx", "author": "A"},
            "header": {"align": "right", "lines": ["Суд"]},
            "title": {"main": "ТИТУЛ"},
            "body": [
                {"type": "paragraph", "text": "[УКАЗАТЬ сумму]", "inline_markdown": False}
            ],
            "checks": {"no_yo": True, "no_en_dash": True, "require_norms_refs": False},
        }
        spec_path = tmp_workspace / "warn_spec.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        out = tmp_workspace / "warn_out.docx"
        run_cli(["build", str(spec_path), str(out)])
        code, payload, _ = run_cli(["validate", str(out), "--strict"])
        assert code == 2
        assert payload["status"] == "degraded"

    def test_validate_missing_file_exit_1(self, tmp_workspace):
        missing = tmp_workspace / "nope.docx"
        code, payload, _ = run_cli(["validate", str(missing)])
        assert code == 1
        assert payload["ok"] is False


class TestDoctorCli:
    def test_doctor_exit_0(self):
        code, payload, _ = run_cli(["doctor"])
        assert code == 0
        assert payload["ok"] is True
        assert "python_version" in payload["data"]["checks"]
        assert "python_docx" in payload["data"]["checks"]
        assert "lxml" in payload["data"]["checks"]
        assert "write_permissions" in payload["data"]["checks"]


class TestMainEntrypoint:
    def test_main_help_returns_non_zero(self):
        with pytest.raises(SystemExit) as exc:
            main(["--help"])
        assert exc.value.code == 0

    def test_main_doctor(self):
        code = main(["doctor"])
        assert code == 0

    def test_main_build(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "main_out.docx"
        code = main(["build", str(example_spec_path), str(out)])
        assert code == 0
        assert out.exists()

    def test_main_validate(self, example_spec_path, tmp_workspace):
        out = tmp_workspace / "main_val.docx"
        main(["build", str(example_spec_path), str(out)])
        code = main(["validate", str(out)])
        assert code == 0

    def test_main_validate_strict(self, tmp_workspace):
        spec = {
            "profile": "gost",
            "output": {"file_name": "x.docx", "author": "A"},
            "header": {"align": "right", "lines": ["Суд"]},
            "title": {"main": "ТИТУЛ"},
            "body": [
                {"type": "paragraph", "text": "[УКАЗАТЬ]", "inline_markdown": False}
            ],
            "checks": {"no_yo": True, "no_en_dash": True, "require_norms_refs": False},
        }
        spec_path = tmp_workspace / "s.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        out = tmp_workspace / "o.docx"
        main(["build", str(spec_path), str(out)])
        code = main(["validate", str(out), "--strict"])
        assert code == 2
