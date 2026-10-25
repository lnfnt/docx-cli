"""Shared fixtures and helpers for tests."""
from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

os.environ.setdefault("TMPDIR", tempfile.mkdtemp(prefix="docx_cli_"))
tempfile.tempdir = os.environ["TMPDIR"]

# Keep CLI tests portable: use the interpreter running pytest instead of assuming
# the release archive contains a project-local .venv.
PYTHON = sys.executable


def run_cli(*args: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run([PYTHON, "-m", "docx_cli", *args], capture_output=True, text=True)

import pytest

from docx_cli.spec import DocSpec


@pytest.fixture
def tmp_workspace():
    path = tempfile.mkdtemp(prefix="docx_cli_test_")
    yield Path(path)
    shutil.rmtree(path, ignore_errors=True)


@pytest.fixture
def minimal_spec_dict():
    return {
        "profile": "gost",
        "output": {"file_name": "test.docx", "author": "Тестов Т.Т."},
        "header": {"align": "right", "lines": ["В тестовый суд", "Истец: Тестов Т.Т."]},
        "title": {"main": "ТЕСТОВЫЙ ДОКУМЕНТ", "subtitle": "к делу № 1"},
        "body": [
            {"type": "paragraph", "text": "Текст документа для проверки.", "inline_markdown": False}
        ],
        "checks": {"no_yo": True, "no_en_dash": True, "require_norms_refs": False},
    }


@pytest.fixture
def minimal_spec(minimal_spec_dict):
    return DocSpec.from_dict(minimal_spec_dict)


@pytest.fixture
def built_path(tmp_workspace, minimal_spec):
    from docx_cli.builder import build_document

    out = tmp_workspace / "out.docx"
    build_document(minimal_spec, str(out))
    return out


@pytest.fixture
def example_spec_path():
    return Path(__file__).parent.parent / "examples" / "spec_example.json"
