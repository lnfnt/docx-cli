"""v0.1.0 §4: релизные метаданные и CJK-гейт файлов волны v0.1.0."""
from __future__ import annotations

import re
from pathlib import Path

ROOT = Path(__file__).parent.parent

V0100_FILES = [
    "docx_cli/mcp_server.py",
    "docx_cli/md_court.py",
    "docx_cli/cli.py",
    "docx_cli/spec.py",
    "templates/vs_sud.json",
    "templates/soyu_rayon.json",
    "docs/design/nested-lists-tables.md",
    "tests/fixtures/v0100_legacy_md.json",
    "CHANGELOG.md",
    "README.md",
]

CJK_RE = re.compile("[\u3040-\u30ff\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff"
                    "\uff01-\uff60]")


def test_v0100_version_bumped_synchronously():
    import tomllib
    import docx_cli
    from importlib.metadata import version as dist_version
    pyproject_version = tomllib.loads(
        (ROOT / "pyproject.toml").read_text(encoding="utf-8"))["project"]["version"]
    assert docx_cli.__version__ == "0.1.1"
    assert pyproject_version == "0.1.1"
    assert docx_cli.__version__


def test_v0100_changelog_has_0100_section():
    text = (ROOT / "CHANGELOG.md").read_text(encoding="utf-8")
    assert "## 0.1.0" in text
    section = text.split("## 0.1.0", 1)[1].split("## 0.0.0", 1)[0]
    for token in ("mcp", "B-10", "T-05", "needs_confirmation"):
        assert token in section, token


def test_v0100_readme_documents_mcp_and_profiles():
    text = (ROOT / "README.md").read_text(encoding="utf-8")
    for token in ("docx-cli mcp", "needs_confirmation", "vs_sud",
                  "soyu_rayon", "block_in_cell", "list_depth_gt4"):
        assert token in text, token


def test_v0100_cjk_gate_wave_files():
    """CJK-гейт: ни один файл волны v0.1.0 не содержит CJK-иконографики
    (документы русские — кириллица; идеограммы = мусор кодировок)."""
    offenders = []
    for rel in V0100_FILES:
        p = ROOT / rel
        assert p.is_file(), f"missing wave file: {rel}"
        hits = CJK_RE.findall(p.read_text(encoding="utf-8"))
        if hits:
            offenders.append((rel, hits[:5]))
    assert offenders == [], offenders
