# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""docx-cli update: безопасное обновление пакета (шаблоны/настройки не трогаем)."""
from __future__ import annotations

import re

import json
import subprocess
import sys
import urllib.request
from typing import Any, Callable

from docx_cli import __version__

Runner = Callable[[list[str]], "subprocess.CompletedProcess"]

PYPI_URL = "https://pypi.org/pypi/docx-cli/json"


def _default_runner(cmd: list[str]):
    return subprocess.run(cmd, capture_output=True, text=True)


def _pypi_latest() -> str:
    with urllib.request.urlopen(PYPI_URL, timeout=10) as resp:
        data = json.load(resp)
    return data["info"]["version"]


def _git_latest(target: str | None) -> str:
    # ponytail: assume target remote on branch main/master; tag-based once needed
    cmd = ["git", "ls-remote", "--tags", target or "https://github.com/anomalyco/docx-cli.git"]
    cp = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
    if cp.returncode != 0:
        raise RuntimeError(f"git ls-remote failed: {cp.stderr.strip()}")
    tags = [line.rsplit("/", 1)[-1].lstrip("v") for line in cp.stdout.splitlines() if "tags/" in line and "^{}" not in line]
    if not tags:
        raise RuntimeError("no tags found on remote")
    return max(tags, key=lambda v: [int(p) for p in v.split(".") if p.isdigit()] or [0])


def _cmp(v1: str, v2: str) -> int:
    # S3-L: наивное сравнение [1,2] vs [1,2,0] считало их разными;
    # дополняем нулями до общей длины.
    def key(v: str):
        parts = [int(p) for p in v.split(".") if p.isdigit()]
        return parts or [0]
    k1, k2 = key(v1), key(v2)
    n = max(len(k1), len(k2))
    k1 += [0] * (n - len(k1))
    k2 += [0] * (n - len(k2))
    return (k1 > k2) - (k1 < k2)



def _valid_version_target(target: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._+\-]*", str(target)))


def _valid_git_target(target: str) -> bool:
    t = str(target).strip()
    return t.startswith(("http://", "https://", "git://", "ssh://", "git@"))


def check_update(source: str = "pypi", target: str | None = None) -> dict[str, Any]:
    if source not in ("pypi", "git"):
        return {"ok": False, "status": "invalid_usage",
                "reason": f"invalid source: {source}", "data": {"data_safe": True}}
    # S3-H2: git-источник из агентского/MCP-контура запрещён (как в run_update);
    # check_update с произвольным URL иначе делает исходящий ls-remote
    # на непроверенный хост — та же поверхность атаки, что и pip install.
    if source == "git":
        return {"ok": False, "status": "invalid_usage",
                "reason": "update_disabled: git source is not allowed; only PyPI package docx-cli",
                "data": {"data_safe": True}}
    if target is not None and not _valid_version_target(target):
        return {"ok": False, "status": "invalid_usage",
                "reason": f"invalid target: {target}",
                "data": {"data_safe": True}}
    try:
        latest = _pypi_latest() if source == "pypi" else _git_latest(target)
    except Exception as exc:  # noqa: BLE001 — fail-closed
        return {"ok": False, "status": "error", "reason": f"check_failed: {exc}",
                "data": {"data_safe": True}}
    current = __version__
    up_to_date = _cmp(latest, current) <= 0
    return {
        "ok": True,
        "status": "ok",
        "data": {"data_safe": True, "current_version": current,
                 "latest_version": latest, "up_to_date": up_to_date,
                 "source": source},
    }


def run_update(source: str = "pypi", target: str | None = None,
               runner: Runner | None = None) -> dict[str, Any]:
    runner = runner or _default_runner
    if source not in ("pypi", "git"):
        return {"ok": False, "status": "invalid_usage",
                "reason": f"invalid source: {source}",
                "data": {"data_safe": True, "data_preserved": True}}
    # S3-H2: git-источник из агентского/MCP-контура запрещён; allowlist = только PyPI-имя docx-cli.
    if source == "git":
        return {"ok": False, "status": "invalid_usage",
                "reason": "update_disabled: git source is not allowed; only PyPI package docx-cli",
                "data": {"data_safe": True, "data_preserved": True}}
    if target is not None and not _valid_version_target(target):
        return {"ok": False, "status": "invalid_usage",
                "reason": f"invalid target: {target}",
                "data": {"data_safe": True, "data_preserved": True}}
    if source == "pypi":
        spec = "docx-cli" if not target else f"docx-cli=={target}"
        args = ["install", "--upgrade", spec]
    else:
        url = target or "https://github.com/anomalyco/docx-cli.git"
        args = ["install", "--upgrade", f"git+{url}" if not url.startswith(("http", "git+", "ssh")) else url]
    cmd = [sys.executable, "-m", "pip", *args]
    try:
        cp = runner(cmd)
    except Exception as exc:  # noqa: BLE001 — fail-closed
        return {"ok": False, "status": "error", "reason": f"update_failed: {exc}",
                "data": {"data_safe": True, "data_preserved": True}}
    data = {"data_safe": True, "data_preserved": True, "source": source,
            "command": " ".join(cmd), "returncode": cp.returncode}
    if cp.returncode != 0:
        return {"ok": False, "status": "error",
                "reason": (cp.stderr or cp.stdout or "pip failed").strip()[-500:],
                "data": {**data, "data_preserved": True}}
    return {"ok": True, "status": "ok", "data": data}
