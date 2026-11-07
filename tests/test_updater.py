# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

import json
import subprocess

import pytest

from docx_cli import cli
from docx_cli.updater import check_update, run_update


def fake_ok(cmd):
    return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")


def fake_fail(cmd):
    return subprocess.CompletedProcess(cmd, 1, stdout="", stderr="boom")


def run_cli(argv):
    parser = cli.build_parser()
    args = parser.parse_args(argv)
    return args.func(args), parser


def test_check_up_to_date(monkeypatch):
    monkeypatch.setattr("docx_cli.updater._pypi_latest", lambda: "0.1.4")
    res = check_update("pypi")
    assert res["ok"] and res["data"]["data_safe"]
    assert res["data"]["up_to_date"] is True


def test_check_update_available(monkeypatch):
    monkeypatch.setattr("docx_cli.updater._pypi_latest", lambda: "9.9.9")
    res = check_update("pypi")
    assert res["data"]["up_to_date"] is False


def test_check_fail_closed(monkeypatch):
    def boom():
        raise RuntimeError("network down")
    monkeypatch.setattr("docx_cli.updater._pypi_latest", boom)
    res = check_update("pypi")
    assert not res["ok"] and "network down" in res["reason"]
    assert res["data"]["data_safe"] is True


def test_check_invalid_source():
    res = check_update("ftp")
    assert not res["ok"] and res["status"] == "invalid_usage"


def test_check_git(monkeypatch):
    # S3-H2: git-источник запрещён и в check_update — иначе исходящий
    # ls-remote на произвольный непроверенный хост.
    calls = []

    def _boom(target):
        calls.append(target)
        return "1.2.3"

    monkeypatch.setattr("docx_cli.updater._git_latest", _boom)
    res = check_update("git", target="https://example.com/x.git")
    assert not res["ok"] and res["status"] == "invalid_usage"
    assert calls == [], "ls-remote до непроверенного хоста всё же выполнен"


def test_run_update_builds_pip_command():
    captured = {}

    def runner(cmd):
        captured["cmd"] = cmd
        return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

    res = run_update("pypi", runner=runner)
    assert res["ok"] and res["data"]["data_preserved"]
    assert captured["cmd"][0].endswith("python") or "python" in captured["cmd"][0]
    assert "-m" in captured["cmd"] and "pip" in captured["cmd"]
    assert "--upgrade" in captured["cmd"]


def test_run_update_pip_failure(monkeypatch):
    res = run_update("pypi", runner=fake_fail)
    assert not res["ok"]
    assert res["data"]["data_preserved"] is True


def test_run_update_invalid_source():
    res = run_update("ftp", runner=fake_ok)
    assert not res["ok"]


def test_cli_check_json(capsys, monkeypatch):
    monkeypatch.setattr("docx_cli.updater._pypi_latest", lambda: "0.1.4")
    code, _ = run_cli(["update", "--check"])
    out = capsys.readouterr().out
    data = json.loads(out)
    assert code == 0 and data["ok"] is True and data["data"]["source"] == "pypi"


def test_cli_update_mock(monkeypatch, capsys):
    monkeypatch.setattr("docx_cli.updater.run_update",
                        lambda *a, **kw: {"ok": True, "status": "ok", "data": {}})
    code, _ = run_cli(["update", "--source", "git"])
    assert code == 0
    assert json.loads(capsys.readouterr().out)["ok"] is True


def test_cli_update_error(capsys):
    monkeypatch = pytest.MonkeyPatch()
    monkeypatch.setattr("docx_cli.updater.run_update", lambda *a, **kw: {"ok": False, "status": "error", "reason": "x"})
    code, _ = run_cli(["update", "--source", "git"])
    monkeypatch.undo()
    assert code == 1
    assert json.loads(capsys.readouterr().out)["ok"] is False
