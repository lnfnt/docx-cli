"""v0.9.0 пины контроля (A-06 verify-signature, S-10 a2a-args):
регрессионные пины на выпущенное поведение."""
from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path

import pytest

import sys

CLI = [sys.executable, "-m", "docx_cli"]
KEY_TEMPLATE = """%echo genkey
Key-Type: RSA
Key-Length: 2048
Name-Real: Docx Test
Name-Email: test@example.invalid
Expire-Date: 0
%no-protection
%commit
"""


def _cli(*args, env_extra=None):
    import os
    if env_extra:
        e = {**os.environ, **env_extra}
        for k, v in list(e.items()):
            if not isinstance(v, str):
                e[k] = str(v)
    else:
        e = None
    return subprocess.run([*CLI, *args], capture_output=True, text=True, env=e)


@pytest.fixture(scope='module')
def signed_dist(tmp_path_factory):
    home = tmp_path_factory.mktemp("gnupg")
    dist = tmp_path_factory.mktemp("dist")
    env = {**__import__('os').environ, "GNUPGHOME": str(home), "HOME": str(home)}
    home.chmod(0o700)
    (home / "tmpl").write_text(KEY_TEMPLATE)
    r = subprocess.run(["gpg", "--batch", "--gen-key", str(home / "tmpl")],
                       capture_output=True, text=True, env=env, timeout=120)
    assert r.returncode == 0, r.stderr
    (dist / "pkg.whl").write_bytes(b"fake wheel bytes")
    sums = subprocess.run(["sha256sum", "pkg.whl"], cwd=dist, capture_output=True, text=True, env=env)
    (dist / "SHA256SUMS").write_text(sums.stdout)
    r2 = subprocess.run(["gpg", "--batch", "--yes", "--detach-sign", "-o", "SHA256SUMS.asc", "SHA256SUMS"],
                        cwd=dist, capture_output=True, text=True, env=env, timeout=60)
    assert r2.returncode == 0, r2.stderr
    return dist, env


def _doctor_sig(dist, env=None):
    return _cli('doctor', '--verify-signature', '--dist', str(dist), env_extra=env)


def test_verify_signature_ok(signed_dist):
    dist, env = signed_dist
    if shutil.which('gpg') is None:
        pytest.skip('gpg not installed')
    p = _doctor_sig(dist, env)
    env_out = json.loads(p.stdout)
    blob = json.dumps(env_out)
    assert 'ok' in str(env_out.get('data', {}).get('verify_signature', '')).lower() or '"verify_signature": "ok"' in blob, blob[:300]


def test_verify_signature_fail_on_tamper(signed_dist):
    dist, env = signed_dist
    if shutil.which('gpg') is None:
        pytest.skip('gpg not installed')
    (dist / "pkg.whl").write_bytes(b"TAMPERED payload")
    p = _doctor_sig(dist, env)
    out = json.loads(p.stdout)
    sig = out.get('data', {}).get('verify_signature') or out.get('verify_signature') or json.dumps(out)[:0]
    assert str(sig).lower().startswith('fail'), str(out)[:300]


def test_verify_signature_unsigned(tmp_path):
    dist = tmp_path / "nodist"
    dist.mkdir()
    (dist / "x.whl").write_bytes(b'x')
    p = _doctor_sig(dist, None)
    out = json.loads(p.stdout)
    sig = out.get('data', {}).get('verify_signature') or out.get('verify_signature')
    assert sig == 'unsigned', str(out)[:200]


def test_a2a_args_present_and_bus_free(tmp_path):
    from docx import Document
    f = tmp_path / 'clean.docx'
    Document().save(f)
    p = _cli('scan', str(f), '--json', '--a2a-args')
    out = json.loads(p.stdout)
    cmd = out.get('data', {}).get('a2a_command') or out.get('a2a_command')
    assert cmd and cmd.startswith('a2a log'), str(out)[:200]
    assert '--verify' in cmd
    # bus-free: команда только напечатана — без флага поля нет
    p2 = _cli('scan', str(f), '--json')
    plain = json.loads(p2.stdout)
    assert 'a2a_command' not in (plain.get('data') or plain)
