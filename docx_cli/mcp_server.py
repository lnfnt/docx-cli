"""docx-cli MCP stdio server: JSON-RPC 2.0 tool surface (v0.1.0, A-05).

Паттерн sud-cli (mcp_server.py): tools/list, tools/call, единый content-конверт
с isError; алиас-толерантность параметров через _pick (два имени с разными
значениями — явная ошибка, не молчаливый выбор). Межрепозиторийной зависимости
нет — это наша реализация того же контракта.

Права (sud W6a): каждый инструмент несёт поле permission read|write (и
permission-поле в каждой inputSchema); write-тулы (docx_build, docx_normalize,
docx_md_build) требуют явный confirm: true — без него status=needs_confirmation
и НИКАКИХ побочных эффектов (файл не пишется). needs_confirmation — данные,
не транспортная ошибка (isError=False).

Схемы: type object, required минимальный, additionalProperties:false —
неизвестный параметр явная ошибка, а не молчаливое игнорирование (fail-closed).
"""
from __future__ import annotations

import argparse
import io
import json
import sys
import tempfile
from contextlib import redirect_stdout
from pathlib import Path
from typing import Any

import docx_cli
from docx_cli import __version__ as _PKG_VERSION

CONFIRM_PROPERTY = {
    "type": "boolean",
    "description": "write-инструмент: без confirm=true вернёт "
                   "status=needs_confirmation, запись не производится",
}


def _schema(properties: dict, required: list[str] | None = None,
            write: bool = False) -> dict:
    props = dict(properties)
    props["permission"] = dict(PROPERTY_PERMISSION_DEFAULT)
    if write:
        props["confirm"] = dict(CONFIRM_PROPERTY)
    return {"type": "object", "properties": props,
            "required": required or [], "additionalProperties": False}


# permission-поле объявлено в КАЖДОЙ схеме (контракт §1); значение сверяется
# с декларируемым правом тула при диспатче (fail-closed на mismatch).
PROPERTY_PERMISSION_DEFAULT = {
    "type": "string",
    "enum": ["read", "write"],
    "description": "декларируемое право: read|write; несовпадение с правом "
                   "тула — явная ошибка",
}

_PATH = {"type": "string", "description": "путь к файлу"}

TOOLS: list[dict[str, Any]] = [
    {"name": "docx_build", "permission": "write",
     "description": "Собрать .docx из DocSpec (машинный контракт профиля "
                    "gost|compact|court|vs_sud|soyu_rayon) по указанному пути. "
                    "write: требует confirm=true.",
     "schema": _schema({
         "spec": {"type": "object",
                  "description": "DocSpec JSON (profile/output/header/title/body)"},
         "out": {**_PATH, "description": "путь итогового .docx; CLI-алиас: output"},
         "output": {**_PATH, "description": "CLI-алиас out"},
     }, required=["spec"], write=True)},
    {"name": "docx_validate", "permission": "read",
     "description": "Проверить существующий .docx валидатором оформления "
                    "(gost|review). Только чтение.",
     "schema": _schema({
         "file": {**_PATH, "description": "путь к .docx; CLI-алиас: path"},
         "path": {**_PATH, "description": "CLI-алиас file"},
         "mode": {"type": "string", "enum": ["gost", "review"], "default": "gost"},
         "strict": {"type": "boolean", "default": False,
                    "description": "w/skip -> warnings_present semantics (as CLI --strict)"}},
         required=["file"])},
    {"name": "docx_scan", "permission": "read",
     "description": "Аудит входящего .docx на скрытые LLM-векторы "
                    "(inject_scan как есть: vanish/микрошрифт/tracked/…). "
                    "Только чтение.",
     "schema": _schema({
         "file": {**_PATH, "description": "путь к .docx; CLI-алиас: path"},
         "path": {**_PATH, "description": "CLI-алиас file"},
         "a2a_args": {"type": "boolean", "default": False,
                      "description": "добавить bus-free a2a_command в конверт"}})},
    {"name": "docx_normalize", "permission": "write",
     "description": "Нормализовать шрифты/интервалы .docx к Times New Roman. "
                    "write: требует confirm=true; каталог — только yes=true.",
     "schema": _schema({
         "path": {**_PATH, "description": "файл или каталог .docx"},
         "file": {**_PATH, "description": "CLI-алиас path"},
         "body_size": {"type": "integer", "default": 12},
         "header_size": {"type": "integer", "default": 11},
         "dry_run": {"type": "boolean", "default": False,
                     "description": "показать изменения без записи"},
         "backup": {"type": "boolean", "default": True},
         "yes": {"type": "boolean", "default": False,
                 "description": "подтвердить каталожный режим"}},
         required=["path"], write=True)},
    {"name": "docx_md_build", "permission": "write",
     "description": "Собрать процессуальный .docx из Markdown (шапка, титул, "
                    "**bold**, вложенные списки и pipe-таблицы v0.10). "
                    "write: требует confirm=true.",
     "schema": _schema({
         "md": {"type": "string",
                "description": "текст markdown; альтернатива md_path"},
         "md_path": {**_PATH, "description": "путь к .md (вместо md)"},
         "out": {**_PATH, "description": "путь итогового .docx; CLI-алиас: output"},
         "output": {**_PATH, "description": "CLI-алиас out"},
         "author": {"type": "string", "default": "colleague"}},
         required=["out"], write=True)},
    {"name": "docx_doctor", "permission": "read",
     "description": "Проверить окружение (python, python-docx, lxml, pydantic, "
                    "запись cwd). Только чтение.",
     "schema": _schema({
         "verify_signature": {"type": "boolean", "default": False},
         "dist": {"type": "string", "default": "dist"}})},
    {"name": "docx_agent_read", "permission": "read",
     "description": "Ограниченное JSON-чтение .docx (окно + курсор). Только "
                    "чтение.",
     "schema": _schema({
         "file": {**_PATH, "description": "путь к .docx; CLI-алиас: path"},
         "path": {**_PATH, "description": "CLI-алиас file"},
         "start": {"type": "integer", "default": 0},
         "max_chars": {"type": "integer", "default": 4000,
                       "description": "CLI-алиас max"},
         "max": {"type": "integer", "description": "CLI-алиас max_chars"}})},
    {"name": "docx_verify_forbidden", "permission": "read",
     "description": "Проверить .docx (+опц. .md) на запрещённые строки "
                    "(desync-guard). Только чтение.",
     "schema": _schema({
         "docx": {**_PATH, "description": "путь к .docx; CLI-алиас: path"},
         "path": {**_PATH, "description": "CLI-алиас docx"},
         "md": {**_PATH, "description": "опциональный .md к проверке"},
         "forbid": {"type": "array", "items": {"type": "string"},
                    "description": "запрещённые подстроки"},
         "forbid_file": {**_PATH, "description": "файл списка паттернов"}},
         required=["docx"])},
]

TOOLS_BY_NAME = {t["name"]: t for t in TOOLS}


def _pick(values: dict, names: tuple[str, ...], label: str):
    """Значение под любым из допустимых имён (CLI-алиас ↔ имя движка).
    Два имени с РАЗНЫМИ значениями — явная ошибка (паттерн sud)."""
    found = {n: values.pop(n) for n in names if n in values}
    if not found:
        return None
    distinct = {repr(v) for v in found.values()}
    if len(distinct) > 1:
        raise ValueError(f"conflicting values for {label}: {sorted(found)} — "
                         "передайте только одно имя")
    return next(iter(found.values()))


def _run_captured(func, ns: argparse.Namespace) -> dict:
    """Запуск cmd_* из cli.py с перехватом stdout-конверта: MCP не дублирует
    логику CLI, а переиспользует ровно тот же код и ровно тот же конверт."""
    from docx_cli import cli as _cli
    buf = io.StringIO()
    with redirect_stdout(buf):
        exit_code = func(ns)
    text = buf.getvalue().strip()
    payload = json.loads(text) if text else {}
    payload.setdefault("exit_code", exit_code)
    return payload


def _call(name: str, args: dict) -> dict:
    """Диспатч инструментов. Ошибки — исключения (handle_request конвертит
    в isError=True), сайд-эффекты только после confirm-гейта в _dispatch_gated."""
    values = dict(args or {})
    if name == "docx_build":
        from docx_cli.builder import build_document
        from docx_cli.spec import DocSpec
        out = _pick(values, ("out", "output"), "out/output")
        if out is None:
            raise ValueError("out is required (путь итогового .docx)")
        spec = DocSpec.from_dict(values["spec"])
        return build_document(spec, str(out))
    if name == "docx_validate":
        from docx_cli import cli as _cli
        file = _pick(values, ("file", "path"), "file/path")
        if file is None:
            raise ValueError("file is required")
        return _run_captured(_cli.cmd_validate, argparse.Namespace(
            file=str(file), json=True,
            strict=bool(values.get("strict", False)),
            mode=str(values.get("mode", "gost"))))
    if name == "docx_scan":
        from docx_cli import cli as _cli
        file = _pick(values, ("file", "path"), "file/path")
        if file is None:
            raise ValueError("file is required")
        return _run_captured(_cli.cmd_scan, argparse.Namespace(
            file=str(file), json=True,
            a2a_args=bool(values.get("a2a_args", False))))
    if name == "docx_normalize":
        from docx_cli.fonts_normalize import normalize_path
        path = _pick(values, ("path", "file"), "path/file")
        if path is None:
            raise ValueError("path is required")
        p = Path(str(path))
        if p.is_dir() and not bool(values.get("yes", False)):
            return {"ok": False, "status": "invalid_usage",
                    "reason": "directory mode requires yes=true"}
        res = normalize_path(str(p),
                             body_size=int(values.get("body_size", 12)),
                             header_size=int(values.get("header_size", 11)),
                             dry_run=bool(values.get("dry_run", False)),
                             backup=bool(values.get("backup", True)),
                             yes=bool(values.get("yes", False)))
        res.update({"ok": True, "status": "ok"})
        return res
    if name == "docx_md_build":
        from docx_cli.md_court import md_build
        out = _pick(values, ("out", "output"), "out/output")
        if out is None:
            raise ValueError("out is required")
        md_text = values.get("md")
        md_path = values.get("md_path")
        if md_text is not None and md_path is not None:
            raise ValueError("pass only one of md / md_path")
        if md_text is None and md_path is None:
            raise ValueError("md or md_path is required")
        if md_text is not None:
            tmp = Path(tempfile.mkdtemp(prefix="docx_mcp_md_")) / "input.md"
            tmp.write_text(str(md_text), encoding="utf-8")
            source = str(tmp)
        else:
            source = str(md_path)
        res = md_build(source, str(out), str(values.get("author", "colleague")))
        res.update({"ok": True, "status": "ok"})
        return res
    if name == "docx_doctor":
        from docx_cli import cli as _cli
        return _run_captured(_cli.cmd_doctor, argparse.Namespace(
            json=True,
            verify_signature=bool(values.get("verify_signature", False)),
            dist=str(values.get("dist", "dist"))))
    if name == "docx_agent_read":
        from docx_cli import cli as _cli
        file = _pick(values, ("file", "path"), "file/path")
        if file is None:
            raise ValueError("file is required")
        max_chars = _pick(values, ("max_chars", "max"), "max_chars/max")
        if max_chars is None:
            max_chars = 4000
        return _run_captured(_cli.cmd_agent_read, argparse.Namespace(
            file=str(file), start=int(values.get("start", 0)),
            max_chars=int(max_chars)))
    if name == "docx_verify_forbidden":
        from docx_cli import cli as _cli
        docx_path = _pick(values, ("docx", "path"), "docx/path")
        if docx_path is None:
            raise ValueError("docx is required")
        return _run_captured(_cli.cmd_verify_forbidden, argparse.Namespace(
            docx=str(docx_path), md=values.get("md"),
            forbid=list(values.get("forbid", [])),
            forbid_file=values.get("forbid_file"), json=True))
    raise KeyError(f"unknown tool: {name}")


def _check_arguments(tool: dict, args: dict) -> dict:
    """fail-closed валидация аргументов ДО движка: unknown fields и
    permission-mismatch не смещаются confirm-гейтом."""
    if not isinstance(args, dict):
        raise TypeError("arguments must be an object")
    values = dict(args)
    schema = tool["schema"]
    allowed = set((schema.get("properties") or {}).keys())
    unknown = set(values) - allowed
    if unknown:
        raise ValueError(f"unknown fields: {sorted(unknown)}")
    perm = values.pop("permission", None)
    if perm is not None and perm != tool["permission"]:
        raise ValueError(
            f"permission_mismatch: tool {tool['name']} declares "
            f"'{tool['permission']}', got '{perm}'")
    for field in schema.get("required", []):
        if field not in values:
            raise ValueError(f"missing required field: {field}")
    confirm = None
    if tool["permission"] == "write":
        confirm = values.pop("confirm", None)
    return {"values": values, "confirm": confirm}


def _dispatch_gated(name: str, args: dict) -> Any:
    tool = TOOLS_BY_NAME.get(name)
    if tool is None:
        raise KeyError(f"unknown tool: {name}")
    prepared = _check_arguments(tool, args)
    if tool["permission"] == "write" and prepared["confirm"] is not True:
        # needs_confirmation — ДАННЫЕ, не ошибка транспорта; запись не
        # производилась ни байта (сайд-эффекты только в _call ниже).
        return {"status": "needs_confirmation", "permission": "write",
                "tool": name,
                "detail": "write-инструмент требует явный confirm: true — "
                          "запись не произведена"}
    return _call(name, prepared["values"])


def tool_list() -> dict:
    return {"tools": [{"name": t["name"], "description": t["description"],
                       "inputSchema": t["schema"], "permission": t["permission"]}
                      for t in TOOLS]}


def handle_request(request: dict) -> dict | None:
    method = request.get("method")
    request_id = request.get("id")
    if method == "initialize":
        result = {"protocolVersion": "2024-11-05",
                  "capabilities": {"tools": {}},
                  "serverInfo": {"name": "docx-cli", "version": _PKG_VERSION}}
    elif method == "notifications/initialized":
        return None
    elif method == "ping":
        result = {}
    elif method == "tools/list":
        result = tool_list()
    elif method == "tools/call":
        try:
            params = request.get("params", {})
            value = _dispatch_gated(params["name"], params.get("arguments", {}))
            if hasattr(value, "to_dict"):
                value = value.to_dict()
            result = {"content": [{"type": "text",
                                   "text": json.dumps(value, ensure_ascii=False,
                                                      default=str)}],
                      "isError": False}
        except (KeyError, TypeError, ValueError, OSError) as exc:
            result = {"content": [{"type": "text", "text": str(exc)}],
                      "isError": True}
        except Exception as exc:  # доменные ошибки: stdio не умирает
            result = {"content": [{"type": "text",
                                   "text": f"unexpected: {type(exc).__name__}: {exc}"}],
                      "isError": True}
    else:
        return {"jsonrpc": "2.0", "id": request_id,
                "error": {"code": -32601, "message": "Method not found"}}
    return {"jsonrpc": "2.0", "id": request_id, "result": result}


def serve_stdio(input_stream=None, output_stream=None) -> None:
    input_stream = input_stream or sys.stdin
    output_stream = output_stream or sys.stdout
    for line in input_stream:
        if not line.strip():
            continue
        try:
            frame = json.loads(line)
        except json.JSONDecodeError as exc:
            response = {"jsonrpc": "2.0", "id": None,
                        "error": {"code": -32700, "message": f"Parse error: {exc}"}}
        else:
            if not isinstance(frame, dict):
                response = {"jsonrpc": "2.0", "id": None,
                            "error": {"code": -32600,
                                      "message": "Invalid Request: not an object"}}
            else:
                response = handle_request(frame)
        if response is not None:
            output_stream.write(json.dumps(response, ensure_ascii=False) + "\n")
            output_stream.flush()


def run_tools_dump() -> int:
    print(json.dumps(tool_list(), ensure_ascii=False))
    return 0


def run_stdio_server() -> int:
    serve_stdio()
    return 0
