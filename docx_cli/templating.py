# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Пакет 5 (TASK_PACKAGE5_TEMPLATING.md): движок шаблонизации docx.

Безопасная замена {{var}} сквозь границы дробленых run'ов (run-reconciliation)
с сохранением форматирования первого run'а, вложенные ключи ({{client.name}}),
табличные циклы ({{#items}} ... {{/items}}), условные блоки
({{#if x}} ... {{/if}}) и обработка колонтитулов/сносок.
"""
from __future__ import annotations

import copy
import io
import json
import os
import re
import zipfile
from typing import Any, Callable

from docx import Document
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .errors import DocxCreatorError, TemplateError

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

VAR_RE = re.compile(r"\{\{([^{}]+)\}\}")
IF_OPEN_RE = re.compile(r"\{\{#if\s+([^{}]+?)\s*\}\}")
ENDIF_RE = re.compile(r"\{\{/if\}\}")
IF_BLOCK_RE = re.compile(r"\{\{#if\s+([^{}]+?)\s*\}\}(.*?)\{\{/if\}\}", re.DOTALL)
LOOP_OPEN_RE = re.compile(r"\{\{#([A-Za-z0-9_.]+)\}\}")

_MISSING = object()
_MAX_TABLE_PASSES = 16  # ponytail: защита от вечного пересканирования вложенных циклов


class _Seg:
    """Кусок редактируемого текста (run в python-docx или w:r в XML) с сеттером."""

    __slots__ = ("text", "_write")

    def __init__(self, text: str, write: Callable[[str], None]):
        self.text = text
        self._write = write

    def write(self) -> None:
        self._write(self.text)


# ------------------------------------------------------------------------------
# Ядро: run-reconciliation
# ------------------------------------------------------------------------------

def _layout(segs: list[_Seg]) -> tuple[list[int], list[int], int]:
    starts: list[int] = []
    sizes: list[int] = []
    pos = 0
    for s in segs:
        starts.append(pos)
        sizes.append(len(s.text))
        pos += len(s.text)
    return starts, sizes, pos


def _apply(segs: list[_Seg], spans: list[tuple[int, int, str]]) -> None:
    """Заменить непересекающиеся спаны [start, end) -> repl в конкатенации
    сегментов, сохранив текст вне спанов. Форматирование смещается на
    первый (левый) сегмент спана, остатки плейсхолдера в хвостовых
    сегментах вычищаются."""
    starts, sizes, total = _layout(segs)
    for s, e, repl in sorted(spans, key=lambda x: x[0], reverse=True):
        for seg, st, ln in zip(segs, starts, sizes):
            if st + ln <= s:
                continue
            if st >= e:
                break
            ls = max(s, st) - st
            le = min(e, st + ln) - st
            if ls == le:
                continue
            ins = repl if st <= s < st + ln else ""
            seg.text = seg.text[:ls] + ins + seg.text[le:]
    for seg in segs:
        seg.write()


def _para_runs(p: Paragraph):
    """S3-H4: runs абзаца + runs внутри w:hyperlink — python-docx p.runs
    гиперссылки не видит, и плейсхолдеры там молча оставались
    незамеченными при status ok."""
    runs = list(p.runs)
    seen = {id(r) for r in runs}
    for hl in p._p.iter(qn("w:hyperlink")):
        for r in hl.iter(qn("w:r")):
            if id(r) not in seen:
                seen.add(id(r))
                runs.append(Run(r, p))
    return runs


def _para_segs(p: Paragraph) -> list[_Seg]:
    def _writer(r):
        def _w(v, _r=r):
            _r.text = v
        return _w
    return [_Seg(r.text, _writer(r)) for r in _para_runs(p)]


# S3-L: _r_elm_writer/_r_elm_segs — мёртвый дубль _para_segs/_writer (ни одного
# вызова вне определения); удалены в волне LOW-финала 02.11.


# ------------------------------------------------------------------------------
# Данные: точечная нотация
# ------------------------------------------------------------------------------

def _fmt(v: Any) -> str:
    if isinstance(v, bool):
        return "true" if v else "false"
    if v is None:
        return ""
    if isinstance(v, (dict, list)):
        return json.dumps(v, ensure_ascii=False)
    return str(v)


def _lookup(data: dict, key: str) -> Any:
    # {{items[].name}} внутри строк цикла = поле текущего элемента
    # (контекст элемента уже слит наверх) — срезаем селектор массива.
    key = key.split("[].")[-1]
    cur: Any = data
    for part in key.split("."):
        part = part.strip()
        if isinstance(cur, dict):
            if part in cur:
                cur = cur[part]
            else:
                return _MISSING
        elif isinstance(cur, list) and part.lstrip("-").isdigit():
            i = int(part)
            if -len(cur) <= i < len(cur):
                cur = cur[i]
            else:
                return _MISSING
        else:
            return _MISSING
    return cur


def _truthy(v: Any) -> bool:
    if v is _MISSING or v is None:
        return False
    return bool(v)


def _record_missing(stats: dict, key: str, strict: bool) -> None:
    stats["missing"].append(key)
    if strict:
        raise TemplateError(f"missing variable '{key}'")


# ------------------------------------------------------------------------------
# Абзацы: условные блоки + переменные
# ------------------------------------------------------------------------------

def _render_paragraph(p: Paragraph, data: dict, strict: bool, stats: dict) -> None:
    all_runs = _para_runs(p)
    if not all_runs:
        return
    full = "".join(r.text for r in all_runs)
    if "{{" not in full:
        return

    # Фаза 1: инлайновые {{#if x}}...{{/if}} в одном абзаце.
    spans = []
    for m in IF_BLOCK_RE.finditer(full):
        key = m.group(1).strip()
        v = _lookup(data, key)
        if v is _MISSING:
            _record_missing(stats, key, strict)
            # S3-L: в strict — fail-closed (блок убираем, как раньше);
            # в non-strict — симметрия с {{var}}: плейсхолдер остаётся.
            if strict:
                spans.append((m.start(), m.end(), ""))  # fail-closed: блок убираем
        elif _truthy(v):
            spans.append((m.start(), m.end(), m.group(2)))
        else:
            spans.append((m.start(), m.end(), ""))
    if spans:
        _apply(_para_segs(p), spans)
        full = "".join(r.text for r in all_runs)

    # Фаза 2: переменные. Отсутствующие в non-strict остаются как есть.
    spans = []
    for m in VAR_RE.finditer(full):
        key = m.group(1).strip()
        v = _lookup(data, key)
        if v is _MISSING:
            _record_missing(stats, key, strict)
            continue
        spans.append((m.start(), m.end(), _fmt(v)))
        stats["replaced"][key] = stats["replaced"].get(key, 0) + 1
    if spans:
        _apply(_para_segs(p), spans)


def _strip_spans(p: Paragraph, cuts: list[tuple[int, int]]) -> None:
    _apply(_para_segs(p), [(s, e, "") for s, e in cuts])


def _process_conditionals(paras: list[Paragraph], data: dict, strict: bool,
                          stats: dict) -> None:
    """Межабзацные {{#if}} ... {{/if}}: truthy — снять маркеры; falsy —
    удалить содержимое блока (хвосты до/после маркеров сохраняются)."""
    i = 0
    while i < len(paras):
        txt = paras[i].text
        m = IF_OPEN_RE.search(txt)
        if not m or ENDIF_RE.search(txt):  # инлайновый — обработается в _render_paragraph
            i += 1
            continue
        key = m.group(1).strip()
        j = None
        for k in range(i + 1, len(paras)):
            if ENDIF_RE.search(paras[k].text):
                j = k
                break
        if j is None:
            if strict:
                raise TemplateError(f"unclosed {{#if {key}}} block")
            i += 1
            continue
        v = _lookup(data, key)
        if v is _MISSING:
            _record_missing(stats, key, strict)
            # S3-L: симметрия с {{var}} — в non-strict блок {{#if}} остаётся
            # как есть (маркеры не трогаем), а не удаляется молча.
            if not strict:
                i += 1
                continue
        if _truthy(v):
            close = ENDIF_RE.search(paras[j].text)
            _strip_spans(paras[i], [(m.start(), m.end())])
            _strip_spans(paras[j], [(close.start(), close.end())])
        else:
            close = ENDIF_RE.search(paras[j].text)
            _strip_spans(paras[i], [(m.start(), len(paras[i].text))])
            _strip_spans(paras[j], [(0, close.end())])
            for k in range(i + 1, j):
                _strip_spans(paras[k], [(0, len(paras[k].text))])
            for k in range(i, j + 1):
                if not paras[k].text.strip():
                    el = paras[k]._p
                    el.getparent().remove(el)
        i = j + 1


# ------------------------------------------------------------------------------
# Таблицы: циклы {{#name}} ... {{/name}}
# ------------------------------------------------------------------------------

def _unique_cells(table):
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            if cell._tc in seen:
                continue
            seen.add(cell._tc)
            yield cell


def _iter_unique_tables(initial: list) -> list:
    seen = set()
    stack = list(initial)
    out = []
    while stack:
        t = stack.pop()
        if t._tbl in seen:
            continue
        seen.add(t._tbl)
        out.append(t)
        for cell in _unique_cells(t):
            stack.extend(cell.tables)
    return out


def _strip_loop_tokens(p: Paragraph, name: str) -> None:
    runs = p.runs
    if not runs:
        return
    full = "".join(r.text for r in runs)
    # ponytail: вложенные циклы не поддерживаются — в клонах снимаются ВСЕ
    # #маркеры, чтобы повторное сканирование таблицы не плодило клоны клона.
    pat = re.compile(
        r"\{\{#\s*[A-Za-z0-9_.]+\s*\}\}|\{\{/\s*[A-Za-z0-9_.]+\s*\}\}"
    )
    spans = [(m.start(), m.end()) for m in pat.finditer(full)]
    if not spans:
        return
    _strip_spans(p, spans)


def _table_loop_pass(table, data: dict, strict: bool, stats: dict) -> bool:
    changed = False
    rows = list(table.rows)
    for i, row in enumerate(rows):
        row_text = "".join(c.text for c in row.cells)
        m = LOOP_OPEN_RE.search(row_text)
        if not m or m.group(1) == "if":
            continue
        name = m.group(1)
        close_tok = "{{/" + name + "}}"
        j = i if close_tok in row_text else None
        if j is None:
            for k in range(i + 1, len(rows)):
                if close_tok in "".join(c.text for c in rows[k].cells):
                    j = k
                    break
        if j is None:
            continue  # незакрытый цикл — оставляем как есть
        items = _lookup(data, name)
        if items is _MISSING:
            _record_missing(stats, name, strict)
            continue
        if not isinstance(items, list):
            if strict:
                raise TemplateError(f"variable '{name}' is not a list")
            continue
        block = [rows[k]._tr for k in range(i, j + 1)]
        anchor = block[-1]
        for item in items:
            if isinstance(item, dict):
                ctx = {**data, **item, "item": item}
            else:
                ctx = {**data, name: item, "item": item}
            for tr in block:
                clone = copy.deepcopy(tr)
                anchor.addnext(clone)
                anchor = clone
                for p_elm in clone.iter(qn("w:p")):
                    p = Paragraph(p_elm, None)
                    _strip_loop_tokens(p, name)
                    _render_paragraph(p, ctx, strict, stats)
        for tr in block:
            tr.getparent().remove(tr)
        stats["rows_cloned"] += len(items)
        changed = True
    return changed


def _table_loops(table, data: dict, strict: bool, stats: dict) -> None:
    for _ in range(_MAX_TABLE_PASSES):
        if not _table_loop_pass(table, data, strict, stats):
            break


# ------------------------------------------------------------------------------
# Колонтитулы, сноски, обход контейнеров
# ------------------------------------------------------------------------------

def _hf_parts(doc):
    for section in doc.sections:
        for part in (section.header, section.footer,
                     section.first_page_header, section.first_page_footer,
                     section.even_page_header, section.even_page_footer):
            if part is None or part.is_linked_to_previous:
                continue
            yield part


def _patch_notes(zip_path: str, data: dict, strict: bool, stats: dict) -> None:
    """Сноски/концевые сноски: python-docx их не раскрывает — правим XML
    внутри пакета (run-reconciliation на уровне w:r)."""
    import lxml.etree as etree

    with zipfile.ZipFile(zip_path) as zf:
        items = [(i, zf.read(i.filename)) for i in zf.infolist()]

    def _write_elm(v, _r):
        for t in _r.findall(qn("w:t")):
            _r.remove(t)
        if v:
            t = _r.makeelement(qn("w:t"), {})
            t.set(XML_SPACE, "preserve")
            t.text = v
            _r.append(t)

    changed = False
    for name in ("word/footnotes.xml", "word/endnotes.xml"):
        raw = next((b for i, b in items if i.filename == name), None)
        if raw is None:
            continue
        root = etree.fromstring(raw)
        touched = False
        for p_elm in root.iter(W + "p"):
            segs = []
            for r in p_elm.iter(W + "r"):
                text = "".join(t.text or "" for t in r.findall(qn("w:t")))
                segs.append(_Seg(text, (lambda rr=r: lambda v: _write_elm(v, rr))()))
            if not segs:
                continue
            full = "".join(s.text for s in segs)
            if "{{#if" in full:
                spans = []
                for m in IF_BLOCK_RE.finditer(full):
                    key = m.group(1).strip()
                    v = _lookup(data, key)
                    if v is _MISSING:
                        _record_missing(stats, key, strict)
                        # S3-L: асимметрия non-strict — {{var}} оставляет
                        # плейсхолдер, {{#if}} молча удалял блок. Не strict:
                        # пропускаем спан, блок остаётся как есть.
                        continue
                    spans.append((m.start(), m.end(),
                                  m.group(2) if _truthy(v) else ""))
                if spans:
                    _apply(segs, spans)
                    # S3-M4: if-only правки применялись в памяти и молча
                    # отбрасывались — touched не выставлялся для {{#if}}-ветки.
                    touched = True
                segs = []
                for r in p_elm.iter(W + "r"):
                    text = "".join(t.text or "" for t in r.findall(qn("w:t")))
                    segs.append(_Seg(text, (lambda rr=r: lambda v: _write_elm(v, rr))()))
                full = "".join(s.text for s in segs)
            spans = []
            for m in VAR_RE.finditer(full):
                key = m.group(1).strip()
                v = _lookup(data, key)
                if v is _MISSING:
                    _record_missing(stats, key, strict)
                    continue
                spans.append((m.start(), m.end(), _fmt(v)))
                stats["replaced"][key] = stats["replaced"].get(key, 0) + 1
            if spans:
                _apply(segs, spans)
                touched = True
        if touched:
            for idx, (info, _) in enumerate(items):
                if info.filename == name:
                    items[idx] = (info, etree.tostring(
                        root, xml_declaration=True, encoding="UTF-8", standalone=True))
            changed = True
    if not changed:
        return
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zout:
        for info, blob in items:
            zout.writestr(info, blob)
    tmp = zip_path + ".tmp"
    with open(tmp, "wb") as f:
        f.write(buf.getvalue())
    os.replace(tmp, zip_path)


# ------------------------------------------------------------------------------
# Публичный API
# ------------------------------------------------------------------------------

def render_template(template_path: str, data: dict, output_path: str,
                    strict: bool = False) -> dict:
    """Заполнить .docx-шаблон данными. Возвращает конверт со счётчиками.

    strict=True — fail-closed: отсутствующая переменная -> TemplateError.
    strict=False — плейсхолдеры остаются как есть, отчёт в "missing"."""
    if not isinstance(data, dict):
        raise TemplateError("data must be a JSON object (dict)")

    # S3-L: выход по месту шаблона уничтожил бы и источник, и результат —
    # fail-closed до любых изменений.
    if str(template_path) == str(output_path):
        raise TemplateError("template and output must differ (refusing to overwrite input)")

    doc = Document(str(template_path))
    stats: dict = {"replaced": {}, "missing": [], "rows_cloned": 0}

    tables = _iter_unique_tables(doc.tables)
    for t in tables:
        _table_loops(t, data, strict, stats)

    # дедуп элементов абзацев: колонтитулы повторяются по секциям
    seen = set()
    groups: list[list[Paragraph]] = []

    def push(paras) -> None:
        fresh = []
        for para in paras:
            if para._p in seen:
                continue
            seen.add(para._p)
            fresh.append(para)
        if fresh:
            groups.append(fresh)

    hf_tables: list = []
    for part in _hf_parts(doc):
        hf_tables.extend(_iter_unique_tables(part.tables))
    for t in hf_tables:
        _table_loops(t, data, strict, stats)

    push(doc.paragraphs)
    all_tables = list(tables) + hf_tables
    for t in all_tables:
        for cell in _unique_cells(t):
            push(cell.paragraphs)
    for part in _hf_parts(doc):
        push(part.paragraphs)
        for t in _iter_unique_tables(part.tables):
            for cell in _unique_cells(t):
                push(cell.paragraphs)

    for paras in groups:
        _process_conditionals(paras, data, strict, stats)
        for p in paras:
            _render_paragraph(p, data, strict, stats)

    # S3-H3: атомарность — пишем во временный файл и переносим по успеху.
    # Падение _patch_notes (или обрыв записи) не оставляет полуфайл и не
    # портит прежний результат по месту output_path.
    import os
    tmp_out = f"{output_path}.tmp-{os.getpid()}"
    try:
        doc.save(tmp_out)
        _patch_notes(tmp_out, data, strict, stats)
        os.replace(tmp_out, str(output_path))
    except Exception:
        try:
            os.unlink(tmp_out)
        except OSError:
            pass
        raise

    missing = sorted(set(stats["missing"]))
    return {
        "ok": True,
        "status": "ok" if not missing else "degraded",
        "reason": None if not missing else "missing variables: " + ", ".join(missing),
        "template": str(template_path),
        "output": str(output_path),
        "replaced": stats["replaced"],
        "missing": missing,
        "rows_cloned": stats["rows_cloned"],
    }


__all__ = ["TemplateError", "DocxCreatorError", "render_template"]
