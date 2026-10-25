"""docx-cli agent-read — ограниченное JSON-чтение docx для агентов.

v0.5.0 (clean-room: собственный agent-read для ИИ-агентов). Контракт: блоки
(параграф/таблица) с id, полными char-offset в конкатенированном тексте,
флагом complete и next_cursor — агент читает большие документы по окнам.

Fail-closed: отсутствующий файл -> ошибка-конверт.
"""
from __future__ import annotations

from pathlib import Path

from docx import Document

DEFAULT_MAX_CHARS = 4000


class AgentReadError(Exception):
    pass


def _blocks(doc) -> list[dict]:
    """Блоки документа по порядку: параграфы и таблицы (строки->ячейки)."""
    out: list[dict] = []
    # python-docx не даёт единый поток тело->таблица по порядку просто; идём
    # по XML тела в порядке следования.
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    for child in body.iterchildren():
        tag = child.tag.rsplit('}', 1)[-1]
        if tag == 'p':
            p = Paragraph(child, doc)
            if p.text.strip():
                out.append({"kind": "paragraph", "text": p.text})
        elif tag == 'tbl':
            t = Table(child, doc)
            for row in t.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    out.append({"kind": "table_row", "text": " | ".join(cells)})
    return out


def read(path: str | Path, start_char: int = 0, max_chars: int = DEFAULT_MAX_CHARS) -> dict:
    if isinstance(start_char, bool) or isinstance(max_chars, bool) or \
       not isinstance(start_char, int) or not isinstance(max_chars, int):
        raise AgentReadError("invalid_usage: max and cursor must be integers")
    if start_char < 0 or max_chars <= 0:
        raise AgentReadError("invalid_usage: max must be >0 and cursor must be >=0")
    p = Path(path)
    if not p.is_file():
        raise AgentReadError(f"file not found: {p}")
    doc = Document(str(p))
    blocks = _blocks(doc)
    # полный текст = конкатенация с \n; каждому блоку — диапазон offsets
    enriched = []
    cursor = 0
    for i, b in enumerate(blocks):
        start = cursor
        text = b["text"]
        cursor += len(text) + 1  # +1 за \n
        enriched.append({"id": f"b{i}", "kind": b["kind"], "text": text,
                         "char_start": start, "char_end": cursor - 1})
    total = cursor
    if start_char > total or (total > 0 and start_char >= total):
        raise AgentReadError(
            f"invalid_usage: cursor {start_char} is beyond end (total_chars={total}); "
            "use a cursor returned by a previous call or 0")
    window_start = start_char
    window_end = min(window_start + max_chars, total)

    result = []
    for b in enriched:
        if b["char_end"] <= window_start or b["char_start"] >= window_end:
            continue
        text = b["text"]
        s_off = max(window_start - b["char_start"], 0)
        e_off = min(window_end - b["char_start"], len(text))
        frag = text[s_off:e_off]
        result.append({"id": b["id"], "kind": b["kind"], "text": frag,
                       "char_start": b["char_start"] + s_off,
                       "complete": s_off == 0 and e_off == len(text)})

    next_cursor = window_end if window_end < total else None
    return {"total_chars": total, "window": [window_start, window_end],
            "blocks": result, "next_cursor": next_cursor}
