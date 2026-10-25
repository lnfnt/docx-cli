"""docx-cli v0.7.0 §2 B-07 vectors (audit F6)."""
from __future__ import annotations

from docx import Document

from docx_cli.validate import _check_numbered_list_sequence


def test_v070_f6_b1_two_sequential_numbered_lists_ok():
    """B1: два последовательных перечня легитимны."""
    doc = Document()
    for t in ("1. первый", "2. второй", "1. третий", "2. четвёртый"):
        doc.add_paragraph(t)
    assert _check_numbered_list_sequence(doc, None) == []


def test_v070_f6_b2_proshu_and_dopolnitelno_reset():
    """B2: «Прошу:» и «Дополнительно:» сбрасывают счётчик."""
    doc = Document()
    for t in ("Прошу:", "1. взыскать", "Дополнительно:", "1. приобщить"):
        doc.add_paragraph(t)
    assert _check_numbered_list_sequence(doc, None) == []


def test_v070_f6_b3_disorder_inside_block_fails():
    """B3: беспорядок внутри блока даёт ошибку."""
    doc = Document()
    for t in ("1. a", "3. c", "2. b"):
        doc.add_paragraph(t)
    errs = _check_numbered_list_sequence(doc, None)
    assert errs and any("numbered_list_sequence" in e for e in errs)


def test_v070_f6_probe3_six():
    """probe3: 6 пунктов в двух блоках — без ложного срабатывания."""
    doc = Document()
    for t in ("1. a", "2. b", "3. c", "Прошу:", "1. d", "2. e"):
        doc.add_paragraph(t)
    assert _check_numbered_list_sequence(doc, None) == []