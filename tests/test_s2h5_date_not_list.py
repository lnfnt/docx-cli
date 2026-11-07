# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-H5: строка-дата («1. января 2020 года…») не должна получать форматирование
списка (висячий отступ, отрицательный first_line_indent).

Реальные пункты списка по-прежнему оформляются как список.
"""
from __future__ import annotations


def _fli(p):
    v = p.paragraph_format.first_line_indent
    return v.cm if v is not None else None


def test_date_line_no_hanging_indent():
    from docx_cli.md_court import build_md_document
    d = build_md_document(
        "# ИСКОВОЕ ЗАЯВЛЕНИЕ\n1. января 2020 года заключён договор.", "tester")
    date_p = [p for p in d.paragraphs if "января" in p.text][0]
    fli = _fli(date_p)
    assert fli is None or fli >= 0, f"дата получила висячий отступ списка: {fli}"


def test_real_list_item_formatted():
    from docx_cli.md_court import build_md_document
    d = build_md_document(
        "# ИСКОВОЕ ЗАЯВЛЕНИЕ\nОбстоятельства.\n1. Первое обстоятельство.\n2. Второе.", "tester")
    item_ps = [p for p in d.paragraphs if "обстоятельство" in p.text.lower()
               and "первое" in p.text.lower()]
    assert item_ps, "пункт списка не найден"
    p = item_ps[0]
    fli = _fli(p)
    li = p.paragraph_format.left_indent
    formatted = (fli is not None and fli < 0) or \
                (li is not None and abs(li.cm - 1.25) < 0.01)
    assert formatted, f"настоящий пункт списка не оформлен: fli={fli} left={li.cm if li else None}"
