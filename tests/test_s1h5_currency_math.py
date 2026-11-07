# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S1-H5: денежные суммы в тексте не должны становиться OMML-формулой.

Защита «закрывающий $ не после пробела» не спасает: «($100) и ($200)»
превращалось в формулу «100) и (». Правило: чисто числовое выражение
(деньги) либо продолжение после закрывающего $ цифрой/скобкой — литерал.
"""
from __future__ import annotations

from docx_cli.blocks import _split_inline


def _math_spans(text: str):
    return [s.text for s in _split_inline(text) if s.math]


def test_currency_parens_not_math():
    assert _math_spans("($100) и ($200)") == []


def test_plain_currency_not_math():
    assert _math_spans("сумма $100 000 рублей") == []
    assert _math_spans("от $100 до $500") == []


def test_real_inline_math_still_works():
    spans = _math_spans("при $x$ меньше нуля")
    assert spans == ["x"]
    spans = _math_spans("формула $E=mc^2$ верна")
    assert spans == ["E=mc^2"]
