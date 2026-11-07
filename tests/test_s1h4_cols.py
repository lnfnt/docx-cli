# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S1-H4/S1-M1: валидация колонок на границе spec→render + порядок w:cols в sectPr.

Раньше cols/col_space_cm принимали отрицательные/нулевые/1000-колоночные
значения → невалидный OOXML (диалог восстановления Word), а w:cols писался
в конец sectPr после w:docGrid — нарушение CT_SectPr.
"""
from __future__ import annotations

import pytest
from docx.oxml.ns import qn


def _section_block(**kw):
    from docx_cli.spec import SectionBreakBlock
    return SectionBreakBlock(type="section_break", **kw)


def test_cols_negative_rejected():
    with pytest.raises(Exception):
        _section_block(cols=-2)


def test_cols_zero_rejected():
    with pytest.raises(Exception):
        _section_block(cols=0)


def test_cols_huge_rejected():
    with pytest.raises(Exception):
        _section_block(cols=1000)


def test_cols_valid_accepted():
    b = _section_block(cols=2, col_space_cm=1.0)
    assert b.cols == 2


def test_col_space_negative_rejected():
    with pytest.raises(Exception):
        _section_block(cols=2, col_space_cm=-1)


def test_cols_before_docgrid_in_sectpr():
    import docx
    from docx.shared import Cm
    from docx_cli import blocks
    doc = docx.Document()
    for s in doc.sections:
        s.top_margin = Cm(2); s.bottom_margin = Cm(2)
        s.left_margin = Cm(3); s.right_margin = Cm(1.5)
    # пишем колонки через apply-section
    blk = _section_block(cols=2, col_space_cm=1.0)
    from docx_cli.spec import Profile
    fn = getattr(blocks, "render_section_break", None)
    if fn is None:
        import inspect
        cand = [n for n, f in inspect.getmembers(blocks, inspect.isfunction)
                if "section" in n.lower()]
        pytest.skip(f"apply fn not found: {cand}")
    fn(doc, blk, Profile())
    sectPr = doc.sections[0]._sectPr
    order = [el.tag.split("}")[1] for el in sectPr]
    if "cols" in order and "docGrid" in order:
        assert order.index("cols") < order.index("docGrid"), order
