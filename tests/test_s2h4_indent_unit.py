# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-H4: единичный мелкий отступ не должен масштабировать глубину списка.

unit = min(отступ): один пробел-опечатка делал unit=1, реальные вложенные
пункты улетали на глубину >4 и удалялись (контент потерян).
"""
from __future__ import annotations


def _all_text(doc) -> str:
    return "\n".join(p.text for p in doc.paragraphs)


def test_tiny_indent_typo_does_not_drop_items():
    from docx_cli.md_court import build_md_document
    md = ("# ИСКОВОЕ ЗАЯВЛЕНИЕ\n"
          "Обоснование.\n"
          "1. Первый пункт.\n"
          " 2. Опечатка с одним пробелом.\n"
          "    3. Реально вложенный пункт.\n"
          "        4. Глубже.\n"
          "            5. Ещё глубже.\n"
          "                6. Максимум.\n")
    report: list = []
    d = build_md_document(md, "tester", report=report)
    text = _all_text(d)
    # контент пунктов не потерян
    assert "Первый пункт" in text
    assert "Опечатка с одним пробелом" in text
    assert "Реально вложенный пункт" in text
