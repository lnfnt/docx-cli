# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S2-C1: слово «проект/черновик» в теле не должно переводить документ в шапку.

Раньше любое вхождение подстроки (напр. «сумма проекта составляет...») во
второй половине текста возвращало in_header=True, и весь последующий текст
уходил в реквизитный блок с отступом 8.1 см — без пути назад.
"""
from __future__ import annotations


def _indents(doc):
    return [round(p.paragraph_format.left_indent.inches, 1)
            if p.paragraph_format.left_indent is not None else 0.0
            for p in doc.paragraphs]


def test_body_word_project_does_not_flip_to_header():
    from docx_cli.md_court import build_md_document
    md = ("# Исковое заявление\n"
          "Истец: Иванов\n"
          "Ответчик: Петров\n"
          "Сумма проекта составляет сто тысяч рублей.\n"
          "Прошу взыскать указанную сумму.\n")
    doc = build_md_document(md, "tester")
    indents = _indents(doc)
    assert 3.2 not in indents, f"тело ушло в шапку: {indents}"
    texts = [p.text for p in doc.paragraphs]
    assert any("Сумма проекта" in t for t in texts)


def test_standalone_stamp_still_centered():
    from docx_cli.md_court import build_md_document
    md = "ПРОЕКТ\n\n# ИСКОВОЕ ЗАЯВЛЕНИЕ\nТело документа."
    doc = build_md_document(md, "tester")
    assert doc.paragraphs[0].text.strip().upper() == "ПРОЕКТ"
    indents = _indents(doc)
    assert 3.2 not in indents, f"тело ушло в шапку: {indents}"
