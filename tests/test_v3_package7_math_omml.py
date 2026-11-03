# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

r"""ПАКЕТ 7: Математический процессор (Office Math / OMML — Домен 12).

Тесты по TASK_PACKAGE7_MATH_OMML.md:
  1. Дроби и корни:
     - \frac{a + b}{\sqrt{c}}
     - Наличие <m:f>, <m:num>, <m:den>, <m:rad> в word/document.xml.
  2. Индексы и суммы:
     - \sum_{i=1}^n x_i^2
     - Наличие <m:nary>, <m:sSubSup> (или <m:sSub>/<m:sSup>).
  3. Инлайн-формула в Markdown:
     - "Формула кинетической энергии: $E_k = \frac{m v^2}{2}$ в физике."
     - Внутри <w:p> появляется <m:oMath>.
  4. Выключная блочная формула:
     - Блок $$ ... $$ рендерится как <m:oMathPara> с центрированием.
  5. Греческие буквы и операторы:
     - \alpha, \beta, \gamma, \times, \pm, \le, \infty
     - Символы Unicode в <m:r><m:t>.
  6. Скобки:
     - \left( ... \right) и ( ... ) -> <m:d><m:e>...</m:e></m:d>.
  7. Сложная формула распределения:
     - $$ f(x) = \frac{1}{\sqrt{2\pi\sigma^2}} e^{-\frac{(x-\mu)^2}{2\sigma^2}} $$
  8. Инлайн в ячейках таблицы.
"""
from __future__ import annotations

import zipfile
import docx
import pytest

from docx_cli.builder import build_document
from docx_cli.md_engine import md_build_v2, MdEngineError
from docx_cli.omml import (
    latex_to_omml,
    latex_to_omml_element,
    add_math_paragraph,
    add_math_to_paragraph,
)
from docx_cli.spec import DocSpec, MathBlock

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build(tmp_workspace, body, profile="corporate") -> str:
    spec = DocSpec.from_dict({
        "profile": profile,
        "output": {"file_name": "math_doc.docx", "author": "MathAgent"},
        "header": {"align": "right", "lines": ["Подготовлено: MathAgent"]},
        "title": {"main": "Тест OMML"},
        "body": body,
    })
    out = str(tmp_workspace / "math_doc.docx")
    envelope = build_document(spec, out)
    assert envelope["ok"], envelope.get("reason")
    return out


def _document_xml(docx_path: str) -> str:
    with zipfile.ZipFile(docx_path) as z:
        return z.read("word/document.xml").decode("utf-8")


def test_omml_frac_and_rad():
    r"""Тест 1: Дроби и корни (\frac{a + b}{\sqrt{c}})."""
    xml = latex_to_omml(r"\frac{a + b}{\sqrt{c}}")
    assert "<m:f>" in xml
    assert "<m:num>" in xml
    assert "<m:den>" in xml
    assert "<m:rad>" in xml
    assert '<m:degHide m:val="1"/>' in xml

    # Проверка n-го корня \sqrt[3]{x}
    xml_deg = latex_to_omml(r"\sqrt[3]{x}")
    assert "<m:rad>" in xml_deg
    assert '<m:degHide m:val="0"/>' in xml_deg
    assert "<m:deg>" in xml_deg


def test_omml_indices_and_sum():
    r"""Тест 2: Индексы и суммы (\sum_{i=1}^n x_i^2)."""
    xml = latex_to_omml(r"\sum_{i=1}^n x_i^2")
    assert "<m:nary>" in xml
    assert '<m:chr m:val="∑"/>' in xml
    assert "<m:sSubSup>" in xml or ("<m:sSub>" in xml and "<m:sSup>" in xml)


def test_omml_greek_and_operators():
    """Тест греческих букв и спецсимволов."""
    xml = latex_to_omml(r"\alpha + \beta = \gamma \times \pi \pm \sigma \le \infty")
    assert "α" in xml
    assert "β" in xml
    assert "γ" in xml
    assert "×" in xml
    assert "π" in xml
    assert "±" in xml
    assert "σ" in xml
    assert "≤" in xml
    assert "∞" in xml


def test_omml_delimiters():
    r"""Тест скобок: \left( ... \right) и ( ... )."""
    xml1 = latex_to_omml(r"\left( \frac{a}{b} \right)")
    assert "<m:d>" in xml1
    assert '<m:begChr m:val="("' in xml1
    assert '<m:endChr m:val=")"' in xml1

    xml2 = latex_to_omml(r"(x + y)")
    assert "<m:d>" in xml2
    assert '<m:begChr m:val="("' in xml2


def test_inline_math_markdown(tmp_workspace):
    """Тест 3: Инлайн-формула в Markdown."""
    md = """# Документ с формулой
Формула кинетической энергии: $E_k = \\frac{m v^2}{2}$ в физике.
"""
    out = str(tmp_workspace / "inline_math.docx")
    res = md_build_v2(md, out, "Author")
    assert res["ok"], res.get("reason")

    doc = docx.Document(out)
    # Ищем параграф с текстом
    found = False
    for p in doc.paragraphs:
        if "Формула кинетической энергии" in p.text:
            # Внутри <w:p> должен быть элемент m:oMath
            omaths = p._p.findall(f".//{{{M_NS}}}oMath")
            assert len(omaths) >= 1
            found = True
            break
    assert found

    xml = _document_xml(out)
    assert "<m:oMath" in xml
    assert "<m:f>" in xml


def test_block_math_markdown(tmp_workspace):
    """Тест 4: Выключная блочная формула ($$ ... $$) с центрированием."""
    md = """# Блочная формула
$$
f(x) = \\frac{1}{\\sqrt{2\\pi\\sigma^2}} e^{-\\frac{(x-\\mu)^2}{2\\sigma^2}}
$$
"""
    out = str(tmp_workspace / "block_math.docx")
    res = md_build_v2(md, out, "Author")
    assert res["ok"], res.get("reason")

    doc = docx.Document(out)
    math_p = None
    for p in doc.paragraphs:
        if p._p.find(f".//{{{M_NS}}}oMathPara") is not None:
            math_p = p
            break
    assert math_p is not None
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    assert math_p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER

    xml = _document_xml(out)
    assert "<m:oMathPara" in xml
    assert "<m:f>" in xml
    assert "<m:rad>" in xml


def test_single_line_math_block(tmp_workspace):
    """Тест однострочной блочной формулы $$ formula $$."""
    md = """# Однострочный блок
$$ E = mc^2 $$
"""
    out = str(tmp_workspace / "single_block.docx")
    res = md_build_v2(md, out, "Author")
    assert res["ok"], res.get("reason")

    xml = _document_xml(out)
    assert "<m:oMathPara" in xml
    assert "mc" in xml or "m" in xml


def test_math_block_with_caption(tmp_workspace):
    """Тест MathBlock с подписью."""
    out = _build(tmp_workspace, [
        {"type": "math", "latex": r"\int_0^\infty e^{-x} dx = 1", "caption": "Формула 1.1 — Интеграл Пуассона"},
    ])
    doc = docx.Document(out)
    captions = [p for p in doc.paragraphs if "Формула 1.1" in p.text]
    assert len(captions) == 1
    cap_p = captions[0]
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    assert cap_p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER

    xml = _document_xml(out)
    assert "<m:nary>" in xml
    assert "Формула 1.1" in xml


def test_inline_math_in_table(tmp_workspace):
    """Тест инлайн-математики в ячейках таблицы."""
    md = """# Таблица с формулами
| Величина | Формула |
|:---|:---|
| Энергия | $E = mc^2$ |
| Импульс | $p = mv$ |
"""
    out = str(tmp_workspace / "table_math.docx")
    res = md_build_v2(md, out, "Author")
    assert res["ok"], res.get("reason")

    xml = _document_xml(out)
    assert "<m:oMath" in xml
    assert "<w:tbl>" in xml


def test_unclosed_math_block_fail_closed():
    """Тест незакрытого $$ блока (fail-closed)."""
    md = """# Ошибка
$$
\\frac{a}{b}
"""
    with pytest.raises(MdEngineError):
        md_build_v2(md, "/tmp/never.docx", "Author")


def test_api_functions_direct():
    """Тест прямых функций API: add_math_paragraph, add_math_to_paragraph."""
    doc = docx.Document()
    p_math = add_math_paragraph(doc, r"\sqrt{a^2 + b^2}")
    assert p_math._p.find(f".//{{{M_NS}}}oMathPara") is not None

    p_text = doc.add_paragraph("Теорема Пифагора: ")
    add_math_to_paragraph(p_text, r"c = \sqrt{a^2 + b^2}")
    assert p_text._p.find(f".//{{{M_NS}}}oMath") is not None
