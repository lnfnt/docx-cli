# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 3: богатая инлайн-типографика + настоящие многоуровневые списки OpenXML.

Спецификация: TASK_PACKAGE3_TYPOGRAPHY.md.
Проверяет XML-теги vertAlign/strike/highlight/Consolas в document.xml,
w:numPr/w:ilvl/w:numId, определения w:abstractNum/w:num в word/numbering.xml,
парсинг вложенных списков по отступам в md_engine.
"""
from __future__ import annotations

import zipfile

import pytest
from docx import Document
from docx.enum.text import WD_COLOR_INDEX

from docx_cli.blocks import _split_inline
from docx_cli.md_engine import MdEngineError, md_build_v2
from docx_cli.numbering import add_multilevel_list

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _xml(path, part: str = "word/document.xml") -> str:
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read(part).decode("utf-8")


MD_TYPOGRAPHY = """# Описание формул

Формула H~2~O и площадь 100 м^2^. Слово ~~удалено~~, а это ==важно== и код `def test()`.
"""


def _runs(doc):
    return [r for p in doc.paragraphs for r in p.runs if r.text]


def _find(doc, text: str):
    return [r for r in _runs(doc) if r.text == text]


# ---------------------------------------------------------------- инлайн

def test_split_inline_tokens():
    segs = _split_inline("H~2~O ~~x~~ ^sup^ ==m== `c` **b** *i*")
    by = {}
    for s in segs:
        by.setdefault(s.text, s)
    assert segs[0].text == "H"
    sub = next(s for s in segs if s.text == "2")
    assert sub.vert_align == "subscript"
    strike = next(s for s in segs if s.text == "x")
    assert strike.strike is True
    sup = next(s for s in segs if s.text == "sup")
    assert sup.vert_align == "superscript"
    mark = next(s for s in segs if s.text == "m")
    assert mark.highlight == "yellow"
    code = next(s for s in segs if s.text == "c")
    assert code.code is True and code.bold is False
    assert next(s for s in segs if s.text == "b").bold is True
    assert next(s for s in segs if s.text == "i").italic is True


def test_split_inline_code_is_raw():
    # внутри кода маркеры не парсятся и бэктик закрывает код
    segs = _split_inline("`a**b~c` ~~s~~")
    assert segs[0].text == "a**b~c" and segs[0].code
    s = next(x for x in segs if x.text == "s")
    assert s.strike and not s.code


def test_md_typography_build_envelope(tmp_path):
    out = tmp_path / "typo.docx"
    res = md_build_v2(MD_TYPOGRAPHY, str(out), "pytest")
    assert res["ok"] is True, res


def test_md_subscript_superscript_xml(tmp_path):
    out = tmp_path / "subs.docx"
    md_build_v2(MD_TYPOGRAPHY, str(out), "pytest")
    xml = _xml(out)
    assert '<w:vertAlign w:val="subscript"/>' in xml
    assert '<w:vertAlign w:val="superscript"/>' in xml
    doc = Document(str(out))
    # нижний индекс у "2" в H~2~O и верхний у "2" в м^2^
    subs = [r for r in _runs(doc) if r.text == "2" and r.font.subscript]
    sups = [r for r in _runs(doc) if r.text == "2" and r.font.superscript]
    assert len(subs) == 1 and len(sups) == 1


def test_md_strike_highlight_code_xml(tmp_path):
    out = tmp_path / "marks.docx"
    md_build_v2(MD_TYPOGRAPHY, str(out), "pytest")
    xml = _xml(out)
    assert "<w:strike/>" in xml
    assert '<w:highlight w:val="yellow"/>' in xml
    doc = Document(str(out))
    strike = next(r for r in _runs(doc) if r.text == "удалено")
    assert strike.font.strike is True
    mark = next(r for r in _runs(doc) if r.text == "важно")
    assert mark.font.highlight_color == WD_COLOR_INDEX.YELLOW
    code = next(r for r in _runs(doc) if r.text == "def test()")
    assert code.font.name == "Consolas"
    assert 'w:ascii="Consolas"' in xml


# ---------------------------------------------------------------- списки

MD_NESTED_OL = """# Порядок работ

1. Пункт 1
   1. Пункт 1.1
      1. Пункт 1.1.1
2. Пункт 2
"""

MD_NESTED_UL = """# Дерево

- корень
  - ветка
    - лист
"""


def test_nested_numbered_list_uses_numpr_and_numbering_xml(tmp_path):
    out = tmp_path / "nested.docx"
    res = md_build_v2(MD_NESTED_OL, str(out), "pytest")
    assert res["ok"] is True, res
    doc_xml = _xml(out)
    # numPr/ilvl в документе
    assert '<w:numPr>' in doc_xml or f'<w:numPr ' in doc_xml
    assert '<w:ilvl w:val="0"/>' in doc_xml
    assert '<w:ilvl w:val="1"/>' in doc_xml
    assert '<w:numId w:val="10"/>' in doc_xml  # > 9 из шаблона python-docx
    # определения в word/numbering.xml
    num_xml = _xml(out, "word/numbering.xml")
    assert '<w:abstractNum' in num_xml
    assert '<w:lvlText w:val="%1."/>' in num_xml
    assert '<w:lvlText w:val="%1.%2."/>' in num_xml
    assert '<w:multiLevelType w:val="multilevel"/>' in num_xml
    assert '<w:numFmt w:val="decimal"/>' in num_xml
    # связь numId -> abstractNumId
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert "Пункт 1" in texts and "Пункт 1.1.1" in texts
    assert "1. Пункт 1" not in texts  # нумерация настоящая, не текст


def test_nested_bullet_list_uses_numpr(tmp_path):
    out = tmp_path / "bullets.docx"
    res = md_build_v2(MD_NESTED_UL, str(out), "pytest")
    assert res["ok"] is True, res
    doc_xml = _xml(out)
    doc = Document(str(out))
    from docx.oxml.ns import qn
    prs = []
    for p in doc.paragraphs:
        if p.text in ("корень", "ветка", "лист"):
            prs.append(p._p.pPr.find(qn("w:numPr")))
    assert len(prs) == 3 and all(p is not None for p in prs)
    ilvls = [int(p.find(qn("w:ilvl")).get(qn("w:val"))) for p in prs]
    assert ilvls == [0, 1, 2]
    num_xml = _xml(out, "word/numbering.xml")
    assert '<w:numFmt w:val="bullet"/>' in num_xml


def test_flat_lists_keep_gost_legacy_render(tmp_path):
    out = tmp_path / "flat.docx"
    md_build_v2("# Д\n\n1. первый\n2. второй\n", str(out), "pytest")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "1. первый" in texts and "2. второй" in texts
    doc_xml = _xml(out)
    assert "<w:numPr>" not in doc_xml and "<w:numPr " not in doc_xml


def test_mixed_nested_list_runs(tmp_path):
    md = "# М\n\n1. a\n   - b\n2. c\n"
    out = tmp_path / "mix.docx"
    md_build_v2(md, str(out), "pytest")
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert texts.count("1. a") == 1 and "b" in texts and "1. c" in texts


def test_nested_numbered_multilevel_text_no_prefix(tmp_path):
    out = tmp_path / "prefix.docx"
    md_build_v2(MD_NESTED_OL, str(out), "pytest")
    texts = [p.text for p in Document(str(out)).paragraphs]
    assert "Пункт 1.1" in texts
    assert "1.1. Пункт 1.1" not in texts


# ---------------------------------------------------------------- numbering.py

def test_add_multilevel_list_allocates_ids_deterministically(tmp_path):
    from docx import Document as Doc
    d = Doc()
    num1 = add_multilevel_list(d, kind="decimal", start=1)
    num2 = add_multilevel_list(d, kind="bullet")
    assert num1 != num2 and num1 >= 10
    numbering = d.part.numbering_part.element
    from docx.oxml.ns import qn
    nums = numbering.findall(qn("w:num"))
    assert nums[-2].get(qn("w:numId")) == str(num1)
    assert nums[-1].get(qn("w:numId")) == str(num2)
    lvl2 = (numbering.findall(qn("w:abstractNum"))[-2]
            .findall(qn("w:lvl"))[2])
    assert lvl2.find(qn("w:lvlText")).get(qn("w:val")) == "%1.%2.%3."
    assert lvl2.find(qn("w:pPr")).find(qn("w:ind")).get(qn("w:left")) == "1080"


def test_start_at_starts_level0_counter(tmp_path):
    from docx_cli.spec import DocSpec, HeaderSpec, NumberedListBlock, \
        OutputSpec, ParagraphBlock, TitleSpec
    from docx_cli import builder
    spec = DocSpec(
        profile="corporate",
        output=OutputSpec(file_name="s.docx", author="t"),
        header=HeaderSpec(lines=["h"]),
        title=TitleSpec(main="С"),
        body=[ParagraphBlock(type="paragraph", text="т"),
              NumberedListBlock(type="numbered_list",
                                items=["хвост", "хвост 1.1"],
                                start_at=5, levels=[0, 1])],
    )
    out = tmp_path / "s.docx"
    res = builder.build_document(spec, str(out))
    assert res["ok"] is True, res
    from docx.oxml.ns import qn
    num_xml = _xml(out, "word/numbering.xml")
    assert '<w:start w:val="5"/>' in num_xml


def test_levels_validator_range():
    from docx_cli.spec import NumberedListBlock
    with pytest.raises(ValueError):
        NumberedListBlock(type="numbered_list", items=["a"], levels=[3])
    with pytest.raises(ValueError):
        NumberedListBlock(type="numbered_list", items=["a", "b"], levels=[0])
    NumberedListBlock(type="numbered_list", items=["a"], levels=[0])  # ок
