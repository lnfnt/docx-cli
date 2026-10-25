"""v0.1.0 §2 B-10: вложенные списки и GFM pipe-таблицы в md-конвертере.

Дизайн: docs/design/nested-lists-tables.md (написан ДО реализации).
Регресс-пин легаси: tests/fixtures/v0100_legacy_md.json снят с кода v0.9.0.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import docx
import pytest

from docx_cli.md_court import build_md_document, md_build

FIXTURES = Path(__file__).parent / "fixtures"

WD = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _build(md: str, tmp_path: Path, name: str = "out.docx"):
    report: list = []
    doc = build_md_document(md, "colleague", report=report)
    path = tmp_path / name
    doc.save(str(path))
    reopened = docx.Document(str(path))
    return doc, reopened, report


# ----------------------------------------------------------------- списки


NESTED_UL = """ХОДАТАЙСТВО

- один
  - два
    - три
"""

NESTED_UL4 = """ХОДАТАЙСТВО

- L1 a
  - L2 b
    - L3 c
      - L4 d
"""

NESTED_DEEP = """ХОДАТАЙСТВО

- L1 a
  - L2 b
    - L3 c
      - L4 d
        - L5 skip me
"""


def _styles(doc):
    return [p.style.name for p in doc.paragraphs]


def test_v0100_nested_ul_level2_uses_list_bullet_2(tmp_path):
    doc, _, _ = _build(NESTED_UL, tmp_path)
    st = _styles(doc)
    assert "List Bullet 2" in st
    p2 = next(p for p in doc.paragraphs if p.style.name == "List Bullet 2")
    assert p2.text == "два"
    # level 1 остаётся легаси-рендером (без стиля, висячий отступ)
    p1 = next(p for p in doc.paragraphs if p.text == "один")
    assert p1.style.name == "Normal"
    assert abs(p1.paragraph_format.left_indent.cm - 1.25) < 0.01


def test_v0100_nested_ul_level3_uses_list_bullet_3(tmp_path):
    doc, _, _ = _build(NESTED_UL, tmp_path)
    assert "List Bullet 3" in _styles(doc)
    p3 = next(p for p in doc.paragraphs if p.style.name == "List Bullet 3")
    assert p3.text == "три"


def test_v0100_nested_ul_level4_falls_back_to_manual_indent(tmp_path):
    doc, _, _ = _build(NESTED_UL4, tmp_path)
    p4 = next(p for p in doc.paragraphs if p.text == "L4 d")
    assert p4.style.name == "Normal"  # стиля List Bullet 4 в шаблоне нет — fallback
    pf = p4.paragraph_format
    assert abs(pf.left_indent.cm - 3.5) < 0.01  # 1.25 + 0.75*3
    assert abs(pf.first_line_indent.cm + 0.75) < 0.01


def test_v0100_depth_gt5_is_honest_skip_marker(tmp_path):
    doc, _, report = _build(NESTED_DEEP, tmp_path)
    texts = [p.text for p in doc.paragraphs]
    assert not any("L5" in t for t in texts)
    marks = [r for r in report if r.get("type") == "skipped_construct"
             and r.get("construct") == "list_depth_gt4"]
    assert len(marks) == 1
    assert marks[0]["line"] == 7  # «        - L5 skip me» — 1-based


def test_v0100_mixed_ul_ol_neighbors(tmp_path):
    md = """ХОДАТАЙСТВО

- a
  1. b
  - c
  1. d
"""
    doc, _, _ = _build(md, tmp_path)
    st = _styles(doc)
    assert st.count("List Number 2") == 2
    assert st.count("List Bullet 2") == 1
    # порядок сохранён: a, b, c, d идут подряд
    idx = [next(i for i, p in enumerate(doc.paragraphs)
                if t in p.text) for t in ("a", "b", "c", "d")]
    assert idx == sorted(idx)


def test_v0100_ol_numbering_continues_across_ul_siblings(tmp_path):
    md = """ХОДАТАЙСТВО

1. one
   1. two
   - ul between
   1. two-continued
"""
    # уровень 2 = List Number 2 (стиль автонумерует), уровень 2 в fallback нет —
    # проверяем счётчик через level-4 fallback ниже отдельно; здесь стили
    doc, _, _ = _build(md, tmp_path)
    st = _styles(doc)
    assert st.count("List Number 2") == 2
    assert st.count("List Bullet 2") == 1


def test_v0100_fallback_ol_renders_computed_numbers(tmp_path):
    # unit=3; level 4 = fallback (стиля List Number 4 нет) — номера computed
    md = """ХОДАТАЙСТВО

1. a
   1. b
      1. c
         1. four-one
         - ul sibling
         1. four-two
"""
    doc, _, _ = _build(md, tmp_path)
    texts = [p.text for p in doc.paragraphs]
    assert "1. four-one" in texts
    assert "2. four-two" in texts  # счётчик продолжился через ul-соседа


# ----------------------------------------------------------------- таблицы


SIMPLE_TABLE = """ХОДАТАЙСТВО

| № | Предмет | Срок |
|---|:---:|---:|
| 1 | **жирный** | 10.01 |
| 2 | обычный | 20.02 |
"""


def test_v0100_table_converted_with_grid_and_alignments(tmp_path):
    doc, reopened, _ = _build(SIMPLE_TABLE, tmp_path)
    assert len(reopened.tables) == 1
    t = reopened.tables[0]
    assert len(t.rows) == 3 and len(t.columns) == 3
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    header_cell = t.cell(0, 0).paragraphs[0]
    assert header_cell.alignment == A.LEFT
    assert t.cell(0, 1).paragraphs[0].alignment == A.CENTER
    assert t.cell(0, 2).paragraphs[0].alignment == A.RIGHT
    assert t.cell(1, 2).paragraphs[0].alignment == A.RIGHT
    assert t.style.name == "Table Grid"


def test_v0100_table_header_bold_and_repeat(tmp_path):
    doc, reopened, _ = _build(SIMPLE_TABLE, tmp_path)
    t = reopened.tables[0]
    assert all(r.bold for p in t.cell(0, 1).paragraphs for r in p.runs if r.text)
    trPr = t.rows[0]._tr.find(f"{{{WD}}}trPr")
    assert trPr is not None
    assert trPr.find(f"{{{WD}}}tblHeader") is not None


def test_v0100_inline_bold_in_cell_no_markers_leak(tmp_path):
    doc, reopened, _ = _build(SIMPLE_TABLE, tmp_path)
    cell = reopened.tables[0].cell(1, 1)
    assert "жирный" in cell.text
    assert "**" not in cell.text
    runs = [r for p in cell.paragraphs for r in p.runs if r.text]
    assert any(r.bold for r in runs)


def test_v0100_inline_code_in_cell_stays_times_new_roman(tmp_path):
    """Перепинен контролем 24.10: бэктик-код в ячейке остаётся Times New Roman —
    валидатор font_name бракует любые другие шрифты, render-контракт важнее
    моноширинной косметики (см. md_court._render_cell_inline: коммент fix)."""
    import subprocess, sys
    src = tmp_path / "c.md"
    src.write_text("# АПЕЛЛЯЦИОННАЯ ЖАЛОБА\n\n| Колонка |\n|---|\n| `код` |\n", encoding="utf-8")
    out = tmp_path / "c.docx"
    r = subprocess.run([sys.executable, "-m", "docx_cli", "md-build", str(src), str(out)],
                       capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    v = subprocess.run([sys.executable, "-m", "docx_cli", "validate", str(out)],
                       capture_output=True, text=True)
    assert v.returncode == 0, v.stdout
    from docx import Document
    d = Document(str(out))
    cell_run = d.tables[0].rows[1].cells[0].paragraphs[0].runs[0]
    assert cell_run.font.name == "Times New Roman"

def test_v0100_inline_italic_and_underline_in_cell(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б |
|---|---|
| _итал_ | <u>чёрк</u> |
"""
    doc, reopened, _ = _build(md, tmp_path)
    t = reopened.tables[0]
    r_it = [r for p in t.cell(1, 0).paragraphs for r in p.runs if r.text]
    assert r_it[0].italic is True and "_" not in t.cell(1, 0).text
    r_un = [r for p in t.cell(1, 1).paragraphs for r in p.runs if r.text]
    assert r_un[0].underline is True and "<u>" not in t.cell(1, 1).text


def test_v0100_escaped_pipe_stays_literal_in_cell(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б |
|---|---|
| left \\| right | x |
"""
    doc, reopened, _ = _build(md, tmp_path)
    cell = reopened.tables[0].cell(1, 0)
    assert cell.text == "left | right"
    assert "\\" not in cell.text


def test_v0100_block_in_cell_skips_row_with_marker(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б |
|---|---|
| > цитата внутри | x |
| ok | fine |
"""
    doc, reopened, report = _build(md, tmp_path)
    t = reopened.tables[0]
    assert len(t.rows) == 2  # header + «ok fine», битая строка пропущена
    assert t.cell(1, 0).text == "ok"
    marks = [r for r in report if r.get("construct") == "block_in_cell"]
    assert len(marks) == 1 and "цитата" in str(marks[0]["text"])


def test_v0100_list_in_cell_skipped_like_quote(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б |
|---|---|
| - x | y |
"""
    doc, reopened, report = _build(md, tmp_path)
    assert len(reopened.tables[0].rows) == 1
    assert any(r.get("construct") == "block_in_cell" for r in report)


def test_v0100_header_only_table_renders_one_row(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б | В |
|---|---|---|
"""
    doc, reopened, report = _build(md, tmp_path)
    assert len(reopened.tables) == 1
    t = reopened.tables[0]
    assert len(t.rows) == 1
    assert not [r for r in report if r.get("type") == "skipped_construct"]


def test_v0100_row_width_mismatch_padded_and_marked(tmp_path):
    md = """ХОДАТАЙСТВО

| А | Б | В |
|---|---|---|
| короткая |
| a | b | c |
"""
    doc, reopened, report = _build(md, tmp_path)
    t = reopened.tables[0]
    assert len(t.rows) == 3
    assert t.cell(1, 2).text == ""
    assert any(r.get("construct") == "row_width_mismatch" for r in report)


def test_v0100_tables_not_parsed_in_header_zone(tmp_path):
    md = """Истец: |фамилия| лицо|

ХОДАТАЙСТВО

Текст.
"""
    doc, reopened, _ = _build(md, tmp_path)
    assert len(reopened.tables) == 0
    assert any("|фамилия| лицо|" in p.text for p in doc.paragraphs)


def test_v0100_no_cjk_leak_after_conversion(tmp_path):
    md = NESTED_UL + "\n| a | b |\n|---|---|\n| 1 | 2 |\n"
    doc, reopened, _ = _build(md, tmp_path, name="cjk.docx")
    texts = [p.text for p in reopened.paragraphs]
    for t in reopened.tables:
        for row in t.rows:
            for c in row.cells:
                texts.append(c.text)
    joined = "\n".join(texts)
    assert not re.search(r"[\u3400-\u9fff\u3040-\u30ff\uf900-\ufaff\uff66-\uff9f]",
                        joined)


# ----------------------------------------------------------------- round-trip


def _extract_structure(path: Path) -> dict:
    d = docx.Document(str(path))
    items = []
    for p in d.paragraphs:
        m = re.match(r"^List (Bullet|Number) (\d+)$", p.style.name)
        if m:
            items.append({"kind": "ul" if m.group(1) == "Bullet" else "ol",
                          "level": int(m.group(2)), "text": p.text})
        else:
            pf = p.paragraph_format
            left = round(pf.left_indent.cm, 2) if pf.left_indent else None
            first = round(pf.first_line_indent.cm, 2) if pf.first_line_indent else None
            if first is not None and first < 0 and left:
                kind = "ol" if re.match(r"^\d+[.)] ", p.text) else "ul"
                lvl = 1 if abs(left - 1.25) < 0.02 else \
                    2 if abs(left - 2.0) < 0.02 else \
                    3 if abs(left - 2.75) < 0.02 else \
                    4 if abs(left - 3.5) < 0.02 else 0
                if lvl:
                    items.append({"kind": kind, "level": lvl, "text": p.text})
    tables = []
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    amap = {A.LEFT: "left", A.CENTER: "center", A.RIGHT: "right", A.JUSTIFY: "justify",
            None: "left"}
    for t in d.tables:
        cols = len(t.columns)
        aligns = [amap.get(t.cell(0, c).paragraphs[0].alignment, "left")
                  for c in range(cols)]
        tables.append({
            "cols": cols,
            "rows": len(t.rows),
            "aligns": aligns,
            "header": [t.cell(0, c).text for c in range(cols)],
            "body": [[t.cell(r, c).text for c in range(cols)]
                     for r in range(1, len(t.rows))],
        })
    return {"items": items, "tables": tables}


def test_v0100_roundtrip_nested_lists(tmp_path):
    md = """ХОДАТАЙСТВО

Текст перед списком.

1. Раз
   1. Два
      - Три-ul
   2. Четыре
2. Пять
"""
    doc, _, _ = _build(md, tmp_path, name="rt1.docx")
    # путь сохранён в _build; переизвлекаем с диска
    struct = _extract_structure(tmp_path / "rt1.docx")
    kinds = [(i["kind"], i["level"]) for i in struct["items"]]
    assert kinds == [("ol", 1), ("ol", 2), ("ul", 3), ("ol", 2), ("ol", 1)]
    texts = [i["text"] for i in struct["items"]]
    assert "Раз" in texts[0] and "Пять" in texts[-1]


def test_v0100_roundtrip_table_alignment_and_cells(tmp_path):
    md = """ХОДАТАЙСТВО

| A | B | C |
|:--|:--:|--:|
| a\\|b | **ж** | 9 |
| x | y | z |
"""
    _build(md, tmp_path, name="rt2.docx")
    struct = _extract_structure(tmp_path / "rt2.docx")
    assert len(struct["tables"]) == 1
    t = struct["tables"][0]
    assert t["header"] == ["A", "B", "C"]
    assert t["aligns"] == ["left", "center", "right"]
    assert t["body"][0] == ["a|b", "ж", "9"]
    assert t["rows"] == 3 and t["cols"] == 3


# ----------------------------------------------------------------- легаси-пин


LEGACY = json.loads((FIXTURES / "v0100_legacy_md.json").read_text(encoding="utf-8"))


def _para_sig(p, i):
    pf = p.paragraph_format
    return {
        "i": i,
        "text": p.text,
        "style": p.style.name if p.style else None,
        "align": str(p.alignment) if p.alignment is not None else None,
        "line_spacing": pf.line_spacing,
        "space_before_pt": pf.space_before.pt if pf.space_before is not None else None,
        "space_after_pt": pf.space_after.pt if pf.space_after is not None else None,
        "left_cm": round(pf.left_indent.cm, 3) if pf.left_indent is not None else None,
        "first_cm": round(pf.first_line_indent.cm, 3) if pf.first_line_indent is not None else None,
        "runs": [
            {"text": r.text, "bold": r.bold, "italic": r.italic,
             "underline": r.underline, "font": r.font.name,
             "size_pt": r.font.size.pt if r.font.size is not None else None}
            for r in p.runs
        ],
    }


def _signature(doc):
    return {"author": doc.core_properties.author,
            "paragraphs": [_para_sig(p, i) for i, p in enumerate(doc.paragraphs)],
            "tables": len(doc.tables)}


@pytest.mark.parametrize("sample", sorted(LEGACY))
def test_v0100_legacy_md_structure_is_byte_stable(sample):
    """Старый md без таблиц/вложенностей: структура вывода и report == v0.9.0."""
    case = LEGACY[sample]
    report: list = []
    doc = build_md_document(case["md"], "colleague", report=report)
    assert _signature(doc) == case["signature"]
    assert report == case["report"]


def test_v0100_md_build_envelope_keys_stable(tmp_path):
    src = tmp_path / "in.md"
    src.write_text(LEGACY["title_first"]["md"], encoding="utf-8")
    out = tmp_path / "env.docx"
    res = md_build(str(src), str(out), "colleague")
    assert {"output_path", "author", "blocks_rendered", "report", "reports",
            "fixes", "changes", "normalized", "dash_fixes",
            "dash_normalized"} <= set(res)


# ----------------------------------------------------------------- дизайн-контракт


def test_v0100_design_doc_exists_before_implementation():
    p = Path(__file__).parent.parent / "docs" / "design" / "nested-lists-tables.md"
    assert p.is_file()
    text = p.read_text(encoding="utf-8")
    assert "skipped_construct" in text and "block_in_cell" in text


def test_table_validate_clean_after_cell_size_fix(tmp_path):
    """Координаторский пин: md-build с таблицей проходит validate без cell_font_size."""
    src = tmp_path / "z.md"
    src.write_text(
        "# АПЕЛЛЯЦИОННАЯ ЖАЛОБА\n\n| А | Б |\n|:--|--:|\n| **x** | y |\n",
        encoding="utf-8")
    out = tmp_path / "z.docx"
    import subprocess, sys
    r = subprocess.run([sys.executable, "-m", "docx_cli", "md-build", str(src), str(out)],
                       capture_output=True, text=True, cwd=str(tmp_path.parent.parent) if False else None)
    assert r.returncode == 0, r.stderr
    v = subprocess.run([sys.executable, "-m", "docx_cli", "validate", str(out)],
                       capture_output=True, text=True)
    doc = json.loads(v.stdout or v.stderr)
    assert v.returncode == 0, doc
