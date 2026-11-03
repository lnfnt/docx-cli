# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Block-level rendering helpers for the DOCX builder."""
from __future__ import annotations

import os
import re
from typing import NamedTuple

from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt, RGBColor
from docx.table import Table

from docx_cli.errors import BuildError
from docx_cli.comments import add_comment_to_runs
from docx_cli.footnotes import add_footnote, split_footnote_markers
from docx_cli.revisions import add_revision
from docx_cli.spec import (
    AttachmentsListBlock,
    BulletListBlock,
    CalloutBlock,
    DocSpec,
    HeaderSpec,
    HeadingBlock,
    HorizontalRuleBlock,
    ImageBlock,
    NumberedListBlock,
    PageBreakBlock,
    ParagraphBlock,
    PleadingBlock,
    Profile,
    SectionBreakBlock,
    SignatureBlock,
    TableBlock,
    TableCell,
    TitleSpec,
    TocBlock,
)

NS = nsdecls("w")

# Фаза 2/3 ТЗ: формат заголовков md → (размер pt, space_before, space_after).
_HEADING_FMT = {1: (18, 12, 6), 2: (16, 10, 4), 3: (14, 8, 2), 4: (12, 0, 0)}

# Порядок дочерних элементов pPr после w:pBdr/w:shd (ECMA-376 CT_PPr).
_PPR_AFTER_SHD = (
    "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
    "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd",
    "w:snapToGrid", "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
    "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
    "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr", "w:sectPr",
    "w:pPrChange",
)
_PPR_AFTER_PBDR = ("w:shd",) + _PPR_AFTER_SHD
# Порядок дочерних pPr после w:numPr (ECMA-376 CT_PPr): numPr стоит до
# suppressLineNumbers/pBdr — те же наследники, что и у pBdr, + pBdr.
_PPR_AFTER_NUMPR = ("w:suppressLineNumbers",) + _PPR_AFTER_PBDR
# Порядок дочерних элементов tcPr после границ/заливки (CT_TcPr).
_TCPR_AFTER_BORDERS = (
    "w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign",
    "w:hideMark", "w:headers", "w:cellIns", "w:cellDel", "w:cellMerge", "w:tcPrChange",
)
_BORDER_EDGES = ("top", "left", "bottom", "right", "insideH", "insideV")


# --------------------------------------------------------------------------------------
# Low-level helpers
# --------------------------------------------------------------------------------------

def _set_run_font(run, profile: Profile, size_pt: int, bold: bool = False, italic: bool = False,
                  strike: bool = False, vert_align: str | None = None,
                  highlight: str | None = None, code: bool = False) -> None:
    if code:
        # ПАКЕТ 3: инлайн-код — моноширинный Consolas на 1pt меньше тела
        # (validate._check_fonts_and_sizes освобождает code-runs от проверки шрифта).
        run.font.name = "Consolas"
        run.font.size = Pt(size_pt - 1)
    else:
        run.font.name = profile.font_name
        run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.rtl = False
    run._element.rPr.rFonts.set(qn("w:eastAsia"), run.font.name)
    # ПАКЕТ 3: богатая инлайн-типографика (python-docx ставит элементы rPr
    # в правильном месте CT_RPrschema — без ручного parse_xml).
    if strike:
        run.font.strike = True
    if vert_align == "superscript":
        run.font.superscript = True
    elif vert_align == "subscript":
        run.font.subscript = True
    if highlight:  # кодовая карта константа: маркер ==...== всегда жёлтый
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

def _body_paragraph_format(p, profile: Profile) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.body_line_spacing
    pf.first_line_indent = Cm(profile.first_line_indent_cm)
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _header_paragraph_format(p, profile: Profile) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.header_line_spacing
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)


def _title_paragraph_format(p, profile: Profile) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.header_line_spacing
    if profile.name == "court":
        pf.space_after = Pt(12)
        pf.space_before = Pt(12)
    else:
        pf.space_after = Pt(6)
        pf.space_before = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER


def _cell_paragraph_format(p, profile: Profile) -> None:
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.header_line_spacing
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)


# --------------------------------------------------------------------------------------
# Inline markdown
# --------------------------------------------------------------------------------------

class InlineSegment(NamedTuple):
    """ПАКЕТ 3: сегмент инлайн-типографики.

    vert_align: None | "superscript" | "subscript";
    highlight: None | "yellow" (маркер ==...==);
    code: инлайн-код (Consolas, на 1pt меньше тела);
    math: инлайн-математика OMML (LaTeX-строка).
    """

    text: str
    bold: bool = False
    italic: bool = False
    strike: bool = False
    vert_align: str | None = None
    highlight: str | None = None
    code: bool = False
    math: bool = False

    # Совместимость со старым контрактом (text, bold, italic).
    def __iter__(self):
        return iter((self.text, self.bold, self.italic))


def _split_inline(text: str) -> list[InlineSegment]:
    r"""Parse inline markdown into InlineSegment. Supported markers:

    ``**bold**``, ``*italic*``, ``~~strike~~``, ``^sup^``, ``~sub~``
    (single ``~`` не конфликтует с ``~~``), ``==highlight==``, ``\`code\````.
    """
    segments: list[InlineSegment] = []
    buf: list[str] = []
    bold = italic = strike = False
    vert: str | None = None
    highlight: str | None = None
    code = False

    def flush() -> None:
        if buf:
            segments.append(InlineSegment(
                "".join(buf), bold, italic, strike, vert, highlight, code, math=False))
            buf.clear()

    i = 0
    n = len(text)
    while i < n:
        if code:
            # Внутри кода всё литерально, закрывает только следующий бэктик.
            if text.startswith("`", i):
                # flush ДО закрытия: сегмент несёт состояние, в котором копился.
                flush()
                code = False
                i += 1
            else:
                buf.append(text[i])
                i += 1
            continue
        if text.startswith("$", i) and not text.startswith("$$", i):
            # ПАКЕТ 7: инлайн-математика $...$ (обрабатываем до экранирования \\, чтобы \frac не ломался)
            # Границы по pandoc: $ не за пробелом, закрывающий $ не перед пробелом
            # (иначе "цена 100 $ и 200 $" превращается в формулу).
            end_idx = text.find("$", i + 1)
            if (end_idx != -1 and end_idx > i + 1
                    and not text[i + 1].isspace()
                    and not text[end_idx - 1].isspace()):
                latex_expr = text[i + 1:end_idx]
                # S1-H5: денежный контекст — литерал, не формула.
                # (а) выражение из одних цифр/разделителей ($100 000);
                # (б) после закрывающего $ идёт цифра или ')' — «($100) и ($200)».
                _nxt = text[end_idx + 1] if end_idx + 1 < len(text) else ""
                _moneyish = (re.fullmatch(r"[\d\s.,]+", latex_expr) is not None
                             and any(ch.isdigit() for ch in latex_expr))
                if latex_expr.strip() and not _moneyish \
                        and not _nxt.isdigit() and _nxt != ")":
                    flush()
                    segments.append(InlineSegment(
                        latex_expr, bold=False, italic=False, strike=False,
                        vert_align=None, highlight=None, code=False, math=True))
                    i = end_idx + 1
                    continue
            buf.append(text[i])
            i += 1
        elif text.startswith("\\", i) and i + 1 < n:
            buf.append(text[i + 1])
            i += 2
        elif text.startswith("**", i):
            flush()
            bold = not bold
            i += 2
        elif text.startswith("==", i):
            flush()
            highlight = None if highlight else "yellow"
            i += 2
        elif text.startswith("~~", i):
            flush()
            strike = not strike
            i += 2
        elif text[i] == "*":
            flush()
            italic = not italic
            i += 1
        elif text[i] == "^":
            flush()
            vert = None if vert == "superscript" else "superscript"
            i += 1
        elif text[i] == "~":
            flush()
            vert = None if vert == "subscript" else "subscript"
            i += 1
        elif text[i] == "`":
            flush()
            code = True
            i += 1
        else:
            buf.append(text[i])
            i += 1
    flush()
    return segments


def _add_inline_text(p, text: str, profile: Profile, size_pt: int, italic: bool = False) -> list:
    """Сегменты инлайн-разметки -> runs параграфа; возвращает список run'ов
    (нужен рендереру комментариев — ПАКЕТ 4)."""
    from docx_cli.omml import add_math_to_paragraph
    runs = []
    for seg in _split_inline(text):
        if seg.math:
            add_math_to_paragraph(p, seg.text)
            continue
        run = p.add_run(seg.text)
        _set_run_font(run, profile, size_pt, bold=seg.bold,
                      italic=italic != seg.italic, strike=seg.strike,
                      vert_align=seg.vert_align, highlight=seg.highlight,
                      code=seg.code)
        runs.append(run)
    return runs


# --------------------------------------------------------------------------------------
# Public block renderers
# --------------------------------------------------------------------------------------

def render_header(doc: Document, spec: HeaderSpec, profile: Profile) -> None:
    if profile.name == "court":
        last_p = None
        for line in spec.lines:
            if line == "":
                if last_p is not None:
                    last_p.paragraph_format.space_after = Pt(3)
                continue
            p = doc.add_paragraph()
            _header_paragraph_format(p, profile)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Cm(spec.left_indent_cm)
            run = p.add_run(line)
            _set_run_font(run, profile, profile.body_font_size_pt)
            last_p = p
        return

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }
    for line in spec.lines:
        p = doc.add_paragraph()
        _header_paragraph_format(p, profile)
        p.paragraph_format.alignment = align_map[spec.align]
        run = p.add_run(line)
        _set_run_font(run, profile, profile.body_font_size_pt)


def render_title(doc: Document, spec: TitleSpec, profile: Profile) -> None:
    from docx_cli.bookmarks import ANCHOR_RE, add_bookmark
    p = doc.add_paragraph()
    _title_paragraph_format(p, profile)
    
    # ПАКЕТ 8: якоря {#anchor-name} в титульном заголовке
    main_text = spec.main
    anchors = []
    for m in ANCHOR_RE.finditer(main_text):
        anchors.append(m.group(1) or m.group(2))
    main_text = ANCHOR_RE.sub("", main_text).strip()

    run = p.add_run(main_text)
    _set_run_font(run, profile, profile.title_font_size_pt, bold=True)
    # Фазы 2/3: корпоративные профили красят заголовки в цвет профиля.
    if profile.heading_color:
        run.font.color.rgb = RGBColor.from_string(profile.heading_color)
    if spec.subtitle:
        p2 = doc.add_paragraph()
        _title_paragraph_format(p2, profile)
        run2 = p2.add_run(spec.subtitle)
        _set_run_font(run2, profile, profile.body_font_size_pt, bold=False)

    for anchor in anchors:
        add_bookmark(p, anchor)


def render_heading(doc: Document, block: HeadingBlock, profile: Profile) -> None:
    """Фаза 2/3 ТЗ §А.2: `#`..`####` -> Заголовок 1-4.

    18pt/12:6, 16pt/10:4, 14pt/8:2, 12pt; bold, keep_with_next; цвет —
    из профиля (corporate/report #1F4E79) или явный block.color.
    """
    from docx_cli.bookmarks import ANCHOR_RE, add_bookmark
    size, before, after = _HEADING_FMT[block.level]
    p = doc.add_paragraph()
    try:
        p.style = doc.styles[f"Heading {block.level}"]
    except (KeyError, IndexError):
        pass  # стиль может отсутствовать в нестандартном шаблоне
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0)
    pf.keep_with_next = True
    pf.alignment = WD_ALIGN_PARAGRAPH.LEFT

    # ПАКЕТ 8: якоря {#anchor-name} или <a id="anchor-name"></a>
    heading_text = block.text
    anchors = []
    for m in ANCHOR_RE.finditer(heading_text):
        anchors.append(m.group(1) or m.group(2))
    heading_text = ANCHOR_RE.sub("", heading_text).strip()

    run = p.add_run(heading_text)
    _set_run_font(run, profile, size, bold=True)
    color = _hex_fill(block.color) or profile.heading_color
    if color:
        run.font.color.rgb = RGBColor.from_string(color)

    for anchor in anchors:
        add_bookmark(p, anchor)



def _render_paragraph_inline_with_refs(p, text: str, profile: Profile, size_pt: int) -> list:
    """Рендерит текст, содержащий якоря, перекрестные ссылки и инлайн-разметку."""
    from docx_cli.bookmarks import ANCHOR_RE, CROSS_REF_RE, add_bookmark, add_cross_reference
    runs = []
    pos = 0
    # Находим все якоря и ссылки в тексте в порядке следования
    tokens = []
    for m in ANCHOR_RE.finditer(text):
        tokens.append((m.start(), m.end(), "anchor", m.group(1) or m.group(2), ""))
    for m in CROSS_REF_RE.finditer(text):
        if m.group(4):
            # [@page:#bm]
            tokens.append((m.start(), m.end(), "cross_ref", m.group(4), "", True))
        else:
            is_page = bool(m.group(2))
            tokens.append((m.start(), m.end(), "cross_ref", m.group(3), m.group(1), is_page))
    tokens.sort(key=lambda t: t[0])

    for tok in tokens:
        start, end, kind = tok[0], tok[1], tok[2]
        if start > pos:
            chunk = text[pos:start]
            runs.extend(_add_inline_text(p, chunk, profile, size_pt))
        if kind == "anchor":
            add_bookmark(p, tok[3])
        elif kind == "cross_ref":
            add_cross_reference(p, bookmark_name=tok[3], display_text=tok[4],
                                page_ref=tok[5], profile=profile, size_pt=size_pt)
        pos = end
    if pos < len(text) or not tokens:
        runs.extend(_add_inline_text(p, text[pos:], profile, size_pt))
    return runs


def render_paragraph(doc: Document, block: ParagraphBlock, profile: Profile) -> None:
    # ПАКЕТ 4: сплиттер разметки рецензирования живет в md_engine (парсинг
    # CriticMarkup); здесь циклический импорт исключается — отложенный импорт.
    from docx_cli.md_engine import ReviewSegment, split_review_segments
    from docx_cli.endnotes import add_endnote, split_endnote_markers

    p = doc.add_paragraph()
    _body_paragraph_format(p, profile)
    for chunk, ref in split_footnote_markers(block.text):
        if chunk:
            for echunk, eref in split_endnote_markers(chunk):
                if echunk:
                    for seg in split_review_segments(echunk):
                        if seg.kind == "text":
                            if block.inline_markdown:
                                _render_paragraph_inline_with_refs(p, seg.text, profile, profile.body_font_size_pt)
                            else:
                                run = p.add_run(seg.text)
                                _set_run_font(run, profile, profile.body_font_size_pt)
                        elif seg.kind == "ins":
                            add_revision(p, inserted_text=seg.text,
                                         font=profile.font_name, size_pt=profile.body_font_size_pt)
                        elif seg.kind == "del":
                            add_revision(p, deleted_text=seg.text,
                                         font=profile.font_name, size_pt=profile.body_font_size_pt)
                        elif seg.kind == "comment":
                            if block.inline_markdown:
                                runs = _render_paragraph_inline_with_refs(p, seg.text, profile, profile.body_font_size_pt)
                            else:
                                run = p.add_run(seg.text)
                                _set_run_font(run, profile, profile.body_font_size_pt)
                                runs = [run]
                            add_comment_to_runs(doc, p, runs, seg.comment_text)
                        else:  # pragma: no cover
                            raise BuildError(f"unknown review segment kind: {seg.kind}")
                if eref is not None:
                    etext = getattr(profile, "endnotes", {}).get(eref)
                    if etext is None:
                        raise BuildError(f"endnote [^{eref}] has no definition")
                    add_endnote(doc, p, "", etext)
        if ref is not None:
            ftext = getattr(profile, "footnotes", {}).get(ref)
            if ftext is None:
                raise BuildError(f"footnote [^{ref}] has no definition")
            add_footnote(doc, p, "", ftext)



def _numbered_paragraph_format(p, profile: Profile) -> None:
    """Формат абзаца под numPr: без ручного w:ind — отступы висячие приходят
    из w:lvl (numbering.xml), иначе pPr перебил бы нумерацию."""
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.body_line_spacing
    pf.space_after = Pt(0)
    pf.space_before = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _apply_num_pr(p, level: int, num_id: int) -> None:
    """ПАКЕТ 3: <w:numPr><w:ilvl/><w:numId/></w:numPr> в pPr параграфа."""
    p._p.get_or_add_pPr().insert_element_before(
        parse_xml(
            f'<w:numPr {NS}><w:ilvl w:val="{level}"/>'
            f'<w:numId w:val="{num_id}"/></w:numPr>'
        ),
        *_PPR_AFTER_NUMPR,
    )


def _lists_have_nesting(block) -> bool:
    """True, если среди items есть уровень >0 (плоские = GOST-легаси)."""
    return any(block.levels)


def _render_native_list(doc: Document, block, kind: str, profile: Profile,
                        start_at: int = 1) -> None:
    """ПАКЕТ 3: вложенный список — настоящая OpenXML-нумерация (numPr)."""
    from docx_cli.numbering import add_multilevel_list
    num_id = add_multilevel_list(doc, kind=kind, start=start_at)
    for item, level in zip(block.items, block.levels):
        p = doc.add_paragraph()
        _numbered_paragraph_format(p, profile)
        _apply_num_pr(p, level, num_id)
        run = p.add_run(item)
        _set_run_font(run, profile, profile.body_font_size_pt)


def render_numbered_list(doc: Document, block: NumberedListBlock, profile: Profile) -> None:
    if _lists_have_nesting(block):
        _render_native_list(doc, block, kind="decimal", profile=profile,
                            start_at=block.start_at)
        return
    for idx, item in enumerate(block.items, start=block.start_at):
        p = doc.add_paragraph()
        _body_paragraph_format(p, profile)
        # Explicit "1. Text" format (GOST requirement)
        text = f"{idx}. {item}"
        # Hanging indent: number stays on the left margin, text wraps aligned
        pf = p.paragraph_format
        pf.left_indent = Cm(1.25)
        pf.first_line_indent = Cm(-0.75)
        run = p.add_run(text)
        _set_run_font(run, profile, profile.body_font_size_pt)


def _borders_fragment(tag: str, edges: dict[str, tuple[str, int, str]]) -> str:
    """XML-фрагмент w:tblBorders/w:tcBorders; edge -> (val, sz, color), None = нет линии."""
    parts = []
    for edge in _BORDER_EDGES:
        spec = edges.get(edge)
        if spec is None:
            parts.append(f'<w:{edge} w:val="none" w:sz="0" w:space="0" w:color="auto"/>')
        else:
            val, sz, color = spec
            parts.append(f'<w:{edge} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>')
    return f'<{tag} {NS}>{"".join(parts)}</{tag}>'


def _set_table_borders(table: Table, edges: dict[str, tuple[str, int, str]]) -> None:
    tblPr = table._tbl.tblPr
    for old in tblPr.findall(qn("w:tblBorders")):
        tblPr.remove(old)
    tblPr.insert_element_before(
        parse_xml(_borders_fragment("w:tblBorders", edges)),
        "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook",
        "w:tblCaption", "w:tblDescription", "w:tblPrChange",
    )


def _set_cell_borders(cell, edges: dict[str, tuple[str, int, str]]) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:tcBorders")):
        tcPr.remove(old)
    tcPr.insert_element_before(
        parse_xml(_borders_fragment("w:tcBorders", edges)), *_TCPR_AFTER_BORDERS
    )


def _set_cell_shading(cell, fill: str) -> None:
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    tcPr.insert_element_before(
        parse_xml(f'<w:shd {NS} w:val="clear" w:color="auto" w:fill="{fill}"/>'),
        *_TCPR_AFTER_BORDERS,
    )


def _set_row_props(table: Table, row_idx: int, tbl_header: bool, cant_split: bool) -> None:
    row = table.rows[row_idx]
    trPr = row._tr.get_or_add_trPr()
    if cant_split and trPr.find(qn("w:cantSplit")) is None:
        trPr.append(parse_xml(f'<w:cantSplit {NS}/>'))
    if tbl_header and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(parse_xml(f'<w:tblHeader {NS}/>'))


_V_ALIGN = {"top": WD_ALIGN_VERTICAL.TOP, "center": WD_ALIGN_VERTICAL.CENTER,
            "bottom": WD_ALIGN_VERTICAL.BOTTOM}
_ALIGN = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}


def _hex_fill(color: str | None) -> str | None:
    """'#RRGGBB'/'RRGGBB' -> 'RRGGBB'; None -> None (дефолты pydantic не проходят валидаторы)."""
    return color.lstrip("#") if color else None


def _fill_cell(cell, raw, block: TableBlock, profile: Profile, header_row: bool) -> None:
    """Одна логическая ячейка: str или TableCell."""
    if isinstance(raw, TableCell):
        text, bold, italic, align, shading = raw.text, raw.bold, raw.italic, raw.align, raw.shading
    else:
        text, bold, italic, align, shading = raw, None, None, None, None
    if header_row and bold is None:
        bold = True  # легаси-контракт: шапка жирная

    p = cell.paragraphs[0]
    _cell_paragraph_format(p, profile)
    align_kind = align or ("center" if header_row else "left")
    p.paragraph_format.alignment = _ALIGN[align_kind]
    if "$" in text:
        _add_inline_text(p, text, profile, profile.cell_font_size_pt, italic=bool(italic))
        if bold:
            for r in p.runs:
                r.bold = True
    else:
        run = p.add_run(text)
        _set_run_font(run, profile, profile.cell_font_size_pt,
                      bold=bool(bold), italic=bool(italic))

    fill = _hex_fill(shading) or (_hex_fill(block.header_shading) if header_row else None)
    if fill:
        _set_cell_shading(cell, fill)
    cell.vertical_alignment = _V_ALIGN[block.vertical_align]


def _grid_widths_cm(block: TableBlock, profile: Profile) -> list[float] | None:
    """Ширины колонок в см: col_widths_cm > col_widths_pct > widths_cm."""
    if block.col_widths_cm is not None:
        return list(block.col_widths_cm)
    if block.col_widths_pct is not None:
        usable = profile.paper_width_cm - profile.margin_left_cm - profile.margin_right_cm
        return [usable * w / 100.0 for w in block.col_widths_pct]
    if block.widths_cm is not None:
        return list(block.widths_cm)
    return None


def _set_table_widths(table: Table, widths_cm: list[float]) -> None:
    """gridCol + ширина каждого физического tc (gridSpan -> суммарная)."""
    widths = [Cm(w) for w in widths_cm]
    for col_idx, width in enumerate(widths):
        table.columns[col_idx].width = width
    for row in table.rows:
        cursor = 0
        for tc in row._tr.tc_lst:
            tcPr = tc.get_or_add_tcPr()
            gs = tcPr.first_child_found_in("w:gridSpan")
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            tc.width = widths[cursor] if span == 1 else sum(widths[cursor:cursor + span])
            cursor += span


def render_table(doc: Document, block: TableBlock, profile: Profile) -> None:
    headers = block.headers
    rows = block.rows
    grid_width = len(headers)
    table: Table = doc.add_table(rows=1 + len(rows), cols=grid_width)
    if block.border_style == "grid":
        table.style = "Table Grid"
    else:
        # Без стиля: границы задаём явно (Table Grid вернул бы внутрилайны).
        pass
    table.autofit = False
    table.allow_autofit = False

    # Раскладка сетки (валидация уже прошла в spec): placement по data-строкам
    # [(col_start, span, raw)], rowspan-объединения через w:vMerge.
    placement = block.grid_layout()

    # Merge ДО заполнения: python-docx сам порождает gridSpan/vMerge restart/
    # continue; продолжения rowspan остаются физически в сетке пустыми.
    for sim_r, cells in enumerate(placement):
        row_idx = sim_r + 1
        for col_start, span, raw in cells:
            rowspan = raw.rowspan if isinstance(raw, TableCell) else 1
            if span > 1 or rowspan > 1:
                # Прямоугольное объединение: python-docx порождает w:gridSpan
                # (горизонталь) и w:vMerge restart/continue (вертикаль).
                table.cell(row_idx, col_start).merge(
                    table.cell(row_idx + rowspan - 1, col_start + span - 1)
                )

    # Шапка
    for col_idx, header in enumerate(headers):
        _fill_cell(table.cell(0, col_idx), header, block, profile, header_row=True)

    # Данные
    for sim_r, cells in enumerate(placement):
        row_idx = sim_r + 1
        for col_start, _span, raw in cells:
            _fill_cell(table.cell(row_idx, col_start), raw, block, profile, header_row=False)

    # Границы
    single = ("single", 4, "000000")
    if block.border_style == "grid":
        pass  # Table Grid уже рисует сетку
    elif block.border_style == "header_only":
        _set_table_borders(table, {e: None for e in _BORDER_EDGES})
        for col_idx in range(grid_width):
            _set_cell_borders(table.cell(0, col_idx), {"bottom": single})
    elif block.border_style == "boxed":
        _set_table_borders(table, {
            "top": single, "left": single, "bottom": single, "right": single,
        })
        for col_idx in range(grid_width):
            _set_cell_borders(table.cell(0, col_idx), {"bottom": single})
    else:  # "none"
        _set_table_borders(table, {e: None for e in _BORDER_EDGES})

    # Стабильность переноса: шапка повторяется, строки не рвутся между страницами.
    for row_idx in range(len(rows) + 1):
        _set_row_props(table, row_idx, tbl_header=(block.tbl_header and row_idx == 0),
                       cant_split=block.cant_split)

    # Ширины колонок
    widths = _grid_widths_cm(block, profile)
    if widths:
        _set_table_widths(table, widths)

    # Explicit cell margins: small padding, deterministic
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = tcPr.first_child_found_in("w:tcMar")
            if tcMar is None:
                tcMar = parse_xml(r'<w:tcMar xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                tcPr.append(tcMar)
            for edge, val in (("top", 100), ("start", 100), ("bottom", 100), ("end", 100)):
                node = tcMar.find(qn(f"w:{edge}"))
                if node is None:
                    node = parse_xml(f'<w:{edge} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:w="{val}" w:type="dxa"/>')
                    tcMar.append(node)
                else:
                    node.set(qn("w:w"), str(val))


def _clear_table_borders(table: Table) -> None:
    for row in table.rows:
        for cell in row.cells:
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = tcPr.first_child_found_in("w:tcBorders")
            if tcBorders is None:
                tcBorders = parse_xml(r'<w:tcBorders xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>')
                tcPr.append(tcBorders)
            for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
                node = tcBorders.find(qn(f"w:{edge}"))
                if node is None:
                    node = parse_xml(f'<w:{edge} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="none"/>')
                    tcBorders.append(node)
                else:
                    node.set(qn("w:val"), "none")


def render_bullet_list(doc: Document, block: BulletListBlock, profile: Profile) -> None:
    if _lists_have_nesting(block):
        _render_native_list(doc, block, kind="bullet", profile=profile)
        return
    for item in block.items:
        p = doc.add_paragraph()
        try:
            p.style = doc.styles["List Bullet"]
        except KeyError:
            # Fallback без стиля: ручной маркер с висячим отступом
            pf = p.paragraph_format
            pf.left_indent = Cm(1.25)
            pf.first_line_indent = Cm(-0.75)
            item = f"• {item}"
        _body_paragraph_format(p, profile)
        run = p.add_run(item)
        _set_run_font(run, profile, profile.body_font_size_pt)


def render_callout(doc: Document, block: CalloutBlock, profile: Profile) -> None:
    texts = [block.text] if isinstance(block.text, str) else list(block.text)
    accent = _hex_fill(block.accent) or "2B579A"
    fill = _hex_fill(block.shading)
    for text in texts:
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = profile.body_line_spacing
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.left_indent = Cm(1.0)
        pf.first_line_indent = Cm(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pPr = p._p.get_or_add_pPr()
        pPr.insert_element_before(
            parse_xml(
                f'<w:pBdr {NS}><w:left w:val="single" w:sz="24" w:space="4" '
                f'w:color="{accent}"/></w:pBdr>'
            ),
            *_PPR_AFTER_PBDR,
        )
        if fill:
            pPr.insert_element_before(
                parse_xml(f'<w:shd {NS} w:val="clear" w:color="auto" w:fill="{fill}"/>'),
                *_PPR_AFTER_SHD,
            )
        if block.inline_markdown:
            _add_inline_text(p, text, profile, profile.body_font_size_pt, italic=block.italic)
        else:
            run = p.add_run(text)
            _set_run_font(run, profile, profile.body_font_size_pt, italic=block.italic)


def render_page_break(doc: Document, block: PageBreakBlock, profile: Profile) -> None:
    doc.add_page_break()


def render_horizontal_rule(doc: Document, block: HorizontalRuleBlock, profile: Profile) -> None:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(6)
    pf.space_after = Pt(6)
    pf.first_line_indent = Cm(0)
    pPr = p._p.get_or_add_pPr()
    pPr.insert_element_before(
        parse_xml(
            f'<w:pBdr {NS}><w:bottom w:val="single" w:sz="4" w:space="0" '
            f'w:color="auto"/></w:pBdr>'
        ),
        *_PPR_AFTER_PBDR,
    )


def _add_field(p, instr: str, profile: Profile, size_pt: int, placeholder: str = "1") -> None:
    """Динамическое поле OpenXML (w:fldSimple) с кэшированным результатом."""
    font = profile.font_name
    sz = str(size_pt * 2)
    rpr = (
        f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" '
        f'w:eastAsia="{font}"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
    )
    p._p.append(parse_xml(
        f'<w:fldSimple {NS} w:instr="{instr}">'
        f'<w:r>{rpr}<w:t>{placeholder}</w:t></w:r>'
        f'</w:fldSimple>'
    ))


def render_footer(doc: Document, spec: DocSpec, profile: Profile) -> None:
    section = doc.sections[0]
    if spec.different_first_page:
        section.different_first_page_header_footer = True
        # Титульная страница: колонтитулы пустые (не наследуют основной).
        section.first_page_footer.is_linked_to_previous = False

    footer = section.footer
    footer.is_linked_to_previous = False

    if spec.footer_text:
        p = footer.paragraphs[0]
        _header_paragraph_format(p, profile)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = p.add_run(spec.footer_text)
        _set_run_font(run, profile, profile.body_font_size_pt)

    if spec.page_numbers:
        p = footer.add_paragraph() if spec.footer_text else footer.paragraphs[0]
        _header_paragraph_format(p, profile)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run("Стр. ")
        _set_run_font(run, profile, profile.body_font_size_pt)
        _add_field(p, " PAGE ", profile, profile.body_font_size_pt)
        run = p.add_run(" из ")
        _set_run_font(run, profile, profile.body_font_size_pt)
        _add_field(p, " NUMPAGES ", profile, profile.body_font_size_pt)


def render_signature_block(doc: Document, block: SignatureBlock, profile: Profile) -> None:
    fmt = block.format if block.format is not None else profile.default_signature_format

    if fmt == "text":
        p = doc.add_paragraph()
        pf = p.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.line_spacing = profile.body_line_spacing
        pf.space_after = Pt(0)
        pf.space_before = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pf.tab_stops.add_tab_stop(Cm(16.5), WD_TAB_ALIGNMENT.RIGHT)
        run_left = p.add_run(block.left)
        _set_run_font(run_left, profile, profile.body_font_size_pt)
        run_tab = p.add_run("\t")
        _set_run_font(run_tab, profile, profile.body_font_size_pt)
        run_right = p.add_run(block.right)
        _set_run_font(run_right, profile, profile.body_font_size_pt)
        return

    table: Table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    table.allow_autofit = False
    table.style = "Table Grid"
    left_cell, right_cell = table.rows[0].cells

    left_cell.text = ""
    p = left_cell.paragraphs[0]
    _cell_paragraph_format(p, profile)
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = p.add_run(block.left)
    _set_run_font(run, profile, profile.cell_font_size_pt)

    right_cell.text = ""
    p2 = right_cell.paragraphs[0]
    _cell_paragraph_format(p2, profile)
    p2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run2 = p2.add_run(block.right)
    _set_run_font(run2, profile, profile.cell_font_size_pt)

    # Hide grid lines by clearing borders for signature block
    _clear_table_borders(table)


def render_pleading_marker(doc: Document, block: PleadingBlock, profile: Profile) -> None:
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    pf.line_spacing = profile.body_line_spacing
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Cm(0)
    run = p.add_run(block.text)
    _set_run_font(run, profile, profile.body_font_size_pt, bold=True)


def render_attachments_list(doc: Document, block: AttachmentsListBlock, profile: Profile) -> None:
    p = doc.add_paragraph()
    _body_paragraph_format(p, profile)
    run = p.add_run("Приложения:")
    _set_run_font(run, profile, profile.body_font_size_pt, bold=True)

    # S2-H2: каждое приложение получает явный маркер «Приложение N» —
    # без него ссылки «см. Приложение 1» в теле висят (осиротевшие
    # приложения), а перечень превращается в безымянный список.
    for i, item in enumerate(block.items, start=1):
        marker = doc.add_paragraph()
        marker.alignment = WD_ALIGN_PARAGRAPH.CENTER
        _body_paragraph_format(marker, profile)
        mrun = marker.add_run(f"Приложение {i}")
        _set_run_font(mrun, profile, profile.body_font_size_pt, bold=True)

        p = doc.add_paragraph()
        _body_paragraph_format(p, profile)
        pf = p.paragraph_format
        pf.left_indent = Cm(1.25)
        pf.first_line_indent = Cm(-0.75)
        run = p.add_run(item)
        _set_run_font(run, profile, profile.body_font_size_pt)


def render_toc(doc: Document, block: TocBlock, profile: Profile) -> None:
    """ПАКЕТ 1: заголовок «Содержание» + нативное поле TOC (w:fldSimple).

    Заголовок — центрированный bold (не Heading 1, чтобы Word не вносил сам
    оглавление в само себя). Поле обновляется Word'ом на открытии:
    settings.xml получает w:updateFields (ensure_update_fields в builder).
    """
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.space_before = Pt(12)
    pf.space_after = Pt(6)
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(block.title)
    _set_run_font(run, profile, _HEADING_FMT[1][0], bold=True)

    fp = doc.add_paragraph()
    pPr = fp._p.get_or_add_pPr()
    pPr.append(parse_xml(f'<w:pStyle {NS} w:val="TOC1"/>'))
    fld = parse_xml(f'<w:fldSimple {NS}/>')
    fld.set(qn("w:instr"), f'TOC \\o "1-{block.levels}" \\h \\z \\u')
    fp._p.append(fld)


def render_image(doc: Document, block: ImageBlock, profile: Profile) -> None:
    """ПАКЕТ 2: изображение DrawingML с авто-масштабированием и подписью.

    Fail-closed: файла нет -> BuildError, битого архива не будет.
    Ширина: width_cm (cap по рабочей области) > width_pct > авто
    (натуральная ширина картинки, если влезает, иначе вся рабочая область).
    """
    if not os.path.isfile(block.path):
        raise BuildError(f"image file not found: {block.path} — fail-closed")

    section = doc.sections[-1]
    avail_width_cm = (
        section.page_width.cm - section.left_margin.cm - section.right_margin.cm
    )

    if block.width_cm is not None:
        target_width_cm = min(block.width_cm, avail_width_cm)
    elif block.width_pct is not None:
        target_width_cm = avail_width_cm * (block.width_pct / 100.0)
    else:
        from docx.image.image import Image as _Img
        try:
            natural_cm = _Img.from_file(block.path).width.cm
        except Exception:
            natural_cm = avail_width_cm
        target_width_cm = min(natural_cm, avail_width_cm)

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
    }
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = align_map[block.align]
    pf.space_before = Pt(6)
    pf.space_after = Pt(2) if block.caption else Pt(6)
    pf.first_line_indent = Cm(0)
    run = p.add_run()
    run.add_picture(block.path, width=Cm(target_width_cm))

    if block.caption:
        cap = doc.add_paragraph()
        cf = cap.paragraph_format
        cf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cf.first_line_indent = Cm(0)
        cf.space_before = Pt(0)
        cf.space_after = Pt(12)
        cap_run = cap.add_run(block.caption)
        _set_run_font(cap_run, profile, profile.body_font_size_pt - 1, italic=True)


def render_section_break(doc: Document, block: SectionBreakBlock, profile: Profile) -> None:
    """ПАКЕТ 1: новый раздел через w:sectPr с реальной сменой ориентации.

    python-docx add_section клонирует свойства предыдущего раздела, поэтому
    landscape↔portrait требует обмена page_width/page_height (условного:
    после landscape-раздела portrait-раздел обязан вернуть геометрию).
    """
    start = WD_SECTION.CONTINUOUS if block.break_type == "continuous" else WD_SECTION.NEW_PAGE
    section = doc.add_section(start)
    section.orientation = (
        WD_ORIENTATION.LANDSCAPE if block.orientation == "landscape" else WD_ORIENTATION.PORTRAIT
    )
    width, height = section.page_width, section.page_height
    dims_landscape = width > height
    if (block.orientation == "landscape") != dims_landscape:
        section.page_width, section.page_height = height, width

    margins = (
        ("left", block.margin_left_cm, profile.margin_left_cm),
        ("right", block.margin_right_cm, profile.margin_right_cm),
        ("top", block.margin_top_cm, profile.margin_top_cm),
        ("bottom", block.margin_bottom_cm, profile.margin_bottom_cm),
    )
    for attr, block_cm, profile_cm in margins:
        setattr(section, f"{attr}_margin", Cm(block_cm if block_cm is not None else profile_cm))

    # ПАКЕТ 9: колонки w:cols в sectPr
    if block.cols > 1:
        sectPr = section._sectPr
        w_ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        for old_cols in sectPr.findall(f"{{{w_ns}}}cols"):
            sectPr.remove(old_cols)
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls
        space_twips = int(block.col_space_cm * 566.929)
        sep_val = "1" if block.col_sep else "0"
        cols_xml = (
            f'<w:cols {nsdecls("w")} w:num="{block.cols}" '
            f'w:space="{space_twips}" w:equalWidth="1" w:sep="{sep_val}"/>'
        )
        cols_el = parse_xml(cols_xml)
        # S1-M1: CT_SectPr требует w:cols ДО w:docsGrid — иначе диалог
        # восстановления Word на стороне получателя.
        docgrid = sectPr.find(f"{{{w_ns}}}docGrid")
        if docgrid is not None:
            docgrid.addprevious(cols_el)
        else:
            sectPr.append(cols_el)


# Позиция w:updateFields в CT_Settings: все элементы, идущие за ним по схеме.
_SETTINGS_AFTER_UPDATE_FIELDS = (
    "w:hdrShapeDefaults", "w:footnotePr", "w:endnotePr", "w:compat", "w:docVars",
    "w:rsids", "w:attachedSchema", "w:themeFontLang", "w:clrSchemeMapping",
    "w:doNotIncludeSubdocsInStats", "w:doNotAutoCompressPictures", "w:forceUpgrade",
    "w:captions", "w:readModeInkLockDown", "w:smartTagType", "w:shapeDefaults",
    "w:doNotEmbedSmartTags", "w:decimalSymbol", "w:listSeparator",
)


def ensure_update_fields(doc: Document) -> None:
    """ПАКЕТ 1: <w:updateFields w:val="true"/> в word/settings.xml —
    Word обновляет TOC/поля страниц при открытии документа."""
    settings = doc.settings.element
    if settings.find(qn("w:updateFields")) is not None:
        return
    settings.insert_element_before(
        parse_xml(f'<w:updateFields {NS} w:val="true"/>'),
        *_SETTINGS_AFTER_UPDATE_FIELDS,
    )

