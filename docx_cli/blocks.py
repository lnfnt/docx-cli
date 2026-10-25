"""Block-level rendering helpers for the DOCX builder."""
from __future__ import annotations

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.table import Table

from docx_cli.spec import (
    AttachmentsListBlock,
    HeaderSpec,
    NumberedListBlock,
    ParagraphBlock,
    PleadingBlock,
    Profile,
    SignatureBlock,
    TableBlock,
    TitleSpec,
)


# --------------------------------------------------------------------------------------
# Low-level helpers
# --------------------------------------------------------------------------------------

def _set_run_font(run, profile: Profile, size_pt: int, bold: bool = False, italic: bool = False) -> None:
    run.font.name = profile.font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    run.font.rtl = False
    run._element.rPr.rFonts.set(qn("w:eastAsia"), profile.font_name)


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

def _split_inline(text: str) -> list[tuple[str, bool, bool]]:
    """Parse **bold** and *italic* markers, returning (text, bold, italic) segments."""
    parts: list[tuple[str, bool, bool]] = []
    i = 0
    bold = False
    italic = False
    buf: list[str] = []
    n = len(text)
    while i < n:
        if text.startswith("**", i):
            parts.append(("".join(buf), bold, italic))
            buf = []
            bold = not bold
            i += 2
        elif text.startswith("*", i):
            parts.append(("".join(buf), bold, italic))
            buf = []
            italic = not italic
            i += 1
        elif text.startswith("\\", i) and i + 1 < n:
            buf.append(text[i + 1])
            i += 2
        else:
            buf.append(text[i])
            i += 1
    parts.append(("".join(buf), bold, italic))
    # Drop empty leading segment
    if parts and parts[0][0] == "":
        parts.pop(0)
    return parts


def _add_inline_text(p, text: str, profile: Profile, size_pt: int) -> None:
    for segment, bold, italic in _split_inline(text):
        run = p.add_run(segment)
        _set_run_font(run, profile, size_pt, bold=bold, italic=italic)


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
    p = doc.add_paragraph()
    _title_paragraph_format(p, profile)
    run = p.add_run(spec.main)
    _set_run_font(run, profile, profile.title_font_size_pt, bold=True)
    if spec.subtitle:
        p2 = doc.add_paragraph()
        _title_paragraph_format(p2, profile)
        run2 = p2.add_run(spec.subtitle)
        _set_run_font(run2, profile, profile.body_font_size_pt, bold=False)


def render_paragraph(doc: Document, block: ParagraphBlock, profile: Profile) -> None:
    p = doc.add_paragraph()
    _body_paragraph_format(p, profile)
    text = block.text
    if block.inline_markdown:
        _add_inline_text(p, text, profile, profile.body_font_size_pt)
    else:
        run = p.add_run(text)
        _set_run_font(run, profile, profile.body_font_size_pt)


def render_numbered_list(doc: Document, block: NumberedListBlock, profile: Profile) -> None:
    for idx, item in enumerate(block.items, start=1):
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


def render_table(doc: Document, block: TableBlock, profile: Profile) -> None:
    headers = block.headers
    rows = block.rows
    col_count = len(headers) if headers else len(rows[0])
    table: Table = doc.add_table(rows=1 + len(rows), cols=col_count)
    table.style = "Table Grid"
    table.autofit = False
    table.allow_autofit = False

    # Header row
    hdr_cells = table.rows[0].cells
    for col_idx, header in enumerate(headers):
        cell = hdr_cells[col_idx]
        cell.text = ""
        p = cell.paragraphs[0]
        _cell_paragraph_format(p, profile)
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        _set_run_font(run, profile, profile.cell_font_size_pt, bold=True)

    # Data rows
    for row_idx, row_data in enumerate(rows, start=1):
        cells = table.rows[row_idx].cells
        for col_idx, value in enumerate(row_data):
            cell = cells[col_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            _cell_paragraph_format(p, profile)
            p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            run = p.add_run(value)
            _set_run_font(run, profile, profile.cell_font_size_pt)

    # Column widths
    if block.widths_cm:
        for row in table.rows:
            for col_idx, width_cm in enumerate(block.widths_cm):
                if col_idx < col_count:
                    row.cells[col_idx].width = Cm(width_cm)

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

    for item in block.items:
        p = doc.add_paragraph()
        _body_paragraph_format(p, profile)
        pf = p.paragraph_format
        pf.left_indent = Cm(1.25)
        pf.first_line_indent = Cm(-0.75)
        run = p.add_run(item)
        _set_run_font(run, profile, profile.body_font_size_pt)

