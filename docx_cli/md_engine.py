# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""docx-cli md engine v2 — Фазы 2/3 + ПАКЕТ 1: Markdown для агентов -> DOCX.

Отличие от md_court (процессуальный fail-closed рендер с титулами ИСК/ОТЗЫВ):
этот движок — для обычной деловой документации. Разбирает:
  - заголовки `#`..`####` -> HeadingBlock (Заголовок 1-4, 18/16/14/12 pt);
  - GFM pipe-таблицы `| a | b |` + `|:---|---:|` -> TableBlock (шапка #F2F2F2,
    выравнивание колонок по двоеточиям; border_style берётся из профиля);
  - выноски `> [!NOTE]` (GitHub/notes-app) и простые цитаты `> ...` -> CalloutBlock;
  - сноски `[^1]` + определения `[^1]: ...` -> нативные OpenXML-сноски
    (word/footnotes.xml + w:footnoteReference) через docx_cli.footnotes;
  - `[TOC]` / `[[toc]]` -> TocBlock (нативное поле Word TOC + w:updateFields);
   - `<!-- section:landscape -->` / `<!-- section:portrait -->` ->
     SectionBreakBlock (реальная смена ориентации страницы);
   - списки `- ` / `1. ` (без вложенности — плоские бизнес-списки);
   - ПАКЕТ 4: разметка рецензирования (CriticMarkup) в обычных абзацах:
     `{++добавлено++}` -> w:ins, `{--удалено--}` -> w:del/w:delText,
     `[фрагмент]{comment: "замечание"}` -> нативный комментарий
     (word/comments.xml + w:commentRangeStart/End + w:commentReference).
Первый `# H1` становится титулом документа. Рендер идёт через DocSpec ->
builder.build_document, поэтому весь профильный конвейер (corporate/report,
валидация, колонтитулы) работает без дублирования.

Fail-closed: пустой md / нет H1 / битая таблица / маркер сноски вне абзаца /
сноска без определения или определение без ссылки -> ошибка, документ не создан.
"""
from __future__ import annotations

import ast
import re
from pathlib import Path
from typing import NamedTuple

from pydantic import ValidationError

from docx_cli import builder
from docx_cli.errors import ProfileError
from docx_cli.footnotes import FOOTNOTE_MARKER_RE
from docx_cli.endnotes import ENDNOTE_MARKER_RE
from docx_cli.spec import (
    BulletListBlock,
    CalloutBlock,
    CheckboxBlock,
    DatePickerBlock,
    DocSpec,
    DropdownBlock,
    HeaderSpec,
    HeadingBlock,
    HorizontalRuleBlock,
    ImageBlock,
    MathBlock,
    NumberedListBlock,
    OutputSpec,
    ParagraphBlock,
    ProfileName,
    SectionBreakBlock,
    TableBlock,
    TableCell,
    TitleSpec,
    TocBlock,
)

# GFM-разделитель таблицы (совместим с md_court._TABLE_SEP_RE).
_TABLE_SEP_RE = re.compile(
    r"^[ \t]*\|?[ \t]*:?-{1,}:?[ \t]*(?:\|[ \t]*:?-{1,}:?[ \t]*)*\|?[ \t]*$")
_HEADING_RE = re.compile(r"^(#{1,4})\s+(.+)$")
_LIST_ITEM_RE = re.compile(r"^(\s*)([-*]|\d+[.)])\s+(.+)$")
# ПАКЕТ 6: `- [x] Пункт А` / `- [ ] Пункт Б` -> нативные SDT-чекбоксы.
_CHECKBOX_MD_RE = re.compile(r"^[-*]\s+\[([xX ])\]\s+(\S.*)$")
# ПАКЕТ 6: {{dropdown: ["Да", "Нет", "Не применимо"]}} и {{date}} отдельной строкой.
_DROPDOWN_MD_RE = re.compile(r"^\{\{\s*dropdown\s*:\s*(.+?)\s*\}\}$")
_DATE_MD_RE = re.compile(r"^\{\{\s*date\s*\}\}$")
_CALLOUT_MARK_RE = re.compile(r"^\[\!([A-Za-z]+)\]$")
_ENDNOTE_DEF_RE = re.compile(r"^\[\^(?:!|e:)(\w+)\]:\s*(.+)$")
_FOOTNOTE_DEF_RE = re.compile(r"^\[\^(\w+)\]:\s*(.+)$")
_FOOTNOTE_REF_RE = re.compile(r"\[\^(\w+)\]")
_TOC_RE = re.compile(r"^(?:\[TOC\]|\[\[toc\]\])$", re.IGNORECASE)
_SECTION_RE = re.compile(
    r"^<!--\s*section:(portrait|landscape)(?::(next_page|continuous))?(?::columns:(\d+))?\s*-->$")
_COLUMNS_ONLY_RE = re.compile(
    r"^<!--\s*section:columns:(\d+)\s*-->$")
_HR_RE = re.compile(r"^(?:-{3,}|\*{3,})\s*$")
# ПАКЕТ 2: `![caption](path)` / `![](path)` — отдельной строкой.
_IMAGE_RE = re.compile(r"^!\[([^\]]*)\]\((.+)\)$")


def _indent_level(ws: str) -> int:
    """ПАКЕТ 3: уровень списка по ведущим пробелам/табам.

    0 -> уровень 0; каждый уровень ~ 2-4 пробела или 1 таб: 2-3 -> 1,
    4-6 -> 2, глубже -> clamp 2 (numbering.xml определяет уровни 0..2).
    В примере ТЗ: 0/3/6 пробелов -> уровни 0/1/2.
    """
    level = (len(ws.replace("\t", "  ")) + 2) // 3
    return min(level, 2)

# Выноски GitHub/notes-app: тип -> accent (HEX без '#').
_CALLOUT_ACCENTS = {
    "note": "0969DA",
    "tip": "1A7F37",
    "warning": "9A6700",
    "important": "8250DF",
    "caution": "D1242F",
}


class MdEngineError(Exception):
    """Fail-closed ошибка сборки md v2 -> docx."""


# -------------------------------------------------------------------------------------
# ПАКЕТ 4: разметка рецензирования (CriticMarkup) — парсится в тексте обычных
# абзацев; отрезки интерпретирует blocks.render_paragraph (doc доступен там).
# -------------------------------------------------------------------------------------

class ReviewSegment(NamedTuple):
    """Отрезок текста абзаца: kind = "text" | "ins" | "del" | "comment"."""

    kind: str
    text: str
    comment_text: str | None = None


# {++добавлено++} | {--удалено--} | [фрагмент]{comment: "замечание"}
_REVIEW_SEGMENT_RE = re.compile(
    r"\{\+\+(?P<ins>.+?)\+\+\}"
    r"|\{--(?P<del>.+?)--\}"
    r"|\[(?P<target>[^\]]+)\]\{comment:\s*(?P<q>[\"'])(?P<comment>.*?)(?P=q)\}",
    re.DOTALL,
)


def split_review_segments(text: str) -> list[ReviewSegment]:
    """ПАКЕТ 4: текст абзаца -> [ReviewSegment]: разметка рецензирования ->
    отдельные отрезки, остальное — kind="text". Незакрытые маркеры остаются
    литеральным текстом (как и в исходном CriticMarkup)."""
    segments: list[ReviewSegment] = []
    pos = 0
    for m in _REVIEW_SEGMENT_RE.finditer(text):
        if m.start() > pos:
            segments.append(ReviewSegment("text", text[pos:m.start()]))
        if m.group("ins") is not None:
            segments.append(ReviewSegment("ins", m.group("ins")))
        elif m.group("del") is not None:
            segments.append(ReviewSegment("del", m.group("del")))
        else:
            segments.append(ReviewSegment("comment", m.group("target"), m.group("comment")))
        pos = m.end()
    if pos < len(text) or not segments:
        segments.append(ReviewSegment("text", text[pos:]))
    return segments


def _align_from_sep(cell: str) -> str | None:
    left, right = cell.startswith(":"), cell.endswith(":")
    if left and right:
        return "center"
    if right:
        return "right"
    if left:
        return "left"
    return None


def _split_row(raw: str) -> list[str]:
    t = raw.strip()
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|") and not t.endswith("\\|"):
        t = t[:-1]
    return [p.strip().replace("\\|", "|") for p in re.split(r"(?<!\\)\|", t)]


class _Parser:
    def __init__(self, md_text: str, profile_name: ProfileName,
                 base_dir: str | None = None) -> None:
        self.lines = md_text.replace("\ufeff", "").split("\n")
        self.profile_name = profile_name
        self.base_dir = base_dir
        self.title: str | None = None
        self.footnotes: dict[str, str] = {}
        self.footnote_refs: list[str] = []
        self.endnotes: dict[str, str] = {}
        self.endnote_refs: list[str] = []
        self.blocks: list = []
        self.i = 0

    def parse(self) -> None:
        n = len(self.lines)
        while self.i < n:
            line = self.lines[self.i].strip()
            if not line:
                self.i += 1
                continue
            if _TOC_RE.match(line):
                self._push_block(TocBlock(type="toc"))
                self.i += 1
                continue
            if m := _SECTION_RE.match(line):
                cols = int(m.group(3)) if m.group(3) else 1
                self._push_block(SectionBreakBlock(
                    type="section_break", orientation=m.group(1),
                    break_type=m.group(2) or "next_page",
                    cols=cols))
                self.i += 1
                continue
            if m := _COLUMNS_ONLY_RE.match(line):
                self._push_block(SectionBreakBlock(
                    type="section_break", orientation="portrait",
                    break_type="continuous",
                    cols=int(m.group(1))))
                self.i += 1
                continue
            if m := _HEADING_RE.match(line):
                # Первый `# H1` — титул; остальные заголовки — блоки.
                # S2-M1: маркер сноски [^id] в заголовке раньше молча уходил
                # в docx литералом (двойной стандарт против _buffer_text в
                # ячейках/списках). Теперь — тот же честный fail-closed.
                heading_text = self._buffer_text(m.group(2).strip())
                if self.title is None and len(m.group(1)) == 1:
                    self.title = heading_text
                else:
                    self._push_block(HeadingBlock(
                        type="heading", level=len(m.group(1)),
                        text=heading_text))
                self.i += 1
                continue
            if line.startswith("$$"):
                self._parse_math_block()
                continue
            if "|" in line and self._is_table_start():
                self._parse_table()
                continue
            # S2-M4: backslash-escape line-конструкций — «\>» и «\{{» остаются
            # обычным текстом, не перехватываются callout/dropdown-парсером.
            if line.startswith(("\\>", "\\{{")):
                self._push_block(ParagraphBlock(type="paragraph", text=line[1:]))
                self.i += 1
                continue
            if line.startswith(">"):
                self._parse_callout()
                continue
            if _HR_RE.match(line):
                self._push_block(HorizontalRuleBlock(type="horizontal_rule"))
                self.i += 1
                continue
            if _CHECKBOX_MD_RE.match(line):
                self._parse_checkbox_run()
                continue
            if m := _DROPDOWN_MD_RE.match(line):
                try:
                    items = ast.literal_eval(m.group(1))
                except (ValueError, SyntaxError) as exc:
                    raise MdEngineError(f"dropdown syntax error: {exc}") from exc
                if (not isinstance(items, list) or not items
                        or not all(isinstance(x, str) for x in items)):
                    raise MdEngineError("dropdown items must be a non-empty list of strings")
                self._push_block(DropdownBlock(type="dropdown", items=items))
                self.i += 1
                continue
            if _DATE_MD_RE.match(line):
                self._push_block(DatePickerBlock(type="date_picker"))
                self.i += 1
                continue
            if m := _IMAGE_RE.match(line):
                caption, img_path = m.group(1).strip(), m.group(2).strip()
                p = Path(img_path)
                if not p.is_absolute() and self.base_dir:
                    p = Path(self.base_dir) / p
                self._push_block(ImageBlock(
                    type="image", path=str(p), caption=caption or None))
                self.i += 1
                continue
            if m := _LIST_ITEM_RE.match(line):
                self._parse_list_run()
                continue
            if m := _ENDNOTE_DEF_RE.match(line):
                eid = m.group(1)
                if eid in self.endnotes:
                    raise MdEngineError(f"endnote [^{eid}] defined twice — fail-closed")
                self.endnotes[eid] = m.group(2).strip()
                self.i += 1
                continue
            if m := _FOOTNOTE_DEF_RE.match(line):
                fid = m.group(1)
                if fid in self.footnotes:
                    raise MdEngineError(f"footnote [^{fid}] defined twice — fail-closed")
                self.footnotes[fid] = m.group(2).strip()
                self.i += 1
                continue
            self._parse_paragraph()
        self._validate_footnotes()
        self._validate_endnotes()

    # --- конструкции -------------------------------------------------------

    def _buffer_text(self, text: str) -> str:
        # Сноски — только в обычных абзацах; в ячейках/списках/цитатах
        # молча выкидывать маркер (прежний суррогат) нельзя — fail-closed.
        if _FOOTNOTE_REF_RE.search(text) or ENDNOTE_MARKER_RE.search(text):
            raise MdEngineError(
                f"footnote/endnote marker {text.strip()!r}: [^id]/[^!id] поддерживается только "
                "в обычных абзацах — fail-closed")
        return text

    def _push_block(self, block) -> None:
        self.blocks.append(block)

    def _is_table_start(self) -> bool:
        if self.i + 1 >= len(self.lines):
            return False
        return bool(_TABLE_SEP_RE.match(self.lines[self.i + 1].strip()))

    def _parse_math_block(self) -> None:
        """ПАКЕТ 7: блочная формула $$ ... $$ (однострочная или многострочная)."""
        first_line = self.lines[self.i].strip()
        # Проверяем однострочный вид: $$ latex $$
        if first_line.startswith("$$") and len(first_line) > 2 and first_line.endswith("$$"):
            latex = first_line[2:-2].strip()
            self.i += 1
            if not latex:
                raise MdEngineError("empty math block $$ $$ (fail-closed)")
            self._push_block(MathBlock(type="math", latex=latex))
            return

        if first_line == "$$" or (first_line.startswith("$$") and not first_line.endswith("$$")):
            # Многострочный вид
            lines = []
            if len(first_line) > 2:
                lines.append(first_line[2:].strip())
            self.i += 1
            closed = False
            while self.i < len(self.lines):
                line = self.lines[self.i].strip()
                if line == "$$":
                    closed = True
                    self.i += 1
                    break
                elif line.endswith("$$"):
                    closed = True
                    lines.append(line[:-2].strip())
                    self.i += 1
                    break
                lines.append(self.lines[self.i])
                self.i += 1
            if not closed:
                raise MdEngineError("unclosed math block $$ (fail-closed)")
            latex = "\n".join(lines).strip()
            if not latex:
                raise MdEngineError("empty math block $$ $$ (fail-closed)")
            self._push_block(MathBlock(type="math", latex=latex))
            return

        raise MdEngineError(f"invalid math block syntax at {first_line!r}")

    def _parse_table(self) -> None:
        # S2-M1: шапка раньше шла мимо _buffer_text — маркер сноски в заголовке
        # колонки уходил в docx литералом, а в теле той же таблицы падал
        # fail-closed. Единый стандарт: заголовки проходят тот же гейт.
        headers_raw = [self._buffer_text(c) for c in _split_row(self.lines[self.i])]
        aligns = [_align_from_sep(c) for c in _split_row(self.lines[self.i + 1])]
        ncols = len(headers_raw)
        aligns = (aligns + [None] * ncols)[:ncols]
        body: list[list[TableCell]] = []
        j = self.i + 2
        while j < len(self.lines):
            row = self.lines[j].strip()
            if not row or "|" not in row or _TABLE_SEP_RE.match(row):
                break
            cells = _split_row(row)
            # S2-H1: поведение GFM-reference — выравнивание по шапке без
            # kill-документа: лишние ячейки отсекаются, недостающие падают
            # пустыми. Прежний fail-closed на несовпадении числа колонок
            # гасил всю сборку из-за одного неэкранированного «|» в тексте.
            if len(cells) > ncols:
                cells = cells[:ncols]
            elif len(cells) < ncols:
                cells = cells + [""] * (ncols - len(cells))
            body.append([
                TableCell(text=self._buffer_text(c), align=aligns[k])
                for k, c in enumerate(cells)
            ])
            j += 1
        if not body:
            raise MdEngineError("table has no body rows (fail-closed)")
        try:
            self._push_block(TableBlock(
                type="table", headers=headers_raw, rows=body,
                border_style=self._table_border()))
        except ValidationError as exc:
            raise MdEngineError(f"table error: {exc}") from exc
        self.i = j

    def _table_border(self) -> str:
        from docx_cli.spec import Profile
        try:
            return Profile(self.profile_name).default_table_border
        except ProfileError:
            return "grid"

    def _parse_callout(self) -> None:
        lines: list[str] = []
        accent: str | None = None
        while self.i < len(self.lines) and self.lines[self.i].strip().startswith(">"):
            content = self.lines[self.i].strip()[1:].strip()
            self.i += 1
            if not accent and (m := _CALLOUT_MARK_RE.match(content)):
                accent = _CALLOUT_ACCENTS.get(m.group(1).lower())
                continue
            if not content:
                continue
            lines.append(self._buffer_text(content))
        if not accent:
            accent = "2B579A"  # цитата без маркера: дефолт CalloutBlock
        if not lines:
            raise MdEngineError("callout/quote has no text (fail-closed)")
        self._push_block(CalloutBlock(
            type="callout", text=lines[0] if len(lines) == 1 else lines,
            accent=accent))

    def _parse_checkbox_run(self) -> None:
        """ПАКЕТ 6: рун GFM-чекбоксов `- [x] ...` — по одному CheckboxBlock."""
        while self.i < len(self.lines) and (m := _CHECKBOX_MD_RE.match(self.lines[self.i].strip())):
            self._push_block(CheckboxBlock(
                type="checkbox", checked=m.group(1).lower() == "x",
                label=self._buffer_text(m.group(2).strip())))
            self.i += 1

    def _parse_list_run(self) -> None:
        """ПАКЕТ 3: списки с уровнями по отступам (2/4 пробела или таб);
        смена вида (ol <-> ul) режет рун на соседние блоки."""
        runs: list[tuple[bool, list[str], list[int]]] = []
        kind: bool | None = None
        items: list[str] = []
        levels: list[int] = []
        while (self.i < len(self.lines)
               and (im := _LIST_ITEM_RE.match(self.lines[self.i].rstrip()))):
            is_ordered = bool(re.match(r"^\d+[.)]$", im.group(2)))
            if kind is None:
                kind = is_ordered
            elif is_ordered != kind:
                runs.append((kind, items, levels))
                kind, items, levels = is_ordered, [], []
            items.append(self._buffer_text(im.group(3)))
            levels.append(_indent_level(im.group(1)))
            self.i += 1
        if kind is None:
            raise MdEngineError("empty list run (fail-closed)")
        runs.append((kind, items, levels))
        for is_ordered, its, lvs in runs:
            if is_ordered:
                self._push_block(NumberedListBlock(
                    type="numbered_list", items=its, levels=lvs))
            else:
                self._push_block(BulletListBlock(
                    type="bullet_list", items=its, levels=lvs))

    def _parse_paragraph(self) -> None:
        buf: list[str] = []
        n = len(self.lines)
        while self.i < n:
            line = self.lines[self.i].strip()
            if (not line or _HEADING_RE.match(line) or line.startswith(">")
                    or "|" in line or _LIST_ITEM_RE.match(line)
                    or _IMAGE_RE.match(line)
                    or _ENDNOTE_DEF_RE.match(line)
                    or _FOOTNOTE_DEF_RE.match(line) or _HR_RE.match(line)
                    or _TOC_RE.match(line) or _SECTION_RE.match(line)
                    or _CHECKBOX_MD_RE.match(line)
                    or _DROPDOWN_MD_RE.match(line) or _DATE_MD_RE.match(line)
                    or line.startswith("$$")):  # ПАКЕТ 7: $$-блок рвёт абзац
                break
            buf.append(line)
            self.i += 1
        text = " ".join(buf)
        if not text:
            raise MdEngineError("empty paragraph (fail-closed)")
        # Маркеры [^id] и [^!id] остаются в тексте — render_paragraph связывает их
        # с нативными сносками (footnotes.xml / endnotes.xml).
        self.footnote_refs.extend(m.group(1) for m in FOOTNOTE_MARKER_RE.finditer(text))
        self.endnote_refs.extend(m.group(1) for m in ENDNOTE_MARKER_RE.finditer(text))
        self._push_block(ParagraphBlock(type="paragraph", text=text))

    def _validate_footnotes(self) -> None:
        """Fail-closed: каждая ссылка имеет определение и наоборот."""
        if not self.footnotes and not self.footnote_refs:
            return
        missing = sorted(set(self.footnote_refs) - self.footnotes.keys())
        if missing:
            raise MdEngineError(
                f"footnote [^{missing[0]}] has no definition — fail-closed")
        orphan = sorted(self.footnotes.keys() - set(self.footnote_refs))
        if orphan:
            raise MdEngineError(
                f"footnote definition [^{orphan[0]}] never referenced — fail-closed")

    def _validate_endnotes(self) -> None:
        """Fail-closed: каждая концевая сноска имеет определение и наоборот."""
        if not self.endnotes and not self.endnote_refs:
            return
        missing = sorted(set(self.endnote_refs) - self.endnotes.keys())
        if missing:
            raise MdEngineError(
                f"endnote [^{missing[0]}] has no definition — fail-closed")
        orphan = sorted(self.endnotes.keys() - set(self.endnote_refs))
        if orphan:
            raise MdEngineError(
                f"endnote definition [^{orphan[0]}] never referenced — fail-closed")



def markdown_to_spec(md_text: str, output_path: str, author: str,
                     profile_name: ProfileName = "corporate",
                     base_dir: str | None = None) -> DocSpec:
    """md-текст -> DocSpec (титул = первый `# H1`, fail-closed без него)."""
    parser = _Parser(md_text, profile_name, base_dir=base_dir)
    parser.parse()
    if not md_text.strip():
        raise MdEngineError("empty md input — fail-closed")
    if parser.title is None:
        raise MdEngineError("нет заголовка `# H1` для титула — fail-closed")
    try:
        spec = DocSpec(
            profile=profile_name,
            output=OutputSpec(file_name=Path(output_path).name, author=author),
            header=HeaderSpec(lines=[f"Подготовлено: {author}"]),
            title=TitleSpec(main=parser.title),
            body=parser.blocks,
            footnotes=parser.footnotes,
            endnotes=parser.endnotes,
        )
    except ValidationError as exc:
        raise MdEngineError(f"doc spec error: {exc}") from exc
    if not spec.body:
        raise MdEngineError("md имеет только титул — тело пустое, fail-closed")
    return spec


def md_build_v2(md_text: str, output_docx: str, author: str,
                profile: ProfileName = "corporate",
                base_dir: str | None = None) -> dict:
    """md -> DOCX (конверт builder'а: ok/status/data/metadata)."""
    if not md_text.strip():
        raise MdEngineError("empty md input — fail-closed")
    profile = profile or "corporate"
    spec = markdown_to_spec(md_text, output_docx, author, profile, base_dir=base_dir)
    return builder.build_document(spec, output_docx)


def md_build_v2_file(input_md: str, output_docx: str, author: str,
                     profile: ProfileName = "corporate") -> dict:
    path = Path(input_md)
    if not path.is_file():
        raise MdEngineError(f"input not found: {input_md}")
    return md_build_v2(path.read_text(encoding="utf-8-sig", errors="strict"),
                       output_docx, author, profile, base_dir=str(path.parent))
