# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""DocSpec pydantic model and profile defaults."""
from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator, model_validator

from docx_cli.errors import ProfileError, SpecError


ProfileName = Literal["gost", "compact", "court", "vs_sud", "soyu_rayon", "corporate", "report"]
Align = Literal["left", "right", "center", "justify"]
BlockType = Literal[
    "paragraph",
    "numbered_list",
    "bullet_list",
    "heading",
    "table",
    "signature_block",
    "attachments_list",
    "pleading_marker",
    "callout",
    "page_break",
    "horizontal_rule",
    "toc",
    "section_break",
    "image",
]

_HEX_DIGITS = set("0123456789abcdefABCDEF")


def _validate_hex_color(value: str | None) -> str | None:
    """HEX-цвет: принимает '#RRGGBB' или 'RRGGBB', возвращает 'RRGGBB'."""
    if value is None:
        return None
    v = value.lstrip("#")
    if len(v) != 6 or any(ch not in _HEX_DIGITS for ch in v):
        raise ValueError(f"hex color must be '#RRGGBB', got {value!r}")
    return v


class OutputSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    file_name: str
    author: str


class HeaderSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    align: Literal["left", "right", "center"] = "right"
    lines: list[str] = Field(..., min_length=1)
    left_indent_cm: float = 7.0
    # Офис-эталон 07.11.2026: строки шапки с space_after 2 pt, интервал 1.5,
    # пустая строка в lines дает space_after 3 pt на предыдущем абзаце.
    line_space_after_pt: float | None = None
    line_spacing: float | None = None

    @field_validator("left_indent_cm")
    @classmethod
    def left_indent_in_range(cls, v: float) -> float:
        # 5.25 см — офис-эталон (07.11.2026, «Дополнения к АЖ2»); нижняя граница 5.0
        if v < 5.0 or v > 8.0:
            raise ValueError("header.left_indent_cm must be between 5.0 and 8.0")
        return v


class TitleSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    main: str
    subtitle: str | None = None
    # Офис-эталон 07.11.2026: заголовок 12 pt bold (не title_font_size профиля),
    # main — space_after 6 pt, subtitle — space_after 18 pt.
    size_pt: float | None = None
    main_space_after_pt: float | None = None
    subtitle_space_after_pt: float | None = None
    subtitle_bold: bool = False  # эталон 07.11.2026: subtitle жирный


class ParagraphBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["paragraph"]
    text: str
    inline_markdown: bool = True
    # Переопределение space_after для подзаголовков/разделителей (эталон 07.11.2026)
    space_after_pt: float | None = None


class NumberedListBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["numbered_list"]
    items: list[str] = Field(..., min_length=1)
    start_at: int = Field(default=1, ge=1)
    # ПАКЕТ 3: ilvl (0..2) для каждого item; пусто/все 0 -> плоский GOST-список
    # с явными "N. " префиксами, уровень >0 -> настоящая нумерация OpenXML (numPr).
    levels: list[int] = Field(default_factory=list)

    @model_validator(mode="after")
    def levels_match_items(self) -> "NumberedListBlock":
        if self.levels:
            if len(self.levels) != len(self.items):
                raise ValueError("levels length must match items length")
            if any(l not in (0, 1, 2) for l in self.levels):
                raise ValueError("levels must be in range 0..2")
            # S1-M7: список без уровня 0 даёт битый lvlText «%1.%2.» (%1 не
            # определён ни одним элементом) — Word показывает мусор в маркере.
            if self.levels[0] != 0:
                raise ValueError("levels must start at 0 (no level-0 item -> broken lvlText)")
        return self


class BulletListBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["bullet_list"]
    items: list[str] = Field(..., min_length=1)
    levels: list[int] = Field(default_factory=list)

    @model_validator(mode="after")
    def levels_match_items(self) -> "BulletListBlock":
        if self.levels:
            if len(self.levels) != len(self.items):
                raise ValueError("levels length must match items length")
            if any(l not in (0, 1, 2) for l in self.levels):
                raise ValueError("levels must be in range 0..2")
        return self


class CalloutBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["callout"]
    text: str | list[str]
    accent: str | None = None  # цвет левой границы; None -> #2B579A
    shading: str | None = "#F9F9F9"  # фон выносного блока
    italic: bool = False
    inline_markdown: bool = True

    @field_validator("accent", "shading")
    @classmethod
    def hex_colors(cls, v: str | None) -> str | None:
        return _validate_hex_color(v)


class PageBreakBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["page_break"]


class HeadingBlock(BaseModel):
    """Фаза 2/3: заголовок Markdown `#`..`####` -> стиль Заголовок 1-4."""
    model_config = ConfigDict(extra="forbid")

    type: Literal["heading"]
    level: int = Field(ge=1, le=4)
    text: str
    color: str | None = None  # HEX; None -> цвет из профиля (corporate/report)

    @field_validator("color")
    @classmethod
    def hex_colors(cls, v: str | None) -> str | None:
        return _validate_hex_color(v)


class HorizontalRuleBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["horizontal_rule"]


class TocBlock(BaseModel):
    """ПАКЕТ 1: автособираемое оглавление — нативное поле Word TOC."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["toc"]
    title: str = "Содержание"
    levels: int = Field(default=3, ge=1, le=4)


class SectionBreakBlock(BaseModel):
    """ПАКЕТ 1 & 9: раздел с реальной сменой геометрии страницы (w:sectPr) и колонок."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["section_break"]
    orientation: Literal["portrait", "landscape"] = "portrait"
    break_type: Literal["next_page", "continuous"] = "next_page"
    margin_top_cm: float | None = None
    margin_bottom_cm: float | None = None
    margin_left_cm: float | None = None
    margin_right_cm: float | None = None
    # S1-H4: валидация на границе spec→render — без неё отрицательные/
    # нулевые/тысячные колонки уходили в невалидный OOXML.
    cols: int = Field(default=1, ge=1, le=10)
    col_space_cm: float = Field(default=1.25, gt=0, le=5)
    col_sep: bool = False


class ImageBlock(BaseModel):
    """ПАКЕТ 2: изображение DrawingML (add_picture) с подписью и масштабированием."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["image"] = "image"
    path: str  # локальный путь к файлу изображения (PNG, JPEG, GIF)
    caption: str | None = None  # подпись под рисунком ("Рисунок 1: Динамика цен")
    width_cm: float | None = Field(default=None, gt=0)  # явная ширина в см
    width_pct: float | None = Field(default=None, gt=0, le=100)  # ширина в % рабочей области (S1-M4: cap 100)
    align: Literal["center", "left", "right"] = "center"

    @model_validator(mode="after")
    def one_width(self) -> "ImageBlock":
        if self.width_cm is not None and self.width_pct is not None:
            raise ValueError("width_cm and width_pct are mutually exclusive")
        return self


class CheckboxBlock(BaseModel):
    """ПАКЕТ 6: интерактивный чекбокс — нативный w:sdt + w14:checkbox."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["checkbox"] = "checkbox"
    label: str = ""
    checked: bool = False


class DropdownItem(BaseModel):
    """ПАКЕТ 6: элемент выпадающего списка с value/displayText."""

    model_config = ConfigDict(extra="forbid", populate_by_name=True)

    value: str
    display_text: str | None = Field(default=None, alias="displayText")

    @property
    def display(self) -> str:
        return self.display_text or self.value


class DropdownBlock(BaseModel):
    """ПАКЕТ 6: выпадающий список — нативный w:sdt + w:dropDownList."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["dropdown"] = "dropdown"
    items: list[str | DropdownItem] = Field(..., min_length=1)
    selected: str | None = None
    label: str | None = None

    @model_validator(mode="after")
    def selected_in_items(self) -> "DropdownBlock":
        values = [it.value if isinstance(it, DropdownItem) else it for it in self.items]
        if self.selected is not None and self.selected not in values:
            raise ValueError("selected must be one of items")
        return self


class DatePickerBlock(BaseModel):
    """ПАКЕТ 6: селектор даты — нативный w:sdt + w:date."""

    model_config = ConfigDict(extra="forbid")

    type: Literal["date_picker"] = "date_picker"
    initial_date: str = ""  # YYYY-MM-DD или DD.MM.YYYY; "" — пустой пикер
    date_format: str = "dd.MM.yyyy"

    @field_validator("initial_date")
    @classmethod
    def date_shape(cls, v: str) -> str:
        # S1-M8/S2-M5: shape-only regex пропускал «2026-13-45» и падал на
        # рендере — валидируем реальным календарём на границе spec.
        if v:
            from datetime import datetime as _dt
            ok = False
            for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
                try:
                    _dt.strptime(v, fmt)
                    ok = True
                    break
                except ValueError:
                    continue
            if not ok:
                raise ValueError("initial_date must be a real date (YYYY-MM-DD or DD.MM.YYYY)")
        return v


class TableCell(BaseModel):
    """Словарная ячейка таблицы v2: {"text": ..., "colspan": 2, ...}."""
    model_config = ConfigDict(extra="forbid")

    text: str = ""
    colspan: int = Field(default=1, ge=1)
    rowspan: int = Field(default=1, ge=1)
    bold: bool | None = None  # None -> унаследовать шапку (bold) / тело (не-bold)
    italic: bool | None = None
    align: Align | None = None  # None -> шапка center, тело left
    shading: str | None = None  # HEX; None -> header_shading в шапке

    @field_validator("shading")
    @classmethod
    def hex_shading(cls, v: str | None) -> str | None:
        return _validate_hex_color(v)


class TableBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["table"]
    headers: list[str | TableCell]
    rows: list[list[str | TableCell]] = Field(..., min_length=1)
    widths_cm: list[float] | None = None  # легаси-алиас col_widths_cm
    col_widths_cm: list[float] | None = None
    col_widths_pct: list[float] | None = None
    border_style: Literal["grid", "header_only", "none", "boxed"] = "grid"
    header_shading: str | None = "#F2F2F2"
    vertical_align: Literal["top", "center", "bottom"] = "center"
    tbl_header: bool = True  # w:tblHeader — шапка повторяется на каждой странице
    cant_split: bool = True  # w:cantSplit — запрет разрыва строк между страницами

    @field_validator("header_shading")
    @classmethod
    def hex_header_shading(cls, v: str | None) -> str | None:
        return _validate_hex_color(v)

    @model_validator(mode="after")
    def grid_valid(self) -> "TableBlock":
        grid_width = len(self.headers)
        # Заголовки — одна ячейка на колонку (colspan/rowspan — только в body).
        for i, h in enumerate(self.headers):
            if isinstance(h, TableCell) and (h.colspan != 1 or h.rowspan != 1):
                raise ValueError(f"header cell {i} must not use colspan/rowspan")

        # Ровно один способ задать ширины; длина = числу колонок.
        width_sources = [
            (self.col_widths_cm, "col_widths_cm"),
            (self.col_widths_pct, "col_widths_pct"),
            (self.widths_cm, "widths_cm"),
        ]
        used = [name for lst, name in width_sources if lst is not None]
        if len(used) > 1:
            raise ValueError(
                f"widths_cm, col_widths_cm and col_widths_pct are mutually exclusive, got {used}"
            )
        widths = next(lst for lst, _ in width_sources if lst is not None) if used else None
        if widths is not None:
            if len(widths) != grid_width:
                raise ValueError(
                    f"{used[0]} length ({len(widths)}) must match headers count ({grid_width})"
                )
            if any(w <= 0 for w in widths):
                raise ValueError("column widths must be positive")
            if self.col_widths_pct is not None and sum(widths) > 100.0 + 1e-9:
                raise ValueError("col_widths_pct must sum to <= 100")

        # Симуляция сетки: colspan/rowspan должны укладываться в grid_width
        # без пересечений с вертикальными продолжениями.
        _grid_placement(self.rows, grid_width)
        return self

    def grid_layout(self) -> list[list[tuple[int, int, "TableCell | str"]]]:
        """Раскладка data-строк по сетке: [(col_start, colspan, ячейка)] на строку."""
        return _grid_placement(self.rows, len(self.headers))


def _grid_placement(
    rows: list[list["TableCell | str"]], grid_width: int
) -> list[list[tuple[int, int, "TableCell | str"]]]:
    """Валидация + раскладка логических ячеек по сетке таблицы.

    Ячейки строки занимают свободные колонки слева направо; колонки, покрытые
    rowspan предыдущих строк, считаются занятыми (вертикальные продолжения
    в spec явно не перечисляются). Бросает ValueError при выходе за сетку.
    """
    placement: list[list[tuple[int, int, "TableCell | str"]]] = []
    # Таблица без span-ячеек вообще — легаси-контракт строгих длин строк.
    spanned = any(
        isinstance(raw, TableCell) and (raw.colspan > 1 or raw.rowspan > 1)
        for row in rows for raw in row
    )
    pending: dict[int, int] = {}  # колонка -> сколько строк ниже ещё занято
    for r_idx, row in enumerate(rows):
        if not row:
            raise ValueError("table rows must contain at least one cell")
        placement_row: list[tuple[int, int, "TableCell | str"]] = []
        occupied = {c for c, remaining in pending.items() if remaining > 0}
        col = 0
        for c_idx, raw in enumerate(row):
            if isinstance(raw, TableCell):
                span, rspan, text = raw.colspan, raw.rowspan, raw.text
            else:
                span, rspan, text = 1, 1, raw
                if text == "":
                    raise ValueError(f"empty cell at row {r_idx}, column {c_idx}")
            while col in occupied:
                col += 1
            start = col
            if start + span > grid_width:
                raise ValueError(
                    f"cell at row {r_idx}, column {c_idx} (colspan {span}) "
                    f"does not fit into {grid_width} columns"
                )
            for k in range(start, start + span):
                if k in occupied:
                    raise ValueError(
                        f"cell at row {r_idx}, column {c_idx} overlaps a "
                        f"rowspan continuation at column {k}"
                    )
            if rspan > 1:
                pending[start] = rspan
            placement_row.append((start, span, raw))
            col = start + span
        # Легаси-контракт: строка короче сетки только когда её колонки заняты
        # rowspan-продолжениями (v2). Обычная таблица с кривыми длинами строк —
        # та же ошибка, что и раньше.
        if not spanned and sum(s for _, s, _ in placement_row) + len(occupied) != grid_width:
            raise ValueError("all table rows must have the same number of cells")
        placement.append(placement_row)
        for c in pending:
            if pending[c] > 0:
                pending[c] -= 1
    stuck = [c for c, remaining in pending.items() if remaining > 0]
    if stuck:
        raise ValueError(f"rowspan extends below the last table row at column {stuck[0]}")
    return placement


class SignatureBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["signature_block"]
    left: str
    right: str
    format: Literal["table", "text"] | None = None


class PleadingBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["pleading_marker"]
    text: str


class CommentSpec(BaseModel):
    """ПАКЕТ 4: комментарий на полях к фрагменту target в собранном документе.

    Целевой фрагмент ищется в параграфах документа и оборачивается в
    w:commentRangeStart/End с w:commentReference; тело — в word/comments.xml.

    S2-H6: при повторных вхождениях target якорь определяется явным
    occurrence (1-based, в порядке документа). Нет вхождений или выход
    за диапазон — fail-closed BuildError, не молчаливое первое совпадение.
    """

    model_config = ConfigDict(extra="forbid")

    target: str  # текст фрагмента в документе, к которому относится замечание
    text: str  # текст замечания
    author: str = "colleague"
    date: str | None = None  # ISO-дата; None -> текущий момент UTC

    @field_validator("date")
    @classmethod
    def date_iso(cls, v: str | None) -> str | None:
        # S1-M8: неотвалидированная строка уходила в w:date и ломала пакет.
        if v is None:
            return v
        from datetime import datetime as _dt
        try:
            _dt.fromisoformat(v.replace("Z", "+00:00"))
        except ValueError:
            raise ValueError("date must be ISO-8601 (e.g. 2026-11-02 or 2026-11-02T12:00:00Z)")
        return v
    occurrence: int = Field(default=1, ge=1)  # какое по счёту вхождение (S2-H6)


class AttachmentsListBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["attachments_list"]
    items: list[str] = Field(..., min_length=1)


class MathBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["math"] = "math"
    latex: str
    caption: str | None = None


BodyBlock = (
    ParagraphBlock
    | NumberedListBlock
    | BulletListBlock
    | HeadingBlock
    | TableBlock
    | SignatureBlock
    | AttachmentsListBlock
    | PleadingBlock
    | CalloutBlock
    | PageBreakBlock
    | HorizontalRuleBlock
    | TocBlock
    | SectionBreakBlock
    | ImageBlock
    | CheckboxBlock
    | DropdownBlock
    | DatePickerBlock
    | MathBlock
)


class ChecksSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    no_yo: bool = True
    no_en_dash: bool = True
    require_norms_refs: bool = False


class DocSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    profile: ProfileName = "gost"
    line_spacing: float | None = None
    output: OutputSpec
    header: HeaderSpec
    title: TitleSpec
    body: list[BodyBlock]
    # Фаза 1 Headless Word В: колонтитулы и нумерация страниц.
    footer_text: str | None = None
    page_numbers: bool = True  # «Стр. {PAGE} из {NUMPAGES}» в нижнем колонтитуле
    different_first_page: bool = True  # титульная страница без колонтитулов
    # ПАКЕТ 1: markdown-определения сносок `[^id]: текст` (id -> текст);
    # ссылки `[^id]` остаются в тексте параграфов и превращаются в нативные
    # w:footnoteReference при рендеринге.
    footnotes: dict[str, str] = Field(default_factory=dict)
    # ПАКЕТ 8: markdown-определения концевых сносок `[^!id]: текст` (id -> текст);
    # ссылки `[^!id]` остаются в тексте параграфов и превращаются в нативные
    # w:endnoteReference при рендеринге.
    endnotes: dict[str, str] = Field(default_factory=dict)
    # ПАКЕТ 4: комментарии рецензента — фрагмент target оборачивается в
    # w:commentRangeStart/End (см. docx_cli.comments) после сборки тела.
    comments: list[CommentSpec] = Field(default_factory=list)
    # ПАКЕТ 9: защита документа (read_only, comments_only, tracked_changes, forms)
    # S1-M2/S1-M8: Literal на границе — произвольная строка уходила в settings.xml
    protection: Literal["read_only", "comments_only", "tracked_changes", "forms"] | None = None
    # S2-H3: пароль защиты; без него enforcement запрещён (empty-password equivalent)
    protection_password: str | None = None
    checks: ChecksSpec = Field(default_factory=ChecksSpec)

    @field_validator("line_spacing")
    @classmethod
    def line_spacing_allowed(cls, v: float | None) -> float | None:
        if v is not None and v not in (1.0, 1.15):
            raise ValueError("line_spacing must be 1.0 or 1.15")
        return v

    @field_validator("body")
    @classmethod
    def body_not_empty(cls, body: list[BodyBlock]) -> list[BodyBlock]:
        if not body:
            raise ValueError("body must contain at least one block")
        return body

    @classmethod
    def from_dict(cls, raw: dict[str, Any]) -> "DocSpec":
        try:
            return cls.model_validate(raw)
        except Exception as exc:
            raise SpecError(f"spec_error: {exc}") from exc


class Profile:
    """Default GOST legal profile."""

    name: ProfileName = "gost"
    paper_width_cm: float = 21.0
    paper_height_cm: float = 29.7
    margin_left_cm: float = 3.0
    margin_right_cm: float = 1.5
    margin_top_cm: float = 2.0
    margin_bottom_cm: float = 2.0
    font_name: str = "Times New Roman"
    body_font_size_pt: int = 12
    title_font_size_pt: int = 14
    body_line_spacing: float = 1.5
    header_line_spacing: float = 1.15
    first_line_indent_cm: float = 1.25
    footnotes: dict[str, str] = {}
    endnotes: dict[str, str] = {}
    body_alignment: int = 3  # WD_ALIGN_PARAGRAPH.JUSTIFY
    cell_font_size_pt: int = 11
    default_signature_format: str = "table"
    # Фаза 2/3: корпоративные профили.
    heading_color: str | None = None  # HEX без '#'; None -> цвет не задаётся
    default_table_border: str = "grid"  # border_style таблиц из md-движка

    def __init__(self, name: ProfileName = "gost", line_spacing: float | None = None) -> None:
        if name == "compact":
            self.name = "compact"
            self.margin_left_cm = 2.0
            self.margin_right_cm = 1.0
            self.margin_top_cm = 1.5
            self.margin_bottom_cm = 1.5
            self.body_line_spacing = 1.15
        elif name == "court":
            self.name = "court"
            self.body_font_size_pt = 14
            self.title_font_size_pt = 14
            self.cell_font_size_pt = 14
            self.body_line_spacing = 1.0
            self.header_line_spacing = 1.0
            self.default_signature_format = "text"
        elif name in ("vs_sud", "soyu_rayon"):
            # v0.1.0 T-05: форматтерские профили ВС / СОЮ первой инстанции.
            # Геометрия идентична gost (профиль = имя + пути); никакого
            # юридического текста в профиле — нормативную часть вносит
            # legal-профиль через spec (плейсхолдеры в шаблонах).
            self.name = name
        elif name == "corporate":
            # Фаза 2/3 ТЗ §Б: чистый корпоративный стиль. Arial, поля 2.0 см,
            # межстрочный 1.15, цвет заголовков #1F4E79.
            self.name = "corporate"
            self.font_name = "Arial"
            self.margin_left_cm = 2.0
            self.margin_right_cm = 2.0
            self.body_line_spacing = 1.15
            self.header_line_spacing = 1.15
            self.body_font_size_pt = 11
            self.title_font_size_pt = 18
            self.cell_font_size_pt = 10
            self.first_line_indent_cm = 0.0
            self.heading_color = "1F4E79"
        elif name == "report":
            # Фаза 2/3 ТЗ §Б: аналитический отчет — Calibri, акцентные
            # выноски и таблицы с линией только под шапкой.
            self.name = "report"
            self.font_name = "Calibri"
            self.margin_left_cm = 2.0
            self.margin_right_cm = 2.0
            self.body_line_spacing = 1.15
            self.header_line_spacing = 1.15
            self.body_font_size_pt = 11
            self.title_font_size_pt = 18
            self.cell_font_size_pt = 10
            self.first_line_indent_cm = 0.0
            self.heading_color = "1F4E79"
            self.default_table_border = "header_only"
        elif name != "gost":
            raise ProfileError(f"unknown profile: {name}")

        if line_spacing is not None:
            self.body_line_spacing = line_spacing
            self.header_line_spacing = line_spacing


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()
