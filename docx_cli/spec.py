"""DocSpec pydantic model and profile defaults."""
from __future__ import annotations

from datetime import datetime, timezone
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, field_validator

from docx_cli.errors import ProfileError, SpecError


ProfileName = Literal["gost", "compact", "court", "vs_sud", "soyu_rayon"]
Align = Literal["left", "right", "center", "justify"]
BlockType = Literal[
    "paragraph",
    "numbered_list",
    "table",
    "signature_block",
    "attachments_list",
    "pleading_marker",
]


class OutputSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    file_name: str
    author: str


class HeaderSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    align: Literal["left", "right", "center"] = "right"
    lines: list[str] = Field(..., min_length=1)
    left_indent_cm: float = 7.0

    @field_validator("left_indent_cm")
    @classmethod
    def left_indent_in_range(cls, v: float) -> float:
        if v < 6.0 or v > 8.0:
            raise ValueError("header.left_indent_cm must be between 6.0 and 8.0")
        return v


class TitleSpec(BaseModel):
    model_config = ConfigDict(extra="forbid")

    main: str
    subtitle: str | None = None


class ParagraphBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["paragraph"]
    text: str
    inline_markdown: bool = True


class NumberedListBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["numbered_list"]
    items: list[str] = Field(..., min_length=1)


class TableBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["table"]
    headers: list[str]
    rows: list[list[str]] = Field(..., min_length=1)
    widths_cm: list[float] | None = None

    @field_validator("rows")
    @classmethod
    def rows_valid(cls, rows: list[list[str]]) -> list[list[str]]:
        if len({len(r) for r in rows}) != 1:
            raise ValueError("all table rows must have the same number of cells")
        width = len(rows[0])
        if width == 0:
            raise ValueError("table rows must contain at least one cell")
        for r_idx, row in enumerate(rows):
            for c_idx, cell in enumerate(row):
                if cell == "":
                    raise ValueError(
                        f"empty cell at row {r_idx}, column {c_idx}"
                    )
        return rows

    @field_validator("widths_cm")
    @classmethod
    def widths_match_headers(
        cls, widths: list[float] | None, info: Any
    ) -> list[float] | None:
        if widths is None:
            return None
        headers = info.data.get("headers")
        header_count = len(headers) if headers else 0
        if len(widths) != header_count:
            raise ValueError(
                f"widths_cm length ({len(widths)}) must match headers count ({header_count})"
            )
        return widths


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


class AttachmentsListBlock(BaseModel):
    model_config = ConfigDict(extra="forbid")

    type: Literal["attachments_list"]
    items: list[str] = Field(..., min_length=1)


BodyBlock = (
    ParagraphBlock
    | NumberedListBlock
    | TableBlock
    | SignatureBlock
    | AttachmentsListBlock
    | PleadingBlock
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
    body_alignment: int = 3  # WD_ALIGN_PARAGRAPH.JUSTIFY
    cell_font_size_pt: int = 11
    default_signature_format: str = "table"

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
        elif name != "gost":
            raise ProfileError(f"unknown profile: {name}")

        if line_spacing is not None:
            self.body_line_spacing = line_spacing
            self.header_line_spacing = line_spacing


def now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()
