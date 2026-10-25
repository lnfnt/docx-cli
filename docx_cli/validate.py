"""Built-in post-build validation: reopen the generated .docx and verify it."""
from __future__ import annotations

import re
import zipfile
from typing import Any

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import qn

from docx_cli.spec import DocSpec, Profile, now_iso


RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES = "[Content_Types].xml"
RELS = "_rels/.rels"
DOC_RELS = "word/_rels/document.xml.rels"
DOCUMENT_XML = "word/document.xml"


_REQUISITE_RE = re.compile(r"(паспорт|снилс|инн|кпп|огрн)", re.IGNORECASE)


def _cm(value: Any) -> float:
    return float(value.cm) if value is not None else 0.0


def _pt(value: Any) -> float | None:
    return float(value.pt) if value is not None else None


def _check_package_integrity(path: str) -> list[str]:
    errors: list[str] = []
    required = {CONTENT_TYPES, RELS, DOC_RELS, DOCUMENT_XML}
    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = set(zf.namelist())
            missing = required - names
            if missing:
                errors.append(f"missing_required_parts: {sorted(missing)}")
                return errors
            for rel_file in (RELS, DOC_RELS):
                data = zf.read(rel_file)
                tree = parse_xml(data)
                for rel in tree.findall(f"{{{RELS_NS}}}Relationship"):
                    target = rel.get("Target")
                    if target is None:
                        continue
                    if target.startswith("/"):
                        target = target[1:]
                    else:
                        base = "word/" if rel_file == DOC_RELS else ""
                        target = base + target
                    target = target.replace("\\", "/")
                    # Normalize ../ paths
                    parts = target.split("/")
                    stack: list[str] = []
                    for part in parts:
                        if part == "..":
                            if stack:
                                stack.pop()
                        elif part and part != ".":
                            stack.append(part)
                    normalized = "/".join(stack)
                    if normalized and normalized not in names:
                        errors.append(f"unresolved_relationship: {rel_file} -> {normalized}")
    except zipfile.BadZipFile as exc:
        errors.append(f"bad_zip: {exc}")
    except Exception as exc:
        errors.append(f"package_integrity: {exc}")
    return errors


def _check_margins(doc: Document, profile: Profile) -> list[str]:
    errors: list[str] = []
    section = doc.sections[0]
    pairs = [
        ("left", _cm(section.left_margin), profile.margin_left_cm),
        ("right", _cm(section.right_margin), profile.margin_right_cm),
        ("top", _cm(section.top_margin), profile.margin_top_cm),
        ("bottom", _cm(section.bottom_margin), profile.margin_bottom_cm),
    ]
    for name, actual, expected in pairs:
        if abs(actual - expected) > 0.05:
            errors.append(f"margin_{name}: expected {expected}cm, got {actual:.3f}cm")
    return errors


def _iter_all_paragraphs(doc: Document):
    """Yield every paragraph in the document, including table cells."""
    idx = 0
    for p in doc.paragraphs:
        yield False, idx, p
        idx += 1
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield True, idx, p
                    idx += 1


def _expected_paragraph_texts(spec: DocSpec, profile: Profile) -> dict[str, set[str]]:
    """Build exact-text sets for paragraphs rendered outside normal body flow."""
    header_texts = {line.strip() for line in spec.header.lines if line != ""}
    title_texts = {spec.title.main.strip()}
    if spec.title.subtitle:
        title_texts.add(spec.title.subtitle.strip())

    pleading_texts: set[str] = set()
    signature_texts: set[str] = set()
    list_texts: set[str] = set()

    for block in spec.body:
        if block.type == "pleading_marker":
            pleading_texts.add(block.text.strip())
        elif block.type == "signature_block":
            fmt = block.format if block.format is not None else profile.default_signature_format
            if fmt == "text":
                signature_texts.add(f"{block.left}\t{block.right}".strip())
        elif block.type == "numbered_list":
            for idx, item in enumerate(block.items, start=1):
                list_texts.add(f"{idx}. {item}".strip())
        elif block.type == "attachments_list":
            for item in block.items:
                list_texts.add(item.strip())

    return {
        "header": header_texts,
        "title": title_texts,
        "pleading": pleading_texts,
        "signature": signature_texts,
        "list": list_texts,
    }


_SIG_HEAD_RE = re.compile(
    r"^(Защитник|Адвокат|Представитель|Ордер|Истец|Ответчик|Заявитель|"
    r"Жалобщик|Подозреваемый|Обвиняемый)\b"
)


def _is_service_signature(p, text: str) -> bool:
    """v0.5.1 (по внешней рецензии): подписная/служебная строка, набранная
    пробелами без tab+'___'. Признаки: ролевое слово в начале строки И
    отсутствие красной строки — обычный body-абзац («Защитник указал, что…»)
    идёт с отступом 1.25 см и так проверяется дальше."""
    if not _SIG_HEAD_RE.match(text):
        return False
    # Настоящая подпись с tab+___ проверяется отдельной веткой; это —
    # «пробельная» подпись без таба, для неё красная строка не требуется.
    if "\t" in text and "___" in text:
        return False
    fl = p.paragraph_format.first_line_indent
    return fl is None or int(fl) == 0


def _classify_paragraphs(doc: Document, spec: DocSpec | None, profile: Profile | None = None,
                      skipped: list | None = None) -> dict[int, str]:
    """Map paragraph index to role: header, title, body, list, pleading, signature, other."""
    roles: dict[int, str] = {}
    if spec is not None:
        profile = profile or Profile(spec.profile)
        expected = _expected_paragraph_texts(spec, profile)
        attachments_seen_spec = False  # v0.5.1: зеркало fallback-логики
        for idx, p in enumerate(doc.paragraphs):
            text = p.text.strip()
            if not text:
                attachments_seen_spec = False
                continue
            if text in expected["header"]:
                roles[idx] = "header"
            elif text in expected["title"]:
                roles[idx] = "title"
            elif text in expected["pleading"]:
                roles[idx] = "pleading"
            elif text in expected["signature"]:
                roles[idx] = "signature"
            elif text in expected["list"]:
                roles[idx] = "list"
            elif re.match(r"^Приложения\s*[:.]?\s*$", text):
                attachments_seen_spec = True
                roles[idx] = "other"
            elif attachments_seen_spec or _is_service_signature(p, text):
                # v0.5.1 (по внешней рецензии): симметрично
                # fallback-ветке — подписные/служебные строки и продолжения
                # списка приложений, не совпавшие с exact-текстом spec,
                # не проверяются как body (красная строка).
                if skipped is not None:
                    if _is_service_signature(p, text):
                        skipped.append({"type": "check_skipped", "zone": idx,
                                        "reason": "service_signature"})
                    elif attachments_seen_spec:
                        skipped.append({"type": "check_skipped", "zone": idx,
                                        "reason": "appendices_fallback"})
                roles[idx] = "signature"
            else:
                roles[idx] = "body"
        return roles

    # Fallback heuristic when no spec is supplied (standalone validate command).
    title_idx: int | None = None
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() and p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            # First centered non-empty paragraph is treated as title
            title_idx = idx
            break
    # Title block = title paragraph + any immediately following centered
    # paragraphs (subtitle chain), e.g. "ОТЗЫВ ИСТЦА" + "на возражения...".
    title_end = title_idx
    if title_idx is not None:
        for idx in range(title_idx + 1, len(doc.paragraphs)):
            p = doc.paragraphs[idx]
            if (
                p.text.strip()
                and p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
            ):
                title_end = idx
            else:
                break
    attachments_seen = False  # v0.5.1: всё после «Приложения:» — служебный блок
    for idx, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            attachments_seen = False
            continue
        if title_idx is not None and title_idx <= idx <= title_end:
            roles[idx] = "title"
        elif title_idx is not None and idx < title_idx:
            roles[idx] = "header"
        elif (
            p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
            and (
                p.paragraph_format.first_line_indent is None
                or int(p.paragraph_format.first_line_indent) == 0
            )
            and text.upper().startswith("ПРОШУ")
        ):
            roles[idx] = "pleading"
        elif (
            p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.CENTER
            and (
                p.paragraph_format.first_line_indent is None
                or int(p.paragraph_format.first_line_indent) == 0
            )
        ):
            # v0.5.3 (фикс 1/2 инвариант): центрированные структурные блоки
            # (подзаголовок «## …», маркеры) — не body: иначе md-build порождает
            # документ, который бракует собственный validate по alignment.
            roles[idx] = "other"
        elif (
            p.paragraph_format.alignment == WD_ALIGN_PARAGRAPH.LEFT
            and "\t" in p.text
            and "___" in p.text
        ):
            roles[idx] = "signature"
        elif attachments_seen:
            # v0.5.1 (по внешней рецензии): продолжения
            # списка приложений («Ордер адвоката …») и подпись после «Приложения:»
            # — служебный блок, красная строка к ним не применяется.
            if skipped is not None:
                skipped.append({"type": "check_skipped", "zone": idx,
                                "reason": "appendices_fallback"})
            roles[idx] = "other"
        elif _is_service_signature(p, text):
            if skipped is not None:
                skipped.append({"type": "check_skipped", "zone": idx,
                                "reason": "service_signature"})
            roles[idx] = "signature"
        elif text == "Приложения:" or re.match(r"^Приложения\s*[:.]?\s*$", text) \
                or re.match(r"^\d+\.\s", text):
            # Numbered-list / attachment-list paragraphs are not GOST body;
            # skip body checks rather than false-fail. In standalone mode we
            # do not have the spec, so any "N. " paragraph is treated as a
            # list item.
            if text.startswith("Приложения"):
                attachments_seen = True
            roles[idx] = "other"
        elif p.style.name.startswith("List") or (
            p.paragraph_format.left_indent is not None
            and int(p.paragraph_format.left_indent) > 0
        ):
            # v0.1.0 B-10 (контроль координатора): style-mode вложенные
            # списки (List Bullet/Number N) и legacy-рендер уровня 1
            # (Normal + левый отступ 35.5pt) — структурные list-параграфы,
            # не body: ГОСТ-красная строка/JUSTIFY к ним не применяется
            # (ровно как уже исключаются "N. " и hanging-indent списки).
            if skipped is not None:
                skipped.append({"type": "check_skipped", "zone": idx,
                                "reason": "list_paragraph"})
            roles[idx] = "other"
        elif (
            abs(_cm(p.paragraph_format.left_indent) - 1.25) <= 0.05
            and abs(_cm(p.paragraph_format.first_line_indent) + 0.75) <= 0.05
        ):
            # v0.2.0 fix: hanging-indent paragraphs (lists/attachments rendered
            # by builder with left=1.25cm, first=-0.75cm) are structural list
            # items even when text heuristics miss them (e.g. text-form
            # signature lines, non-numeric list items). Точный маркер
            # builder'а надёжнее текстовых эвристик.
            roles[idx] = "other"
        else:
            roles[idx] = "body"
    return roles


def _all_runs(p):
    from docx.oxml.ns import qn
    from docx.text.run import Run
    for r in p._p.iter(qn("w:r")):
        yield Run(r, p)


def _check_fonts_and_sizes(doc: Document, spec: DocSpec | None, guessed_profile: str = "gost") -> list[str]:
    errors: list[str] = []
    profile = Profile(spec.profile if spec else guessed_profile, line_spacing=spec.line_spacing if spec else None)
    roles = _classify_paragraphs(doc, spec, profile)

    for is_cell, p_idx, p in _iter_all_paragraphs(doc):
        role = roles.get(p_idx, "cell" if is_cell else "other")
        for run in _all_runs(p):
            text = run.text
            if not text:
                continue
            name = run.font.name
            if name != profile.font_name:
                errors.append(f"font_name: paragraph {p_idx} run uses '{name}'")
            size = _pt(run.font.size)
            if role == "title":
                is_title_main = run.bold and size == profile.title_font_size_pt
                if not is_title_main and size != profile.body_font_size_pt:
                    errors.append(f"title_font_size: paragraph {p_idx} run size {size}pt")
            elif role == "cell":
                if size is not None and size != profile.cell_font_size_pt:
                    errors.append(f"cell_font_size: paragraph {p_idx} run size {size}pt")
            else:
                if size is not None and size != profile.body_font_size_pt:
                    errors.append(f"body_font_size: paragraph {p_idx} run size {size}pt")
    return errors


def _check_body_formatting(doc: Document, spec: DocSpec | None, guessed_profile: str = "gost") -> list[str]:
    errors: list[str] = []
    profile = Profile(spec.profile if spec else guessed_profile, line_spacing=spec.line_spacing if spec else None)
    roles = _classify_paragraphs(doc, spec, profile)

    for p_idx, p in enumerate(doc.paragraphs):
        role = roles.get(p_idx, "other")
        text = p.text.strip()
        is_svc = _is_service_signature(p, text)
        if role == "body":
            pass
        elif role == "signature" and is_svc:
            pass
        else:
            continue
        pf = p.paragraph_format
        align = pf.alignment
        if align != WD_ALIGN_PARAGRAPH.JUSTIFY:
            errors.append(f"alignment: paragraph {p_idx} expected JUSTIFY, got {align}")
        ls = pf.line_spacing
        if isinstance(ls, (int, float)) and abs(float(ls) - profile.body_line_spacing) > 0.05:
            errors.append(
                f"line_spacing: paragraph {p_idx} expected {profile.body_line_spacing}, got {ls}"
            )
        if role == "body":
            indent = _cm(pf.first_line_indent)
            if abs(indent - profile.first_line_indent_cm) > 0.05:
                errors.append(
                    f"first_line_indent: paragraph {p_idx} expected {profile.first_line_indent_cm}cm, got {indent:.3f}cm"
                )
    return errors


def _check_list_formatting(doc: Document, spec: DocSpec | None, profile: Profile) -> list[str]:
    """Hanging indent for numbered_list and attachments_list items."""
    errors: list[str] = []
    if spec is None:
        return errors

    roles = _classify_paragraphs(doc, spec, profile)
    for p_idx, p in enumerate(doc.paragraphs):
        if roles.get(p_idx) != "list":
            continue
        pf = p.paragraph_format
        left = _cm(pf.left_indent)
        first = _cm(pf.first_line_indent)
        if abs(left - 1.25) > 0.05:
            errors.append(f"list_left_indent: paragraph {p_idx} expected 1.25cm, got {left:.3f}cm")
        if abs(first - (-0.75)) > 0.05:
            errors.append(f"list_first_line_indent: paragraph {p_idx} expected -0.75cm, got {first:.3f}cm")
    return errors


def _check_header_formatting(doc: Document, spec: DocSpec | None, profile: Profile) -> list[str]:
    errors: list[str] = []
    if spec is None or profile.name != "court":
        return errors

    roles = _classify_paragraphs(doc, spec, profile)
    expected_left = spec.header.left_indent_cm
    for p_idx, p in enumerate(doc.paragraphs):
        if roles.get(p_idx) != "header":
            continue
        pf = p.paragraph_format
        if pf.alignment != WD_ALIGN_PARAGRAPH.LEFT:
            errors.append(f"header_alignment: paragraph {p_idx} expected LEFT, got {pf.alignment}")
        left = _cm(pf.left_indent)
        if abs(left - expected_left) > 0.05:
            errors.append(f"header_left_indent: paragraph {p_idx} expected {expected_left}cm, got {left:.3f}cm")
        ls = pf.line_spacing
        if isinstance(ls, (int, float)) and abs(float(ls) - profile.header_line_spacing) > 0.05:
            errors.append(f"header_line_spacing: paragraph {p_idx} expected {profile.header_line_spacing}, got {ls}")
    return errors


def _check_pleading_marker(doc: Document, spec: DocSpec | None, profile: Profile) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    if spec is None:
        return errors, warnings

    # Warn when a plain body paragraph starts with ПРОШУ in court profile.
    if profile.name == "court":
        for block in spec.body:
            if block.type == "paragraph" and block.text.strip().upper().startswith("ПРОШУ"):
                warnings.append("pleading_marker: use dedicated 'pleading_marker' block for text starting with 'ПРОШУ'")
                break

    roles = _classify_paragraphs(doc, spec, profile)
    for p_idx, p in enumerate(doc.paragraphs):
        if roles.get(p_idx) != "pleading":
            continue
        pf = p.paragraph_format
        if pf.alignment != WD_ALIGN_PARAGRAPH.CENTER:
            errors.append(f"pleading_alignment: paragraph {p_idx} expected CENTER, got {pf.alignment}")
        first = _cm(pf.first_line_indent)
        if abs(first - 0.0) > 0.05:
            errors.append(f"pleading_first_line_indent: paragraph {p_idx} expected 0cm, got {first:.3f}cm")
        for run in p.runs:
            if not run.text:
                continue
            if not run.bold:
                errors.append(f"pleading_bold: paragraph {p_idx} run is not bold")
            size = _pt(run.font.size)
            if size is not None and size != profile.body_font_size_pt:
                errors.append(f"pleading_font_size: paragraph {p_idx} run size {size}pt")
    return errors, warnings


def _check_en_dash(doc: Document, spec: DocSpec | None) -> tuple[list[str], list[str]]:
    errors: list[str] = []
    warnings: list[str] = []
    en_dash = "\u2013"
    no_en_dash = spec.checks.no_en_dash if spec else True
    for is_cell, p_idx, p in _iter_all_paragraphs(doc):
        if en_dash in p.text:
            msg = f"en_dash: paragraph {p_idx} contains '–'"
            if no_en_dash:
                errors.append(msg)
            else:
                warnings.append(msg)
    return errors, warnings


def _is_requisite_line(text: str) -> bool:
    return bool(re.search(r"\d{6,}", text) or _REQUISITE_RE.search(text))


def _check_yo(doc: Document, spec: DocSpec | None) -> list[str]:
    warnings: list[str] = []
    if spec is not None and not spec.checks.no_yo:
        return warnings
    whitelist: set[str] = set()
    if spec is not None:
        whitelist = {line.strip() for line in spec.header.lines if _is_requisite_line(line)}
    for is_cell, p_idx, p in _iter_all_paragraphs(doc):
        text = p.text
        if any(text.strip() == wl for wl in whitelist):
            continue
        if "ё" in text or "Ё" in text:
            warnings.append(f"yo_present: paragraph {p_idx} contains 'ё' outside requisites")
    return warnings


def _check_placeholders(doc: Document, spec: DocSpec | None) -> list[str]:
    warnings: list[str] = []
    pattern = re.compile(r"\[УКАЗАТЬ[^\]]*\]")
    for is_cell, p_idx, p in _iter_all_paragraphs(doc):
        if pattern.search(p.text):
            warnings.append(f"placeholder_present: paragraph {p_idx}")
    return warnings


def _check_norm_refs(doc: Document, spec: DocSpec | None) -> list[str]:
    warnings: list[str] = []
    if spec is None or not spec.checks.require_norms_refs:
        return warnings
    full_text = "\n".join(p.text for p in doc.paragraphs)
    if not re.search(
        r"(?:ст\.?\s*\d|Стать[ья]|ГК\s*РФ|КоАП|НК\s*РФ|ТК\s*РФ|ГПК\s*РФ|КАС\s*РФ)",
        full_text,
    ):
        warnings.append("norm_refs_missing: no legal norm references detected")
    return warnings


def _check_numbered_list_sequence(doc: Document, spec: DocSpec | None) -> list[str]:
    errors: list[str] = []
    pattern = re.compile(r"^(\d+)\.\s")
    # v0.5.1 (round-trip приёмка): блок приложений («Приложения:») — отдельный
    # нумерованный список, его нумерация легитимно начинается с 1 после списка
    # просительной части; в v0.5.0 законные [1 | 1] давали «out of order».
    # v0.5.3 (фикс 3): границы сегментов — «ПРОШУ:» и
    # «Приложения:» без учёта регистра и двоеточия/точки (рабочие документы
    # пишут «прошу:», «приложения.», «ПРОШУ СУД:» и т.п.).
    boundary = re.compile(r"^(?:приложения|прошу|дополнительно)(?:\b|\s*[:.])", re.IGNORECASE)
    segments: list[list[int]] = [[]]
    for p in doc.paragraphs:
        t = p.text.strip()
        if boundary.match(t):
            if segments[-1]:
                segments.append([])
            continue
        m = pattern.match(t)
        if m:
            segments[-1].append(int(m.group(1)))
    for numbers in segments:
        if not numbers:
            continue
        groups: list[list[int]] = []
        current: list[int] = []
        for n in numbers:
            # A new list starts at 1 ONLY when the previous group is complete
            # ([1..k]); otherwise a repeated/out-of-order number stays in the
            # current group so the sequence check flags it (fail-closed).
            # v0.5.3 фикс 3: полная группа может иметь ЛЮБУЮ длину, включая
            # [1] — независимые одноэлементные перечни законны (склейка
            # [1|1] в v0.5.2 браковала реальные документы).
            if n == 1 and current:
                groups.append(current)
                current = [n]
            else:
                current.append(n)
        if current:
            groups.append(current)

        for grp in groups:
            expected = list(range(1, len(grp) + 1))
            if grp != expected:
                errors.append(f"numbered_list_sequence: {grp} out of order")
    return errors


def _guess_profile(doc) -> str:
    """Standalone validate has no spec: guess profile from dominant font size.
    14pt body => court profile; otherwise gost (12pt)."""
    from collections import Counter
    sizes = Counter()
    for p in doc.paragraphs:
        for run in p.runs:
            if run.text.strip() and run.font.size is not None:
                sizes[round(run.font.size.pt)] += 1
    return "court" if sizes and sizes.most_common(1)[0][0] == 14 else "gost"


# v0.5.3 (фикс 5; согласовано с внешним рецензентом): полосы геометрии для входящих
# процессуальных актов судов и оппонентов. В --mode review они дают ТОЛЬКО
# warnings; errors в review — строго контентные дефекты (плейсхолдеры,
# en dash, norm_refs, склейка нумерации и т.п.).
REVIEW_MARGIN_BAND_CM = (1.0, 3.0)
REVIEW_FONT_BAND_PT = (10.0, 14.0)
REVIEW_LINE_SPACING_BAND = (1.0, 1.5)
REVIEW_FIRST_LINE_BAND_CM = (0.5, 1.5)


def _in_band(v: float, lo: float, hi: float) -> bool:
    return lo <= v <= hi


def _check_geometry_review(doc: Document, spec: DocSpec | None,
                           profile: Profile) -> list[str]:
    """Review-режим: геометрия как warnings (полосы реальной судебной практики)."""
    warnings: list[str] = []
    lo, hi = REVIEW_MARGIN_BAND_CM
    section = doc.sections[0]
    for name, actual in (("left", _cm(section.left_margin)),
                         ("right", _cm(section.right_margin)),
                         ("top", _cm(section.top_margin)),
                         ("bottom", _cm(section.bottom_margin))):
        if not _in_band(actual, lo, hi):
            warnings.append(
                f"review_margin_{name}: {actual:.3f}cm вне полосы {lo}-{hi}cm")
    f_lo, f_hi = REVIEW_FONT_BAND_PT
    for is_cell, p_idx, p in _iter_all_paragraphs(doc):
        for run in p.runs:
            if not run.text:
                continue
            if run.font.name and run.font.name != profile.font_name:
                warnings.append(
                    f"review_font_name: paragraph {p_idx} run uses '{run.font.name}'")
            size = _pt(run.font.size)
            if size is not None and not _in_band(size, f_lo, f_hi):
                warnings.append(
                    f"review_font_size: paragraph {p_idx} run size {size:.1f}pt "
                    f"вне полосы {f_lo:.0f}-{f_hi:.0f}pt")
    s_lo, s_hi = REVIEW_LINE_SPACING_BAND
    i_lo, i_hi = REVIEW_FIRST_LINE_BAND_CM
    roles = _classify_paragraphs(doc, spec)
    for p_idx, p in enumerate(doc.paragraphs):
        if roles.get(p_idx) != "body":
            continue
        pf = p.paragraph_format
        ls = pf.line_spacing
        if isinstance(ls, (int, float)) and not _in_band(float(ls), s_lo, s_hi):
            warnings.append(
                f"review_line_spacing: paragraph {p_idx} = {float(ls):.2f} "
                f"вне полосы {s_lo}-{s_hi}")
        ind = _cm(pf.first_line_indent)
        if _in_band(ind, i_lo, i_hi):
            continue
        if ind == 0.0:
            # Отсутствие красной строки легитимно при явном межпараграфном
            # интервале (полоса, согласованная с внешним рецензентом).
            after = pf.space_after.pt if pf.space_after is not None else 0.0
            before = pf.space_before.pt if pf.space_before is not None else 0.0
            if max(after, before) <= 0:
                warnings.append(
                    f"review_first_line_indent: paragraph {p_idx} — отступ "
                    f"отсутствует и межпараграфный интервал не задан")
        else:
            warnings.append(
                f"review_first_line_indent: paragraph {p_idx} = {ind:.3f}cm "
                f"вне полосы {i_lo}-{i_hi}cm")
    return warnings


def validate_document(path: str, spec: DocSpec | None = None,
                      mode: str = "gost") -> dict[str, Any]:
    """Reopen generated .docx and run all built-in checks.

    mode='gost' (default) — строгая геометрия по профилю (собственные
    процессуальные документы), прежнее поведение без изменений.
    mode='review' — входящие акты судов/оппонентов: геометрия даёт только
    warnings (полосы реальной практики), errors остаются строго контентными
    (фикс 5, согласовано с внешним рецензентом).
    """
    if mode not in ("gost", "review"):
        raise ValueError(f"unknown validate mode: {mode!r} (expected: gost|review)")
    # Standalone mode: guess the profile from the document itself so court
    # documents are checked against court rules, not gost.
    effective_profile_name = spec.profile if spec else None
    if effective_profile_name is None:
        try:
            from docx import Document as _Document
            effective_profile_name = _guess_profile(_Document(path))
        except Exception:
            effective_profile_name = "gost"
    profile = Profile(
        effective_profile_name,
        line_spacing=spec.line_spacing if spec else None,
    )
    errors: list[str] = []
    warnings: list[str] = []
    checks: dict[str, str] = {}

    try:
        doc = Document(path)
    except Exception as exc:
        return {
            "checks": {"open": "fail"},
            "warnings": [],
            "errors": [f"open_failed: {exc}"],
            "skipped": [],
            "skips": [],
            "findings": [],
        }

    skipped: list = []
    try:
        _classify_paragraphs(doc, spec, profile, skipped=skipped)
    except Exception:
        pass
    # de-duplicate identical (zone, reason) records collected across passes
    _seen = set()
    _uniq = []
    for _f in skipped:
        _key = (_f.get("zone"), _f.get("reason"))
        if _key in _seen:
            continue
        _seen.add(_key)
        _uniq.append(_f)
    skipped = _uniq

    pi_errors = _check_package_integrity(path)
    errors.extend(pi_errors)
    checks["package_integrity"] = "fail" if pi_errors else "pass"

    if mode == "review":
        # v0.5.3 фикс 5: геометрия — только warnings; строгие профильные
        # проверки (margins/fonts/body/list/header/pleading geometry) не
        # блокируют review входящего документа.
        geom_warnings = _check_geometry_review(doc, spec, profile)
        warnings.extend(geom_warnings)
        checks["margins"] = "warn" if any(
            w.startswith("review_margin") for w in geom_warnings) else "pass"
        checks["fonts"] = "warn" if any(
            w.startswith("review_font") for w in geom_warnings) else "pass"
        checks["body_formatting"] = "warn" if any(
            w.startswith(("review_line_spacing", "review_first_line"))
            for w in geom_warnings) else "pass"
        checks["list_formatting"] = "n/a"
        checks["header_formatting"] = "n/a"
        checks["pleading_marker"] = "n/a"
    else:
        margin_errors = _check_margins(doc, profile)
        errors.extend(margin_errors)
        checks["margins"] = "fail" if margin_errors else "pass"

        guessed = effective_profile_name
        font_errors = _check_fonts_and_sizes(doc, spec, guessed_profile=guessed)
        errors.extend(font_errors)
        checks["fonts"] = "fail" if font_errors else "pass"

        fmt_errors = _check_body_formatting(doc, spec, guessed_profile=guessed)
        errors.extend(fmt_errors)
        checks["body_formatting"] = "fail" if fmt_errors else "pass"

        list_errors = _check_list_formatting(doc, spec, profile)
        errors.extend(list_errors)
        checks["list_formatting"] = "fail" if list_errors else "pass"

        header_errors = _check_header_formatting(doc, spec, profile)
        errors.extend(header_errors)
        checks["header_formatting"] = "fail" if header_errors else "pass"

        pleading_errors, pleading_warnings = _check_pleading_marker(doc, spec, profile)
        errors.extend(pleading_errors)
        warnings.extend(pleading_warnings)
        checks["pleading_marker"] = "fail" if pleading_errors else ("warn" if pleading_warnings else "pass")

    en_errors, en_warnings = _check_en_dash(doc, spec)
    errors.extend(en_errors)
    warnings.extend(en_warnings)
    checks["no_en_dash"] = "fail" if en_errors else ("warn" if en_warnings else "pass")

    yo_warnings = _check_yo(doc, spec)
    warnings.extend(yo_warnings)
    checks["no_yo"] = "warn" if yo_warnings else "pass"

    ph_findings = _check_placeholders(doc, spec)
    if mode == "review":
        # Фикс 5: остаточные плейсхолдеры — контентный error и в review.
        errors.extend(ph_findings)
        checks["placeholders"] = "fail" if ph_findings else "pass"
    else:
        warnings.extend(ph_findings)
        checks["placeholders"] = "warn" if ph_findings else "pass"

    norm_findings = _check_norm_refs(doc, spec)
    if mode == "review":
        # Фикс 5: отсутствие ссылок на нормы (когда проверка активна) —
        # контентный error и в review.
        errors.extend(norm_findings)
        checks["norm_refs"] = "fail" if norm_findings else "pass"
    else:
        warnings.extend(norm_findings)
        checks["norm_refs"] = "warn" if norm_findings else "pass"

    seq_errors = _check_numbered_list_sequence(doc, spec)
    errors.extend(seq_errors)
    checks["numbered_list_sequence"] = "fail" if seq_errors else "pass"

    return {
        "validated_at": now_iso(),
        "mode": mode,
        "profile": effective_profile_name,
        "checks": checks,
        "warnings": warnings,
        "errors": errors,
        "skipped": skipped,
        "skips": skipped,
        "findings": skipped,
    }
