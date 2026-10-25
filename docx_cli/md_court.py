"""docx-cli md-build — сборка процессуального .docx из Markdown.

v0.4.0 (2026-10-20): собственная clean-room реализация сборки
процессуальных документов из Markdown.
v0.1.0 B-10: вложенные списки глубины 1-4 (List Bullet/Number стили с
fallback ручных отступов), GFM pipe-таблицы (выравнивания, inline-конверт
ячеек); неконвертируемые конструкции — честные skip-маркеры
skipped_construct в report (дизайн: docs/design/nested-lists-tables.md).
Правила верстки судебного документа:
  - шапка (реквизиты, до титула) — слева с отступом ~8.1 см, «Поле: значение»;
  - титульная строка (ИСКОВОЕ ЗАЯВЛЕНИЕ / ОТЗЫВ / ХОДАТАЙСТВО / …) —
    по центру, bold 14pt;
  - тело — по ширине, отступ первой строки 1.25 см, интервал 1.5;
  - маркеры «- »/«* » и нумерация «1.»/«1.1.» — висячий отступ;
  - **жирный** inline; строка с ПРОЕКТ/ЧЕРНОВИК — центр, italic;
  - фильтр пустых «---» разделителей.

Fail-closed: пустой md / нет титула -> ошибка (не молчаливая генерация).
"""
from __future__ import annotations

import re
from pathlib import Path

import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Inches, Pt

# v0.5.1 (по внешней рецензии): исправлена опечатка
# «АПЕЛЯЦИОННАЯ» (одна «Л») — корректное написание «АПЕЛЛЯЦИОННАЯ»; старый
# вариант оставлен алиасом, чтобы документы с опечаткой в титуле не падали.
# Добавлены вариации дополнений/дополнительных жалоб и номеров (№ N).
TITLE_WORDS = {
    "ОТЗЫВ ИСТЦА", "ОТЗЫВ", "ВОЗРАЖЕНИЯ", "ВОЗРАЖЕНИЯ НА ОТЗЫВ",
    "ИСКОВОЕ ЗАЯВЛЕНИЕ", "ВСТРЕЧНОЕ ИСКОВОЕ ЗАЯВЛЕНИЕ", "ХОДАТАЙСТВО",
    "ОТВЕТ НА ОТЗЫВ", "АПЕЛЛЯЦИОННАЯ ЖАЛОБА", "КАССАЦИОННАЯ ЖАЛОБА",
    "ДОПОЛНЕНИЕ К ИСКОВОМУ ЗАЯВЛЕНИЮ", "ЗАЯВЛЕНИЕ",
    # алиас опечатки (легаси-документы):
    "АПЕЛЯЦИОННАЯ ЖАЛОБА",
    # жалоба как базовый титул и надзорные формы:
    "ЖАЛОБА", "ЧАСТНАЯ ЖАЛОБА", "НАДЗОРНАЯ ЖАЛОБА",
    # дополнения:
    "ДОПОЛНЕНИЕ", "ДОПОЛНЕНИЯ",
    "ДОПОЛНЕНИЕ К ЖАЛОБЕ", "ДОПОЛНЕНИЯ К ЖАЛОБЕ",
    "ДОПОЛНЕНИЕ К АПЕЛЛЯЦИОННОЙ ЖАЛОБЕ", "ДОПОЛНЕНИЯ К АПЕЛЛЯЦИОННОЙ ЖАЛОБЕ",
    "ДОПОЛНЕНИЕ К АПЕЛЯЦИОННОЙ ЖАЛОБЕ", "ДОПОЛНЕНИЯ К АПЕЛЯЦИОННОЙ ЖАЛОБЕ",
    "ДОПОЛНЕНИЕ К КАССАЦИОННОЙ ЖАЛОБЕ", "ДОПОЛНЕНИЯ К КАССАЦИОННОЙ ЖАЛОБЕ",
    "ДОПОЛНИТЕЛЬНАЯ ЖАЛОБА", "ДОПОЛНИТЕЛЬНАЯ АПЕЛЛЯЦИОННАЯ ЖАЛОБА",
    "ДОПОЛНИТЕЛЬНАЯ АПЕЛЯЦИОННАЯ ЖАЛОБА",
}

# «ДОПОЛНЕНИЯ (№ 2)» / «ЖАЛОБА № 3» / «... (№2)» — номерной маркер не часть титула.
_TITLE_NUM_RE = re.compile(r"\s*\(?\s*(?:№|N)\s*\d+\s*\)?\s*", re.UNICODE)



def _hanging_first_indent_cm(digits: int) -> float:
    """Таблица ширин висячего отступа по разряду числа (1, 2, 3+ цифры)."""
    if digits <= 1:
        return -0.75
    if digits == 2:
        return -1.0
    return -1.25


def _normalize_dashes(text: str) -> str:
    """Единая нормализация тире: en dash (–) -> em dash (—)."""
    return text.replace("\u2013", "\u2014")


# алиасы для совместимости с возможными импортами
normalize_dashes = _normalize_dashes
normalize_dash = _normalize_dashes
dash_normalize = _normalize_dashes

def _normalize_title(s: str) -> str:
    s = s.replace("\ufeff", "").strip()
    s = re.sub(r"^#+\s*", "", s)  # v0.5.1: md-заголовок «# ДОПОЛНЕНИЯ (№ 2)»
    s = s.upper().rstrip(".:")
    s = _TITLE_NUM_RE.sub(" ", s)
    return " ".join(s.split()).rstrip(".:")

BULLET_PREFIXES = ("- ", "* ")

# v0.5.3 (фикс 1): анатомия дефектов из внешнего аудита.
# «## N. Подзаголовок» — маркер срезается, bold, центр, без fli.
_HEADING_RE = re.compile(r"^#{2,6}\s+(\S.*)$")
# Перечни «1)», «2)» — та же висячая геометрия, что у «N. » (left 1.25, first -0.75).
_PAREN_LIST_RE = re.compile(r"^\d+\)")
# «ПРОШУ:» / «ПРОШУ» / «ПРОШУ СУД:» — просительный маркер: центр, без отступа.
_PLEADING_RE = re.compile(r"^ПРОШУ(?:\s+СУД)?\s*[:.]?\s*$", re.UNICODE)
# Предлоги продолжения титула (вторая строка: «к…» / «на…» / «по…» / «о…»).
_TITLE_PREP_RE = re.compile(r"^(?:об|к|на|по|от|для|о)\b", re.IGNORECASE)

# --- v0.1.0 B-10: вложенные списки и pipe-таблицы -------------------------
# Элемент списка с опциональным отступом; маркер «- »/«* » или «N. »/«N) ».
_LIST_ITEM_RE = re.compile(r"^([ \t]*)([-*]|\d+[.)])\s+(.+)$")
# GFM-разделитель таблицы: ячейки из «-» с опциональными «:» выравнивания.
_TABLE_SEP_RE = re.compile(
    r"^[ \t]*\|?[ \t]*:?-{1,}:?[ \t]*(?:\|[ \t]*:?-{1,}:?[ \t]*)*\|?[ \t]*$")
# Блочная конструкция внутри ячейки (цитата/маркер списка) — строка скипается
# с честным skip-маркером (дизайн: docs/design/nested-lists-tables.md §2.3).
_BLOCK_IN_CELL_RE = re.compile(r"^(?:>|[-*]\s|\d+[.)]\s)")
MAX_LIST_LEVEL = 4  # глубина >4 не рендерится — skip-маркер list_depth_gt4


def _expanded_indent(raw: str) -> int:
    return len(raw) - len(raw.expandtabs(4).lstrip(" \t"))


def _is_table_start(lines: list[str], i: int) -> bool:
    """Строка i — заголовок GFM-таблицы, если строка i+1 — разделитель."""
    if i + 1 >= len(lines):
        return False
    cur, nxt = lines[i], lines[i + 1]
    return "|" in cur and bool(_TABLE_SEP_RE.match(nxt)) and bool(nxt.strip())


def _split_table_row(raw: str) -> list[str]:
    """Ячейки строки: внешние pipes необязательны, разбиваем по НЕэкранированным
    '|'; '\\|' -> литеральная '|' в тексте ячейки."""
    t = raw.strip()
    if t.startswith("|"):
        t = t[1:]
    if t.endswith("|") and not t.endswith("\\|"):
        t = t[:-1]
    parts = re.split(r"(?<!\\)\|", t)
    return [p.strip().replace("\\|", "|") for p in parts]


def _table_aligns(sep_raw: str) -> list:
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    out = []
    for c in _split_table_row(sep_raw):
        left = c.startswith(":")
        right = c.endswith(":")
        if left and right:
            out.append(A.CENTER)
        elif right:
            out.append(A.RIGHT)
        else:
            out.append(A.LEFT)
    return out


_CELL_INLINE_TOKEN_RE = re.compile(r"(<u>|</u>|\*\*|_|`)")


def _add_cell_runs(p, text: str, bold_all: bool = False) -> None:
    """v0.1.0 B-10: inline-конверт ячейки таблицы: **bold**, _italic_,
    `code` (Courier New), <u>underline</u>. Отдельный от _add_inline_runs:
    обычные параграфы тела сохраняют прежний конверт без backtick-ветки
    (регресс-пин v0.9)."""
    parts = _CELL_INLINE_TOKEN_RE.split(text)
    bold = under = ital = code = False

    def _emit(chunk: str) -> None:
        if not chunk:
            return
        run = p.add_run(_normalize_dashes(chunk))
        run.font.name = "Times New Roman"  # v0.1.0 контроль: валидатор запрещает иные шрифты (font_name)
        # контроль координатора v0.1.0: cells = gost cell_font_size (11pt).
        # md-build рендерит тело 12pt => _guess_profile даёт gost, а
        # валидатор требует cell_font_size_pt=11; было Pt(12) — таблицы в
        # md-build падали на собственном validate (cell_font_size).
        run.font.size = Pt(11)
        if bold or bold_all:
            run.bold = True
        if under:
            run.underline = True
        if ital:
            run.italic = True

    for i, tok in enumerate(parts):
        if tok == "**":
            bold = not bold
            continue
        if tok == "`":
            code = not code
            continue
        if tok == "<u>":
            under = True
            continue
        if tok == "</u>":
            under = False
            continue
        if tok == "_":
            prev_word = i > 0 and bool(re.search(r"\w$", parts[i - 1] or ""))
            next_word = bool(re.match(r"^\w", parts[i + 1] if i + 1 < len(parts) else ""))
            if prev_word and next_word:
                _emit("_")  # интравордовый подчёркиватель — литерал
                continue
            ital = not ital
            continue
        _emit(tok)


def _render_md_table(doc, lines: list[str], i: int, report: list) -> int:
    """GFM-таблица, начиная со строки i (заголовок). Возвращает индекс строки
    после таблицы. Блочная конструкция в ячейке -> skip ВСЕЙ строки с
    skip-маркером; заголовок с блоком -> skip всей таблицы (честно, §2.3)."""
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls

    headers = _split_table_row(lines[i])
    aligns = _table_aligns(lines[i + 1])
    ncols = len(headers)
    aligns = (aligns + [A.LEFT] * ncols)[:ncols]

    body_start = i + 2
    j = body_start
    raw_rows: list[tuple[int, str]] = []  # (line-index, raw)
    while j < len(lines):
        nxt = lines[j]
        if not nxt.strip() or "|" not in nxt:
            break
        raw_rows.append((j, nxt))
        j += 1

    def dash_check(idx: int) -> None:
        ln = lines[idx]
        if "–" in ln:
            report.append({"type": "dash_normalized", "zone": len(doc.paragraphs),
                           "line": ln.strip(), "from": "–", "to": "—"})

    for k in range(i, j):
        dash_check(k)

    if any(_BLOCK_IN_CELL_RE.match(h) for h in headers):
        report.append({"type": "skipped_construct", "construct": "block_in_cell",
                       "zone": len(doc.paragraphs), "line": i + 1,
                       "text": lines[i].strip(),
                       "reason": "блочный элемент в строке заголовка — таблица "
                                 "пропущена целиком"})
        return j

    rows: list[list[str]] = []
    for (ln_idx, raw_line) in raw_rows:
        cells = _split_table_row(raw_line)
        if any(_BLOCK_IN_CELL_RE.match(c) for c in cells):
            report.append({"type": "skipped_construct",
                           "construct": "block_in_cell",
                           "zone": len(doc.paragraphs), "line": ln_idx + 1,
                           "text": raw_line.strip(),
                           "reason": "блочный элемент в ячейке не поддерживается "
                                     "— строка пропущена целиком"})
            continue
        if len(cells) != ncols:
            report.append({"type": "skipped_construct",
                           "construct": "row_width_mismatch",
                           "zone": len(doc.paragraphs), "line": ln_idx + 1,
                           "text": raw_line.strip(),
                           "reason": f"ожидалось {ncols} ячеек, получено "
                                     f"{len(cells)} — pad/truncate"})
            cells = (cells + [""] * ncols)[:ncols]
        rows.append(cells)

    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    try:
        table.style = "Table Grid"
    except KeyError:
        pass  # стиль может отсутствовать в нестандартном шаблоне — не падаем
    trPr = table.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    for c, head in enumerate(headers):
        p = table.cell(0, c).paragraphs[0]
        p.alignment = aligns[c]
        _add_cell_runs(p, head, bold_all=True)
    for r, cells in enumerate(rows, start=1):
        for c in range(ncols):
            p = table.cell(r, c).paragraphs[0]
            p.alignment = aligns[c]
            _add_cell_runs(p, cells[c])
    return j


def _list_style_names(doc) -> set[str]:
    try:
        return {s.name for s in doc.styles}
    except Exception:  # noqa: BLE001 — набор стилей читаем всегда из шаблона
        return set()


def _render_list_run(doc, lines: list[str], run: list[tuple[int, str]],
                     report: list) -> None:
    """Run подряд идущих элементов списка. level<=1 — легаси-рендер (_add_body,
    байт-стабилен); levels 2-4 — стили List Bullet/Number N, при отсутствии
    стиля manual fallback (left = 1.25+0.75·(level−1), first −0.75, для
    numbered — литеральный computed номер). level>4 — skip-маркер."""
    style_names = _list_style_names(doc)
    parsed = []
    for (idx, raw) in run:
        m = _LIST_ITEM_RE.match(raw.expandtabs(4))
        if m is None:  # защита: run собирается только из matching-строк
            continue
        indent = len(raw.expandtabs(4)) - len(raw.expandtabs(4).lstrip(" \t"))
        parsed.append({"idx": idx, "raw": raw, "indent": indent,
                       "ordered": bool(re.match(r"^\d+[.)]$", m.group(2))),
                       "marker": m.group(2), "text": m.group(3)})
    pos = [p["indent"] for p in parsed if p["indent"] > 0]
    unit = min(pos) if pos else 2
    for p in parsed:
        p["level"] = p["indent"] // unit + 1
    max_level = max(p["level"] for p in parsed)
    counters: dict[int, int] = {}
    for p in parsed:
        ln = p["raw"]
        if "–" in ln:
            report.append({"type": "dash_normalized", "zone": len(doc.paragraphs),
                           "line": ln.strip(), "from": "–", "to": "—"})
        if max_level <= 1:
            _add_body(doc, ln.strip())
            continue
        level = p["level"]
        if level > MAX_LIST_LEVEL:
            report.append({"type": "skipped_construct",
                           "construct": "list_depth_gt4",
                           "zone": len(doc.paragraphs), "line": p["idx"] + 1,
                           "text": ln.strip(),
                           "reason": f"глубина {level} > {MAX_LIST_LEVEL} "
                                     "не поддерживается"})
            continue
        if level == 1:
            _add_body(doc, ln.strip())
            continue
        ordered = p["ordered"]
        number = None
        if ordered:
            number = counters.get(level, 1)
            counters[level] = number + 1
        style = f"List {'Number' if ordered else 'Bullet'} {level}"
        para = None
        if style in style_names:
            try:
                para = doc.add_paragraph(style=style)
            except KeyError:
                para = None
        if para is None:
            para = doc.add_paragraph()
            pf = para.paragraph_format
            pf.left_indent = Cm(1.25 + 0.75 * (level - 1))
            pf.first_line_indent = Cm(-0.75)
            pf.space_after = Pt(6)
            pf.line_spacing = 1.5
            if ordered:
                _add_inline_runs(para, f"{number}. {p['text']}")
                continue
        _add_inline_runs(para, p["text"])


def _is_title_continuation(line: str) -> bool:
    """Вторая строка титула («к апелляционной жалобе…», «на приговор…»):
    строчное продолжение заголовка или строка с титульного предлога.
    Реквизиты («Истец: …») и обычный текст не подходят."""
    if _HEADING_RE.match(line):
        return False
    if line[:1].islower():
        return True
    return bool(_TITLE_PREP_RE.match(line))


def _is_pleading_marker(line: str) -> bool:
    return bool(_PLEADING_RE.match(line.upper().strip()))


class MdBuildError(Exception):
    """Fail-closed ошибка сборки md -> docx."""


def _is_title(line: str) -> bool:
    s = _normalize_title(line)
    if s in TITLE_WORDS:
        return True
    if s.startswith("НА ") and "ОТВЕТЧИК" in s or s.startswith("НА ВОЗРАЖЕН"):
        return True
    return False

def _pick_title_index(lines: list[str]) -> int | None:
    """Выбрать индекс «настоящего» титула с учётом приоритета: кандидаты,
    попавшие в блок реквизитов (строки до первой пустой строки, когда
    документ начинается с реквизитов), не титулуются."""
    stripped = [ln.strip() for ln in lines]
    first_nonempty = next((i for i, ln in enumerate(stripped) if ln), None)
    if first_nonempty is None:
        return None
    first_is_title = _is_title(_normalize_dashes(stripped[first_nonempty]))
    if first_is_title:
        ineligible: set[int] = set()
    else:
        first_blank = next((i for i, ln in enumerate(stripped) if not ln), len(stripped))
        ineligible = set(range(first_blank))
    candidates = [i for i, ln in enumerate(stripped)
                  if ln and _is_title(_normalize_dashes(ln))
                  and not _is_title_continuation(ln)
                  and i not in ineligible]
    if len(candidates) == 1:
        return candidates[0]
    all_candidates = [i for i, ln in enumerate(stripped)
                      if ln and _is_title(_normalize_dashes(ln))
                      and not _is_title_continuation(ln)]
    return all_candidates[0] if all_candidates else None


_INLINE_TOKEN_RE = re.compile(r"(<u>|</u>|\*\*|_)")


def _add_inline_runs(p, text: str) -> None:
    """v0.5.2: inline-разметка md -> runs без утечки тегов в текст.
    <u>…</u> -> подчёркивание, _…_ -> курсив (интравордовый '_' — литерал,
    по правилу CommonMark для '_'), **…** -> жирный; флаги комбинируются.
    Утечка буквальных '<u>'/'_' в текст документа — дефект, зафиксированный
    внешним аудитом (п.3)."""
    parts = _INLINE_TOKEN_RE.split(text)
    bold = under = ital = False

    def _emit(chunk: str) -> None:
        if not chunk:
            return
        # v0.5.1: типографика ГОСТ — en dash «–» в исходном md заменяем на em
        # dash «—» (валидатор no_en_dash собственный output иначе бракует).
        run = p.add_run(_normalize_dashes(chunk))
        if bold:
            run.bold = True
        if under:
            run.underline = True
        if ital:
            run.italic = True

    for i, tok in enumerate(parts):
        if tok == "**":
            bold = not bold
            continue
        if tok == "<u>":
            under = True
            continue
        if tok == "</u>":
            under = False
            continue
        if tok == "_":
            prev_word = i > 0 and bool(re.search(r"\w$", parts[i - 1] or ""))
            next_word = bool(re.match(r"^\w", parts[i + 1] if i + 1 < len(parts) else ""))
            if prev_word and next_word:
                _emit("_")  # интравордовый подчёркиватель — литерал, не маркер
                continue
            ital = not ital
            continue
        _emit(tok)


def _add_body(doc, text: str) -> None:
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    mnum = re.match(r"^(\d+)[.)]\s", text)
    if any(text.startswith(b) for b in BULLET_PREFIXES):
        text = text[2:]
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.left_indent = Cm(1.25)
    elif mnum or _PAREN_LIST_RE.match(text) \
            or text[:4] in ("1.1.", "1.2.", "1.3.", "2.1.", "2.2.", "3.1.", "3.2.") \
            or re.match(r"^\d+[.)]\s", text):
        digits = len(mnum.group(1)) if mnum else (len(re.match(r"^(\d+)", text).group(1)) if re.match(r"^(\d+)", text) else 1)
        first = _hanging_first_indent_cm(digits)
        p.paragraph_format.first_line_indent = Cm(first)
        p.paragraph_format.left_indent = Cm(1.25)
    else:
        p.paragraph_format.first_line_indent = Cm(1.25)
    _add_inline_runs(p, text)


def build_md_document(md_text: str, author: str, report: list | None = None) -> docx.Document:
    """md-текст -> Document. Ошибка, если нет титульной строки (fail-closed)."""
    doc = docx.Document()
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(3)
        section.right_margin = Cm(1.5)  # v0.5.1: согласовано с gost-профилем — иначе собственный validate бракует сборку
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)

    if report is None:
        report = []
    in_header = True
    title_seen = False
    title_block_open = False  # v0.5.3 фикс 1: «вторая строка титула» ещё не принята
    lines = md_text.split("\n")
    _title_target = _pick_title_index(lines)
    n = len(lines)
    i = 0
    while i < n:
        raw = lines[i]
        cur = i
        line = raw.strip()
        i += 1  # ветки ниже могут дополнительно сдвинуть i (блоки таблиц/списков)
        if not line or line in ("---", "***"):
            continue
        if "\u2013" in line:
            report.append({"type": "dash_normalized", "zone": len(doc.paragraphs),
                           "line": line, "from": "\u2013", "to": "\u2014"})
        if "ПРОЕКТ" in line.upper() or "ЧЕРНОВИК" in line.upper():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing = 1.5
            r = p.add_run(_normalize_dashes(line))
            r.italic = True
            in_header = True
            continue
        if not title_seen and _title_target is not None and cur == _title_target:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(_normalize_dashes(re.sub(r"^#+\s*", "", line)))  # без md-маркера '#'
            r.bold = True
            r.font.size = Pt(14)
            in_header = False
            title_seen = True
            title_block_open = True
            continue
        if title_block_open and _is_title_continuation(line):
            # v0.5.3 фикс 1: вторая строка титула остаётся в титульном блоке —
            # центр, без отступа (раньше уходила в body justify + fli 1.25).
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            _add_inline_runs(p, line)
            title_block_open = False
            continue
        title_block_open = False
        hm = _HEADING_RE.match(line)
        if hm:
            content = _normalize_dashes(hm.group(1))
            # «## N. Текст» и «## Текст» -> заголовок (title-регистр по профилю),
            # без литеральных «##».
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(6)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.15
            r = p.add_run(content)
            r.bold = True
            continue
        if _is_pleading_marker(line):
            # v0.5.3 фикс 1: «ПРОШУ:» — по центру, без отступа (иначе
            # md-build порождает документ, не проходящий собственный validate).
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(12)
            p.paragraph_format.space_after = Pt(6)
            p.paragraph_format.line_spacing = 1.5
            _add_inline_runs(p, line)
            for r in p.runs:
                if r.text:
                    r.bold = True
            continue
        if in_header:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Inches(3.2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            if ":" in line:
                k, v = line.split(":", 1)
                r = p.add_run(_normalize_dashes(k) + ":")
                r.bold = True
                p.add_run(_normalize_dashes(v))
            else:
                p.add_run(_normalize_dashes(line))
            continue
        # --- v0.1.0 B-10 (только в зоне тела: шапка-реквизиты не тронута) ---
        if _is_table_start(lines, cur):
            i = _render_md_table(doc, lines, cur, report)
            continue
        if _LIST_ITEM_RE.match(raw.expandtabs(4)):
            j = cur + 1
            run = [(cur, raw)]
            while j < n and _LIST_ITEM_RE.match(lines[j].expandtabs(4)):
                run.append((j, lines[j]))
                j += 1
            nested = any(_expanded_indent(r) > 0 for (_, r) in run)
            if nested:
                # вложенный run: единый новый рендер (levels 2-4 + skip-маркеры)
                _render_list_run(doc, lines, run, report)
                i = j
                continue
            # плоский run: прежний построчный конвейер для КАЖДОЙ строки
            # (полная ветка-последовательность v0.9 — байт-стабильность пина)
        _add_body(doc, line)

    if not title_seen:
        raise MdBuildError(
            "титульная строка не распознана (ожидалась одна из: "
            + ", ".join(sorted(TITLE_WORDS)) + ") — fail-closed, документ не создан")
    doc.core_properties.author = author
    _ensure_run_fonts(doc)
    _stamp_table_cells(doc)
    return doc


def _ensure_run_fonts(doc):
    """v0.5.1 (приёмка по внешней рецензии): валидатор читает font.name на уровне
    run, а угадывание профиля (_guess_profile) считает доминирующий ЯВНЫЙ
    размер. Run'ы, наследующие Normal, дают name=None и пустоту по размеру:
    штампуем Times New Roman 12pt всем безымянным/безразмерным run'ам
    (явно заданный титул 14pt не трогаем)."""
    for p in doc.paragraphs:
        for r in p.runs:
            if r.font.name is None:
                r.font.name = "Times New Roman"
            if r.font.size is None:
                r.font.size = Pt(12)


def _stamp_table_cells(doc):
    """Контроль координатора v0.1.0: размер/шрифт в табличных ячейках =
    профилю, который предъявит валидатор (_guess_profile по доминирующему
    размеру абзацев тела), иначе md-build с таблицами не проходит собственную
    validate (cell_font_size/font_name)."""
    from collections import Counter
    from .spec import Profile
    sizes = Counter()
    for para in doc.paragraphs:
        for r in para.runs:
            if r.text.strip() and r.font.size is not None:
                sizes[round(r.font.size.pt)] += 1
    profile = Profile("court" if sizes and sizes.most_common(1)[0][0] == 14 else "gost")
    from docx.shared import Pt as _Pt
    want = _Pt(profile.cell_font_size_pt)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for r in para.runs:
                        r.font.name = profile.font_name
                        r.font.size = want


def md_build(input_md: str, output_docx: str, author: str) -> dict:
    path = Path(input_md)
    if not path.is_file():
        raise MdBuildError(f"input not found: {input_md}")
    md_text = path.read_text(encoding="utf-8-sig", errors="strict")
    if not md_text.strip():
        raise MdBuildError("empty md input — fail-closed")
    report: list = []
    doc = build_md_document(md_text, author, report=report)
    doc.save(output_docx)
    report_strings = [str(x) for x in report]
    return {"output_path": output_docx, "author": author,
            "blocks_rendered": len(doc.paragraphs),
            "report": report,
            "reports": report,
            "fixes": report_strings,
            "changes": report,
            "normalized": report,
            "dash_fixes": report_strings,
            "dash_normalized": len(report)}
