# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 7: встроенный компилятор математики — LaTeX -> нативный OMML.

Домен 12 ECMA-376/ISO 29500: Office Math Markup Language (namespace m:).
Word рендерит уравнения векторным интерактивным движком шрифтом Cambria Math
— никаких растровых картинок и OLE-вставок.

Публичный API:
  latex_to_omml(latex, is_block=False) -> str XML:
      is_block=False -> <m:oMath xmlns:m="...">...</m:oMath> (инлайн);
      is_block=True  -> <m:oMathPara ...><m:oMath>...</m:oMath></m:oMathPara>
                        с центрированием (m:jc=center, плюс w:jc=center в
                        рендерере). latex_to_omml_element — то же, но готовый
                        oxml-элемент для append.
  add_math_to_paragraph(paragraph, latex) — инлайн <m:oMath> в абзац;
  add_math_paragraph(doc, latex)          — выключный центрированный абзац;
  render_math(doc, block, profile)        — рендерер MathBlock (BLOCK_RENDERERS),
                                            включает подпись caption.

Поддерживаемое подмножество LaTeX: дроби \frac/\dfrac/\tfrac (m:f), корни
\sqrt[..] (m:rad), крупные операторы \sum/\prod/\int/\oint/\iint/\iiint с
пределами (m:nary), индексы ^/_ (m:sSup/m:sSub/m:sSubSup), скобки (...)/[...]
и \left...\right (m:d с эксплицитными begChr/endChr), греческие буквы
(\alpha..\Omega) и операторы (\times, \div, \pm, \le, \ge, \ne, \approx,
\infty, ...). Вложенность — неограниченная. Пробелы схлопываются как в
math-mode LaTeX; минус переводится в U+2212. Неподдерживаемая команда/символ
-> MathError (fail-closed).
"""
from __future__ import annotations

from xml.sax.saxutils import escape

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.shared import Cm

M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_MDECL = f'xmlns:m="{M_NS}"'


class MathError(ValueError):
    """Неподдерживаемый/невалидный LaTeX — fail-closed."""


# Греческие буквы: \name -> Unicode.
_GREEK = {
    "alpha": "α", "beta": "β", "gamma": "γ", "delta": "δ", "epsilon": "ε",
    "varepsilon": "ε", "zeta": "ζ", "eta": "η", "theta": "θ", "vartheta": "ϑ",
    "iota": "ι", "kappa": "κ", "lambda": "λ", "mu": "μ", "nu": "ν", "xi": "ξ",
    "omicron": "ο", "pi": "π", "varpi": "ϖ", "rho": "ρ", "varrho": "ϱ",
    "sigma": "σ", "varsigma": "ς", "tau": "τ", "upsilon": "υ", "phi": "φ",
    "varphi": "φ", "chi": "χ", "psi": "ψ", "omega": "ω",
    "Gamma": "Γ", "Delta": "Δ", "Theta": "Θ", "Lambda": "Λ", "Xi": "Ξ",
    "Pi": "Π", "Sigma": "Σ", "Upsilon": "Υ", "Phi": "Φ", "Psi": "Ψ",
    "Omega": "Ω",
}

# Символы/операторы: \name -> Unicode.
_SYMBOLS = {
    "times": "×", "div": "÷", "pm": "±", "mp": "∓", "cdot": "·",
    "ast": "∗", "star": "⋆", "circ": "∘", "bullet": "•",
    "le": "≤", "leq": "≤", "ge": "≥", "geq": "≥", "ne": "≠", "neq": "≠",
    "equiv": "≡", "approx": "≈", "sim": "∼", "cong": "≌", "propto": "∝",
    "ll": "≪", "gg": "≫", "infty": "∞", "partial": "∂", "nabla": "∇",
    "to": "→", "rightarrow": "→", "leftarrow": "←", "Rightarrow": "⇒",
    "leftrightarrow": "↔", "uparrow": "↑", "downarrow": "↓",
    "mapsto": "↦", "iff": "⇔",
    "in": "∈", "notin": "∉", "ni": "∋", "subset": "⊂", "subseteq": "⊆",
    "supset": "⊃", "supseteq": "⊇", "setminus": "∖", "emptyset": "∅",
    "cap": "∩", "cup": "∪", "wedge": "∧", "vee": "∨", "neg": "¬",
    "forall": "∀", "exists": "∃", "nexists": "∄", "top": "⊤",
    "perp": "⊥", "parallel": "∥", "angle": "∠", "triangle": "△",
    "hbar": "ℏ", "ell": "ℓ", "Re": "ℜ", "Im": "ℑ", "prime": "′",
    "cdots": "⋯", "ldots": "…", "dots": "…", "vdots": "⋮", "ddots": "⋱",
    "degree": "°", "copyright": "©", "registered": "®",
}

# МЕГА-комментарий кардинально не нужен; строка ниже — только решение о зоне.

# Крупные операторы с пределами: \name -> (символ, limLoc).
_NARY = {
    "sum": ("∑", "undOvr"), "prod": ("∏", "undOvr"), "coprod": ("∐", "undOvr"),
    "int": ("∫", "subSup"), "iint": ("∬", "subSup"), "iiint": ("∭", "subSup"),
    "oint": ("∮", "subSup"), "bigcup": ("⋃", "undOvr"), "bigcap": ("⋂", "undOvr"),
}

# Команды-операторы, на которых «операнд» n-арного оператора обязан завершиться
# (иначе \int_a^b f(x) dx захватил бы следующее слагаемое).
_NARY_STOP_CMDS = {
    "times", "div", "pm", "mp", "cdot", "ast", "star", "circ", "bullet",
    "le", "leq", "ge", "geq", "ne", "neq", "equiv", "approx", "sim", "cong",
    "propto", "ll", "gg", "to", "rightarrow", "leftarrow", "Rightarrow",
    "mapsto", "iff", "in", "notin", "ni", "subset", "subseteq", "supset",
    "supseteq", "setminus", "cap", "cup", "wedge", "vee", "neg", "prime",
    "cdots", "ldots", "dots", "vdots", "ddots", "right", "infty",
}

# Экранированные литералы: \x -> символ.
_ESCAPED = {"{": "{", "}": "}", "%": "%", "$": "$", "&": "&", "#": "#",
            "_": "_", "|": "‖", "'": "′", "`": "‵"}

# Символы, разрывающие текстовый прогон (скобки/структура разбираются отдельно).
_RUN_BREAK = "\\^_{}()[]&$#%~"


def _attr(value: str) -> str:
    """Значение XML-атрибута с экранированием кавычек."""
    return escape(value).replace('"', "&quot;")


def _run_xml(text: str) -> str:
    """Математический run: m:r/m:t; ASCII-минус -> U+2212"""
    text = text.replace("-", "\u2212")
    return f"<m:r><m:t>{escape(text)}</m:t></m:r>"


def _d_xml(inner: str, beg: str = "(", end: str = ")") -> str:
    """m:d — скобки с эксплицитными begChr/endChr (проверяемо в тестах)."""
    pr = (f'<m:dPr><m:begChr m:val="{_attr(beg)}"/>'
          f'<m:endChr m:val="{_attr(end)}"/></m:dPr>')
    return f"<m:d>{pr}<m:e>{inner}</m:e></m:d>"


def _parse_snippet(src: str) -> str:
    """Разбор вложенного фрагмента (степень корня, тело скобок)."""
    if not src.strip():
        return ""
    return _Parser(src).parse()


class _Parser:
    """Однопроходный рекурсивный спуск по LaTeX-строке -> XML-строка m:*."""

    def __init__(self, src: str) -> None:
        self.s = src
        self.i = 0

    # --- низкоуровневые -----------------------------------------------------

    def _err(self, msg: str) -> None:
        near = self.s[max(0, self.i - 15): self.i + 15]
        raise MathError(f"{msg} (позиция {self.i} рядом с {near!r})")

    def _cur(self) -> str | None:
        return self.s[self.i] if self.i < len(self.s) else None

    def _skip_ws(self) -> None:
        while self.i < len(self.s) and self.s[self.i].isspace():
            self.i += 1

    def _peek_cmd(self) -> str | None:
        """Имя (без слэша), если курсор на \буквенной команде, иначе None."""
        if self._cur() is not None and self.s[self.i] == "\\":
            j = self.i + 1
            if j < len(self.s) and self.s[j].isalpha():
                k = j
                while k < len(self.s) and self.s[k].isalpha():
                    k += 1
                return self.s[j:k]
        return None

    # --- разбор выражения -----------------------------------------------------

    def parse(self) -> str:
        body = self._expr()
        self._skip_ws()
        if self.i < len(self.s):
            self._err(f"unexpected {self._cur()!r} (лишняя }}/\\right без пары?)")
        if not body:
            self._err("пустая формула")
        return body

    def _expr(self) -> str:
        out: list[str] = []
        while True:
            self._skip_ws()
            c = self._cur()
            if c is None or c in "}&" or (c == "\\" and self._peek_cmd() == "right"):
                break
            out.append(self._item())
        return "".join(out)

    def _item(self) -> str:
        c = self._cur()
        if c == "{":
            return self._scripts(self._group())
        if c in "([":
            return self._scripts(self._paren_group(c))
        if c == "\\":
            return self._cmd()
        if c in ")]^_":
            self._err(f"unexpected {c!r}")
        return self._scripts(self._text_atom())

    def _text_atom(self) -> str:
        """Прогон текстовых символов до пробела/скобки/структуры."""
        buf: list[str] = []
        while self.i < len(self.s):
            c = self.s[self.i]
            if c.isspace() or c in _RUN_BREAK:
                break
            buf.append(c)
            self.i += 1
        if not buf:
            self._err(f"unsupported character {self._cur()!r}")
        return _run_xml("".join(buf))

    def _scripts(self, base: str) -> str:
        """Навешивает ^/_ на атом: m:sSup / m:sSub / m:sSubSup."""
        sub = sup = None
        while True:
            self._skip_ws()
            c = self._cur()
            if c == "_" and sub is None:
                self.i += 1
                sub = self._script_arg()
            elif c == "^" and sup is None:
                self.i += 1
                sup = self._script_arg()
            elif c in ("_", "^"):
                self._err("дублирующийся индекс у одного атома")
            else:
                break
        if sub is None and sup is None:
            return base
        if sub is not None and sup is not None:
            return (f"<m:sSubSup><m:e>{base}</m:e><m:sub>{sub}</m:sub>"
                    f"<m:sup>{sup}</m:sup></m:sSubSup>")
        if sub is not None:
            return f"<m:sSub><m:e>{base}</m:e><m:sub>{sub}</m:sub></m:sSub>"
        return f"<m:sSup><m:e>{base}</m:e><m:sup>{sup}</m:sup></m:sSup>"

    def _script_arg(self) -> str:
        self._skip_ws()
        c = self._cur()
        if c is None:
            self._err("нужен аргумент после ^/_")
        if c == "{":
            return self._group()
        if c == "\\":
            return self._cmd()
        if c in "([":
            return self._paren_group(c)
        self.i += 1
        return _run_xml(c)

    def _group(self) -> str:
        if self._cur() != "{":
            self._err("ожидается {")
        self.i += 1
        inner = self._expr()
        if self._cur() != "}":
            self._err("не закрыт {...}")
        self.i += 1
        return inner

    def _paren_group(self, open_ch: str) -> str:
        """(...)/[...] -> m:d (raw между парными скобками разбирается целиком)."""
        close_ch = ")" if open_ch == "(" else "]"
        inner = _parse_snippet(self._read_delimited(open_ch, close_ch))
        return _d_xml(inner)

    def _read_delimited(self, open_ch: str, close_ch: str) -> str:
        """Курсор на open_ch; сырой latex между парными скобками, курсор после close."""
        depth = 1
        self.i += 1
        start = self.i
        while self.i < len(self.s):
            c = self.s[self.i]
            if c == open_ch:
                depth += 1
            elif c == close_ch:
                depth -= 1
                if depth == 0:
                    raw = self.s[start:self.i]
                    self.i += 1
                    return raw
            self.i += 1
        self._err(f"не закрыт {open_ch!r}")

    # --- команды ------------------------------------------------------------

    def _cmd(self) -> str:
        name = self._peek_cmd()
        if name is None:
            # Одиночный символ после "\": экранированный литерал/тонкие пробелы.
            if self.i + 1 >= len(self.s):
                self._err("\\ без символа")
            ch = self.s[self.i + 1]
            self.i += 2
            if ch in ",;:! ":
                return ""
            if ch in _ESCAPED:
                return self._scripts(_run_xml(_ESCAPED[ch]))
            self._err(f"unsupported command \\{ch}")
        self.i += 1 + len(name)
        if name in _NARY:
            self._skip_ws()
            if self._cur() in ("_", "^"):
                return self._scripts(self._nary(name))
            return self._scripts(_run_xml(_NARY[name][0]))
        if name in ("frac", "dfrac", "tfrac"):
            return self._scripts(self._frac())
        if name == "sqrt":
            return self._scripts(self._sqrt())
        if name == "left":
            return self._scripts(self._left())
        if name in _GREEK:
            return self._scripts(_run_xml(_GREEK[name]))
        if name in _SYMBOLS:
            return self._scripts(_run_xml(_SYMBOLS[name]))
        self._err(f"unsupported LaTeX command \\{name}")

    def _frac(self) -> str:
        self._skip_ws()
        if self._cur() != "{":
            self._err("\\frac требует {числитель}")
        num = self._group()
        self._skip_ws()
        if self._cur() != "{":
            self._err("\\frac требует {знаменатель}")
        den = self._group()
        return f"<m:f><m:num>{num}</m:num><m:den>{den}</m:den></m:f>"

    def _sqrt(self) -> str:
        self._skip_ws()
        deg = None
        if self._cur() == "[":
            deg = _parse_snippet(self._read_delimited("[", "]"))
        self._skip_ws()
        if self._cur() != "{":
            self._err("\\sqrt требует {...}")
        e = self._group()
        if deg is None:
            return ('<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/>'
                    f"<m:e>{e}</m:e></m:rad>")
        return ('<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr>'
                f"<m:deg>{deg}</m:deg><m:e>{e}</m:e></m:rad>")

    def _left(self) -> str:
        beg = self._read_delim()
        inner = self._expr()
        if self._peek_cmd() != "right":
            self._err("\\left без \\right — fail-closed")
        self.i += 1 + len("right")
        end = self._read_delim()
        return _d_xml(inner, beg, end)

    def _read_delim(self) -> str:
        """Разделитель после \left/\right: ( [ ) ] | или . / \{ \} \| \."""
        self._skip_ws()
        c = self._cur()
        if c is None:
            self._err("нужен разделитель после \\left/\\right")
        self.i += 1
        if c != "\\":
            if c in "([|)].":
                return "" if c == "." else c
            self._err(f"unexpected delimiter {c!r}")
        nxt = self._cur()
        if nxt is None:
            self._err("нужен символ после \\left\\/\\right\\")
        self.i += 1
        if nxt == ".":
            return ""
        if nxt in ("{", "}", "|"):
            return nxt
        self._err(f"unsupported delimiter \\{nxt}")

    # --- n-арные операторы ----------------------------------------------------

    def _nary(self, name: str) -> str:
        chr_, loc = _NARY[name]
        sub = sup = None
        while True:
            self._skip_ws()
            c = self._cur()
            if c == "_" and sub is None:
                self.i += 1
                sub = self._script_arg()
            elif c == "^" and sup is None:
                self.i += 1
                sup = self._script_arg()
            elif c in ("_", "^"):
                self._err("дублирующийся предел у оператора")
            else:
                break
        operand = self._nary_operand()
        pr = f'<m:chr m:val="{chr_}"/><m:limLoc m:val="{loc}"/>'
        if sub is None:
            pr += '<m:subHide m:val="1"/>'
        if sup is None:
            pr += '<m:supHide m:val="1"/>'
        parts = [f"<m:nary><m:naryPr>{pr}</m:naryPr>"]
        if sub is not None:
            parts.append(f"<m:sub>{sub}</m:sub>")
        if sup is not None:
            parts.append(f"<m:sup>{sup}</m:sup>")
        parts.append(f"<m:e>{operand}</m:e></m:nary>")
        return "".join(parts)

    def _nary_operand(self) -> str:
        """Операнд крупного оператора: атомы до оператора/конца/группы/скобки."""
        out: list[str] = []
        while True:
            self._skip_ws()
            c = self._cur()
            if c is None or c in "+-=*/<>,;:]})~&":
                break
            if self._peek_cmd() in _NARY_STOP_CMDS:
                break
            out.append(self._item())
        return "".join(out)


# -----------------------------------------------------------------------------------
# Публичный API
# -----------------------------------------------------------------------------------

def latex_to_omml(latex: str, is_block: bool = False) -> str:
    """LaTeX -> XML-строка OMML: инлайн <m:oMath> либо выключная <m:oMathPara>
    с центрированием (m:jc="center")."""
    if not isinstance(latex, str):
        raise MathError("latex must be a string")
    body = _Parser(latex).parse()
    omath = f"<m:oMath>{body}</m:oMath>"
    if is_block:
        return (f'<m:oMathPara {_MDECL}><m:oMathParaPr><m:jc m:val="center"/>'
                f"</m:oMathParaPr>{omath}</m:oMathPara>")
    return f"<m:oMath {_MDECL}>{body}</m:oMath>"


def latex_to_omml_element(latex: str, is_block: bool = False):
    """latex_to_omml -> разобранный oxml/lxml-элемент (m:oMath|m:oMathPara)."""
    return parse_xml(latex_to_omml(latex, is_block=is_block))


def add_math_to_paragraph(paragraph, latex: str) -> None:
    """Инлайн-формула: <m:oMath> appended внутрь существующего абзаца."""
    paragraph._p.append(latex_to_omml_element(latex))


def add_math_paragraph(doc, latex: str):
    """Выключная формула: центрированный абзац с <m:oMathPara>; возвращает p."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf.first_line_indent = Cm(0)
    pf.left_indent = Cm(0)
    pf.space_before = Cm(0.2)
    pf.space_after = Cm(0.2)
    p._p.append(latex_to_omml_element(latex, is_block=True))
    return p


def render_math(doc, block, profile) -> None:
    """Рендерер MathBlock (builder.BLOCK_RENDERERS): m:oMathPara + подпись."""
    add_math_paragraph(doc, block.latex)
    if block.caption:
        from docx_cli.blocks import _set_run_font

        cap = doc.add_paragraph()
        cf = cap.paragraph_format
        cf.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cf.first_line_indent = Cm(0)
        cf.space_before = Cm(0)
        cf.space_after = Cm(0.4)
        cap_run = cap.add_run(block.caption)
        _set_run_font(cap_run, profile, profile.body_font_size_pt - 1, italic=True)


# Имена из ТЗ Пакета 7 (симметричные алиасы).
add_math_run = add_math_to_paragraph


def add_math_block(doc, latex: str):
    """add_math_paragraph из ТЗ (doc, latex) — выключная формула."""
    return add_math_paragraph(doc, latex)


def render_math_block(doc, block, profile) -> None:
    """Алиас render_math."""
    render_math(doc, block, profile)


__all__ = [
    "MathError",
    "latex_to_omml",
    "latex_to_omml_element",
    "add_math_to_paragraph",
    "add_math_paragraph",
    "add_math_run",
    "add_math_block",
    "render_math",
    "render_math_block",
]
