"""v0.1.0 §3 T-05: форматтерские профили vs_sud и soyu_rayon.

Структурные скелеты, НЕ юридический текст: плейсхолдеры, никаких нормативных
«ст.NNN» в шаблонах. Профиль = имя + пути (структура как gost), юридическое
содержимое вносит legal-профиль через spec.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pytest

from docx_cli.spec import DocSpec, Profile, ProfileError

ROOT = Path(__file__).parent.parent
TPL = ROOT / "templates"

NEW_TEMPLATES = {
    "vs_sud": TPL / "vs_sud.json",
    "soyu_rayon": TPL / "soyu_rayon.json",
}


def _fill(o):
    """Тот же канал, что в v0.4.0 регресс-тесте: {cyrl-плейсхолдер} -> DEMO."""
    if isinstance(o, dict):
        return {k: _fill(v) for k, v in o.items()}
    if isinstance(o, list):
        return [_fill(x) for x in o]
    if isinstance(o, str):
        return re.sub(r"\{[а-яё_0-9]+\}", "DEMO", o)
    return o


def _spec_from_template(name: str) -> DocSpec:
    d = json.loads(NEW_TEMPLATES[name].read_text(encoding="utf-8-sig"))
    d = _fill(d)
    d["output"]["file_name"] = "t.docx"
    return DocSpec.model_validate(d)


def _synthetic_spec(profile: str) -> dict:
    return {
        "profile": profile,
        "output": {"file_name": "p.docx", "author": "Тест Т.Т."},
        "header": {"align": "right", "lines": ["В DEMO суд", "Заявитель: DEMO"]},
        "title": {"main": "ОБРАЩЕНИЕ", "subtitle": "тестовое"},
        "body": [{"type": "paragraph", "text": "Проверка профиля.",
                  "inline_markdown": False}],
    }


@pytest.mark.parametrize("profile", ["vs_sud", "soyu_rayon"])
def test_v0100_new_profile_builds_and_validates_green(profile, tmp_path):
    from docx_cli.builder import build_document
    spec = DocSpec.from_dict(_synthetic_spec(profile))
    out = tmp_path / "p.docx"
    res = build_document(spec, str(out))
    assert res["status"] == "ok", res.get("reason")
    assert out.exists()


@pytest.mark.parametrize("profile", ["vs_sud", "soyu_rayon"])
def test_v0100_new_profile_layout_equals_gost(profile):
    g = Profile("gost")
    p = Profile(profile)
    for attr in ("font_name", "body_font_size_pt", "body_line_spacing",
                 "first_line_indent_cm", "margin_left_cm", "margin_right_cm",
                 "margin_top_cm", "margin_bottom_cm", "paper_width_cm",
                 "paper_height_cm"):
        assert getattr(p, attr) == getattr(g, attr), attr
    assert p.name == profile


@pytest.mark.parametrize("name", sorted(NEW_TEMPLATES))
def test_v0100_template_parses_as_docspec_with_placeholders(name):
    spec = _spec_from_template(name)
    assert spec.profile == name
    raw = NEW_TEMPLATES[name].read_text(encoding="utf-8-sig")
    d = json.loads(raw)
    body = json.dumps(d, ensure_ascii=False)
    assert "{" in body  # плейсхолдеры заявлены
    assert any("заяв" in l or "истец" in l for l in d["header"]["lines"])
    assert spec.body[-1].type == "signature_block"
    assert any(b.type == "attachments_list" for b in spec.body)


@pytest.mark.parametrize("name", sorted(NEW_TEMPLATES))
def test_v0100_templates_have_no_normative_refs(name):
    raw = NEW_TEMPLATES[name].read_text(encoding="utf-8-sig")
    assert re.search(r"ст\.\s*\d", raw) is None
    assert "ст." not in raw  # «ст.» без номера тоже не нужен в скелете


@pytest.mark.parametrize("profile", ["vs_sud", "soyu_rayon"])
def test_v0100_built_artifact_scans_clean(profile, tmp_path):
    from docx_cli.builder import build_document
    from docx_cli.inject_scan import scan_all
    spec = DocSpec.from_dict(_synthetic_spec(profile))
    out = tmp_path / "s.docx"
    res = build_document(spec, str(out))
    assert res["status"] == "ok"
    scan = scan_all(str(out))
    assert scan["clean"] is True, scan.get("findings")


def test_v0100_back_compat_existing_profiles_pinned():
    g = Profile("gost")
    assert (g.font_name, g.body_font_size_pt, g.body_line_spacing) == \
        ("Times New Roman", 12, 1.5)
    c = Profile("court")
    assert (c.body_font_size_pt, c.body_line_spacing,
            c.default_signature_format) == (14, 1.0, "text")
    k = Profile("compact")
    assert (k.margin_left_cm, k.body_line_spacing) == (2.0, 1.15)
    with pytest.raises(ProfileError):
        Profile("no_such")
    # прежние шаблоны не тронуты: профиль court + юридические отметки на месте
    isk = json.loads((TPL / "iskovoe_zayavlenie.json").read_text(encoding="utf-8-sig"))
    assert isk["profile"] == "court"
    assert "333.21" in json.dumps(isk, ensure_ascii=False)
    for old in ("hodatajstvo.json", "otzyv_na_isk.json"):
        d = json.loads((TPL / old).read_text(encoding="utf-8-sig"))
        assert d["profile"] == "court"


def test_v0100_no_duplicate_mirovoy_profile_or_template():
    """Мировой участка — слот в soyu_rayon, отдельного шаблона нет (не дублируем)."""
    names = sorted(p.stem for p in TPL.glob("*.json"))
    assert names == ["hodatajstvo", "iskovoe_zayavlenie", "otzyv_na_isk",
                     "soyu_rayon", "vs_sud"]
    raw = NEW_TEMPLATES["soyu_rayon"].read_text(encoding="utf-8-sig")
    assert "участок" in raw  # мировой слот есть
    assert "мировой" in raw.lower()


def test_v0100_cli_profile_choices_extended():
    import argparse
    from docx_cli.cli import build_parser
    parser = build_parser()
    sub = next(a for a in parser._actions
               if isinstance(a, argparse._SubParsersAction))
    build_p = sub.choices["build"]
    prof = next(a for a in build_p._actions if a.dest == "profile")
    assert {"gost", "compact", "court", "vs_sud", "soyu_rayon"} <= \
        set(prof.choices)
