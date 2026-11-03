# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 2 (TASK_PACKAGE2_IMAGES.md): изображения DrawingML — ImageBlock,
render_image (масштабирование/выравнивание/подписи), md-синтаксис
`![caption](path)`, fail-closed при отсутствии файла. Проверки структуры
внутри ZIP-пакета (word/media, rels, w:drawing).
"""
from __future__ import annotations

import struct
import zipfile
import zlib
import xml.etree.ElementTree as ET
from pathlib import Path

import pytest

from docx_cli.blocks import render_image
from docx_cli.builder import build_document
from docx_cli.errors import BuildError
from docx_cli.md_engine import md_build_v2
from docx_cli.spec import DocSpec, ImageBlock, Profile

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
IMG_REL_NS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)


def _make_png(path, width: int = 300, height: int = 150) -> None:
    """Минимальный валидный PNG (RGB) без PIL."""

    def chunk(tag: bytes, data: bytes) -> bytes:
        return (
            struct.pack(">I", len(data)) + tag + data
            + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)
        )

    ihdr = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)
    raw = b"".join(b"\x00" + b"\xA0" * width * 3 for _ in range(height))
    payload = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", ihdr)
        + chunk(b"IDAT", zlib.compress(raw))
        + chunk(b"IEND", b"")
    )
    Path(path).write_bytes(payload)


def _zip_part(path, name: str) -> str:
    with zipfile.ZipFile(str(path)) as zf:
        return zf.read(name).decode("utf-8")


def _names(path) -> set[str]:
    with zipfile.ZipFile(str(path)) as zf:
        return set(zf.namelist())


def _spec_dict(body: list[dict]) -> dict:
    return {
        "profile": "gost",
        "output": {"file_name": "test.docx", "author": "Тестов Т.Т."},
        "header": {"align": "right", "lines": ["В тестовый суд", "Истец: Тестов Т.Т."]},
        "title": {"main": "Тестовый документ"},
        "body": body,
    }


# ------------------------------------------------------------------------------
# 1. Сборка с реальной картинкой: media, rels, w:drawing, подпись
# ------------------------------------------------------------------------------

def test_image_package_structure(tmp_path):
    img = tmp_path / "chart.png"
    _make_png(img)
    out = tmp_path / "out.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img), "caption": "Рисунок 1: Тестовый график"},
    ]))
    res = build_document(spec, str(out))
    assert res["ok"] is True, res

    # 1. Файл изображения внутри word/media/
    names = _names(out)
    media = [n for n in names if n.startswith("word/media/")]
    assert media, f"no word/media/ in {names}"
    assert any(n.endswith("image1.png") for n in media), media

    # 2. Связь image в word/_rels/document.xml.rels
    rels = _zip_part(out, "word/_rels/document.xml.rels")
    assert IMG_REL_NS in rels
    assert "media/" in rels

    # 3. Content-Type для png
    ct = _zip_part(out, "[Content_Types].xml")
    assert "png" in ct

    # 4. <w:drawing> в document.xml
    doc_xml = _zip_part(out, "word/document.xml")
    assert "<w:drawing>" in doc_xml

    # 5. Подпись присутствует, курсив
    assert "Рисунок 1: Тестовый график" in doc_xml
    root = ET.fromstring(doc_xml)
    caption_run = None
    for r in root.iter(f"{{{W_NS}}}r"):
        if "Рисунок 1" in "".join(t.text or "" for t in r.findall(f"{{{W_NS}}}t")):
            rpr = r.find(f"{{{W_NS}}}rPr")
            assert rpr is not None
            assert rpr.find(f"{{{W_NS}}}i") is not None, "caption must be italic"
            caption_run = r
    assert caption_run is not None


# ------------------------------------------------------------------------------
# 2. Масштабирование: extent не вылезает за рабочую область
# ------------------------------------------------------------------------------

def _extents(path) -> list[tuple[int, int]]:
    doc_xml = _zip_part(path, "word/document.xml")
    root = ET.fromstring(doc_xml)
    out = []
    for ext in root.iter("{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"):
        out.append((int(ext.get("cx")), int(ext.get("cy"))))
    return out


def test_image_auto_scale_fits_margins(tmp_path):
    # Картинка 40x20 см (при 96 dpi) — шире рабочей области A4 (16.5 см gost).
    img = tmp_path / "wide.png"
    _make_png(img, width=1512, height=756)
    out = tmp_path / "out.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img)},
    ]))
    res = build_document(spec, str(out))
    assert res["ok"] is True, res

    profile = Profile("gost")
    avail_cm = profile.paper_width_cm - profile.margin_left_cm - profile.margin_right_cm
    # twips-округление Word: реальная геометрия раздела чуть шире номинала,
    # допуск 0.01 см.
    avail_emu = int(avail_cm * 360000) + 3600
    for cx, _cy in _extents(out):
        assert cx <= avail_emu, f"image cx={cx} exceeds avail {avail_emu}"

    # Точная проверка против фактической геометрии раздела в собранном файле.
    from docx import Document
    sec = Document(str(out)).sections[-1]
    real_avail = (sec.page_width.cm - sec.left_margin.cm - sec.right_margin.cm) * 360000
    for cx, _cy in _extents(out):
        assert cx <= real_avail + 1, f"image cx={cx} exceeds real avail {real_avail}"


def test_image_small_keeps_natural_width(tmp_path):
    # Маленькая картинка (5 см) не растягивается до всей ширины.
    img = tmp_path / "small.png"
    _make_png(img, width=189, height=94)  # ~5 см при 96 dpi
    out = tmp_path / "out.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img)},
    ]))
    res = build_document(spec, str(out))
    assert res["ok"] is True, res
    cx, _ = _extents(out)[0]
    assert cx < int(7.0 * 360000)  # заметно меньше 16.5 см рабочей области


def test_image_width_cm_and_pct(tmp_path):
    img = tmp_path / "img.png"
    _make_png(img)
    avail_cm = 21.0 - 3.0 - 1.5

    out = tmp_path / "w_cm.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img), "width_cm": 10.0},
    ]))
    assert build_document(spec, str(out))["ok"] is True
    cx, _ = _extents(out)[0]
    assert abs(cx - 10.0 * 360000) < 3600  # ±0.01 см

    out2 = tmp_path / "w_pct.docx"
    spec2 = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img), "width_pct": 50.0},
    ]))
    assert build_document(spec2, str(out2))["ok"] is True
    cx2, _ = _extents(out2)[0]
    assert abs(cx2 - avail_cm * 0.5 * 360000) < 3600

    # width_cm больше рабочей области — cap по avail
    out3 = tmp_path / "w_cap.docx"
    spec3 = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img), "width_cm": 99.0},
    ]))
    assert build_document(spec3, str(out3))["ok"] is True
    cx3, _ = _extents(out3)[0]
    assert cx3 <= int(avail_cm * 360000) + 3600


def test_image_align_left(tmp_path):
    img = tmp_path / "img.png"
    _make_png(img)
    out = tmp_path / "out.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(img), "align": "left"},
    ]))
    res = build_document(spec, str(out))
    assert res["ok"] is True, res
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:jc w:val="left"/>' in doc_xml


def test_render_image_direct(tmp_path):
    img = tmp_path / "img.png"
    _make_png(img)
    from docx import Document
    doc = Document()
    render_image(doc, ImageBlock(path=str(img), caption="Схема"), Profile("gost"))
    assert any("Схема" in p.text for p in doc.paragraphs)


# ------------------------------------------------------------------------------
# 3. Markdown-синтаксис ![caption](path)
# ------------------------------------------------------------------------------

def test_md_build_with_image(tmp_path):
    img = tmp_path / "test.png"
    _make_png(img)
    out = tmp_path / "md.docx"
    md = "# Отчет\n\n![Схема процесса](test.png)\n\nТекст после рисунка."
    res = md_build_v2(md, str(out), "pytest", base_dir=str(tmp_path))
    assert res["ok"] is True, res
    doc_xml = _zip_part(out, "word/document.xml")
    assert "<w:drawing>" in doc_xml
    assert "Схема процесса" in doc_xml


def test_md_build_image_no_caption(tmp_path):
    img = tmp_path / "photo.jpg"
    _make_png(img)  # расширение .jpg, содержимое PNG — python-docx сниффает байты
    out = tmp_path / "md.docx"
    md = "# Отчет\n\n![](photo.jpg)\n\nТекст."
    res = md_build_v2(md, str(out), "pytest", base_dir=str(tmp_path))
    assert res["ok"] is True, res
    doc_xml = _zip_part(out, "word/document.xml")
    assert "<w:drawing>" in doc_xml


# ------------------------------------------------------------------------------
# 4. Fail-closed: несуществующий файл
# ------------------------------------------------------------------------------

def test_image_missing_file_fail_closed(tmp_path):
    out = tmp_path / "out.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "image", "path": str(tmp_path / "nope.png")},
    ]))
    # F-07 (аудит w2): ошибка рендерера возвращается конвертом, не исключением.
    res = build_document(spec, str(out))
    assert res["ok"] is False and res["status"] == "error"
    assert "not found" in res["reason"]
    # Битый архив не остался
    assert not out.exists()


def test_image_missing_file_md_fail_closed(tmp_path):
    out = tmp_path / "md.docx"
    md = "# Отчет\n\n![Нет файла](absent.png)\n\nТекст."
    # F-07 (аудит w2): конверт error вместо исключения.
    res = md_build_v2(md, str(out), "pytest")
    assert res["ok"] is False and res["status"] == "error"
    assert "not found" in res["reason"]
    assert not out.exists()


# ------------------------------------------------------------------------------
# 5. Регрессия: базовая сборка без картинок не сломана
# ------------------------------------------------------------------------------

def test_regression_no_images(tmp_path):
    out = tmp_path / "plain.docx"
    spec = DocSpec.model_validate(_spec_dict([
        {"type": "paragraph", "text": "Обычный абзац документа."},
    ]))
    res = build_document(spec, str(out))
    assert res["ok"] is True, res
    assert "word/media" not in " ".join(_names(out))
