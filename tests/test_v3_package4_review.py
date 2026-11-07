# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 4 (TASK_PACKAGE4_REVIEW.md): слой рецензирования — нативные
комментарии (word/comments.xml + w:commentRangeStart/End + w:commentReference),
Track Changes (w:ins / w:del / w:delText, w:trackRevisions), CriticMarkup в
md-движке ({++ins++} / {--del--} / [фрагмент]{comment: "..."}).
Проверки структуры внутри ZIP-пакета.
"""
from __future__ import annotations

import zipfile
import xml.etree.ElementTree as ET

import pytest
from docx import Document
from docx.oxml.ns import qn

from docx_cli.comments import CommentError, add_comment
from docx_cli.errors import BuildError
from docx_cli.md_engine import md_build_v2, split_review_segments
from docx_cli.revisions import add_revision, ensure_track_revisions
from docx_cli.spec import (
    CommentSpec,
    DocSpec,
    HeaderSpec,
    OutputSpec,
    ParagraphBlock,
    TitleSpec,
)
from docx_cli.builder import build_document

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
COMMENTS_REL = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)


def _zip_part(path, name: str) -> str:
    with zipfile.ZipFile(str(path)) as zf:
        if name not in zf.namelist():
            raise AssertionError(f"part {name} missing from ZIP: {zf.namelist()}")
        return zf.read(name).decode("utf-8")


def _names(path) -> set[str]:
    with zipfile.ZipFile(str(path)) as zf:
        return set(zf.namelist())


# ------------------------------------------------------------------------------
# А. Комментарии (md-синтаксис [фрагмент]{comment: "..."} -> comments.xml)
# ------------------------------------------------------------------------------

MD_COMMENTS = """# Договор поставки

Цена договора составляет [150 000 рублей]{comment: "Проверить курс валют"} без НДС.

Срок поставки [оговорен в п. 3]{comment: 'Сослаться на спецификацию'}.
"""


def test_comments_package_structure(tmp_path):
    out = tmp_path / "comments.docx"
    res = md_build_v2(MD_COMMENTS, str(out), "pytest")
    assert res["ok"] is True, res

    # 1. Часть word/comments.xml существует в пакете
    assert "word/comments.xml" in _names(out)

    # 2. Связь на comments.xml в word/_rels/document.xml.rels
    rels = _zip_part(out, "word/_rels/document.xml.rels")
    assert "comments.xml" in rels
    assert COMMENTS_REL in rels

    # 3. Content-Type зарегистрирован в [Content_Types].xml
    assert COMMENTS_CT in _zip_part(out, "[Content_Types].xml")

    # 4. Разметка комментария в document.xml: RangeStart/End + Reference
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:commentRangeStart w:id="1"/>' in doc_xml
    assert '<w:commentRangeEnd w:id="1"/>' in doc_xml
    assert '<w:commentReference w:id="1"/>' in doc_xml
    assert '<w:commentRangeStart w:id="2"/>' in doc_xml

    # 5. Замечание и автор — в word/comments.xml
    root = ET.fromstring(_zip_part(out, "word/comments.xml"))
    assert root.tag == f"{{{W_NS}}}comments"
    by_id = {c.get(f"{{{W_NS}}}id"): c for c in root.findall(f"{{{W_NS}}}comment")}
    assert by_id["1"].get(f"{{{W_NS}}}author") == "colleague"
    assert by_id["1"].get(f"{{{W_NS}}}date")
    assert "Проверить курс валют" in "".join(by_id["1"].itertext())
    assert "Сослаться на спецификацию" in "".join(by_id["2"].itertext())

    # 6. Текст документа открыт корректно, md-разметка не осталась в тексте
    doc = Document(str(out))
    texts = [p.text for p in doc.paragraphs]
    assert any("150 000 рублей" in t for t in texts)
    assert not any("comment:" in t for t in texts)
    assert not any("{++" in t or "{--" in t for t in texts)


def test_comments_range_wraps_target_runs(tmp_path):
    out = tmp_path / "comments_span.docx"
    assert md_build_v2(MD_COMMENTS, str(out), "pytest")["ok"]
    doc = Document(str(out))
    p = next(p for p in doc.paragraphs if "150 000" in p.text)
    # Собственный текст цели остаётся в тексте параграфа
    assert "150 000 рублей" in "".join(r.text for r in p.runs)
    _p = p._p
    start = _p.findall(qn("w:commentRangeStart"))
    end = _p.findall(qn("w:commentRangeEnd"))
    assert len(start) == 1 and len(end) == 1
    # RangeStart стоит раньше RangeEnd в документе
    children = list(_p)
    assert children.index(start[0]) < children.index(end[0])


# ------------------------------------------------------------------------------
# Б. Комментарии через DocSpec.comments (CommentSpec) — API spec-сборки
# ------------------------------------------------------------------------------

def _spec_with_target(text: str, comments: list[CommentSpec]) -> DocSpec:
    return DocSpec(
        profile="corporate",
        output=OutputSpec(file_name="c.docx", author="pytest"),
        header=HeaderSpec(lines=["Подготовлено: pytest"]),
        title=TitleSpec(main="Акт"),
        body=[ParagraphBlock(type="paragraph", text=text)],
        comments=comments,
    )


def test_comment_spec_target_in_document(tmp_path):
    out = tmp_path / "spec_comment.docx"
    res = build_document(
        _spec_with_target(
            "Сумма договора 100 000 рублей уплачена.",
            [CommentSpec(target="100 000 рублей", text="Проверить НДС",
                         author="reviewer", date="2026-01-01T00:00:00Z")],
        ),
        str(out),
    )
    assert res["ok"] is True, res
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:commentRangeStart w:id="1"/>' in doc_xml
    comments_xml = _zip_part(out, "word/comments.xml")
    assert "Проверить НДС" in comments_xml
    assert 'w:author="reviewer"' in comments_xml
    assert 'w:date="2026-01-01T00:00:00Z"' in comments_xml


def test_comment_spec_target_not_found_fail_closed(tmp_path):
    out = tmp_path / "spec_comment_bad.docx"
    res = build_document(
        _spec_with_target(
            "Обычный текст.",
            [CommentSpec(target="отсутствующий фрагмент", text="Замечание")],
        ),
        str(out),
    )
    # F-07 (аудит w2): конверт error вместо исключения; fail-closed сохранён.
    assert res["ok"] is False and res["status"] == "error"
    assert "comment target not found" in res["reason"]
    assert not out.exists()


# ------------------------------------------------------------------------------
# В. add_comment: поиск target по run'ам с резкой пограничных run'ов
# ------------------------------------------------------------------------------

def test_add_comment_splits_boundary_runs():
    doc = Document()
    p = doc.add_paragraph("Всего 123 456 подлежит оплате")
    cid = add_comment(doc, p, "123 456", "Уточнить сумму")
    assert cid == 1
    # Цель вырезана в отдельный run, соседние куски — в своих run'ах
    texts = "".join(r.text for r in p.runs)
    assert texts == "Всего 123 456 подлежит оплате"
    covered = [r for r in p.runs if r.text == "123 456"]
    assert len(covered) == 1
    children = list(p._p)
    start = p._p.find(qn("w:commentRangeStart"))
    end = p._p.find(qn("w:commentRangeEnd"))
    assert children.index(start) < children.index(covered[0]._element) < children.index(end)


def test_add_comment_ids_increment():
    doc = Document()
    p = doc.add_paragraph("aaa bbb ccc")
    assert add_comment(doc, p, "aaa", "раз") == 1
    assert add_comment(doc, p, "ccc", "два") == 2


def test_add_comment_fail_closed():
    doc = Document()
    p = doc.add_paragraph("Текст.")
    with pytest.raises(CommentError):
        add_comment(doc, p, "отсутствует", "замечание")
    with pytest.raises(CommentError):
        add_comment(doc, p, "", "пустая цель")


# ------------------------------------------------------------------------------
# Г. Track Changes: md {++ins++} / {--del--} -> w:ins / w:del(w:delText)
# ------------------------------------------------------------------------------

MD_REVISIONS = """# Протокол сверки

Стоимость {--не изменилась--} {++составляет 90 000++} рублей.
"""


def test_revisions_ins_del_in_package(tmp_path):
    out = tmp_path / "revisions.docx"
    res = md_build_v2(MD_REVISIONS, str(out), "pytest")
    assert res["ok"] is True, res

    doc_xml = _zip_part(out, "word/document.xml")
    # Вставка: w:ins с автором и датой
    assert 'w:author="colleague"' in doc_xml
    assert 'w:date="' in doc_xml
    root = ET.fromstring(doc_xml)
    ins = root.iter(f"{{{W_NS}}}ins")
    ins_list = list(ins)
    dels = list(root.iter(f"{{{W_NS}}}del"))
    assert len(ins_list) == 1
    assert len(dels) == 1
    # атрибуты правок
    assert ins_list[0].get(f"{{{W_NS}}}author") == "colleague"
    assert ins_list[0].get(f"{{{W_NS}}}date")
    # текст вставки — в w:t, текст удаления — в w:delText (не w:t!)
    ins_text = "".join(t.text or "" for t in ins_list[0].iter(f"{{{W_NS}}}t"))
    assert ins_text == "составляет 90 000"
    del_text = "".join(
        t.text or "" for t in dels[0].iter(f"{{{W_NS}}}delText"))
    assert del_text == "не изменилась"
    assert "<w:delText" in doc_xml
    # у правок уникальные w:id
    assert ins_list[0].get(f"{{{W_NS}}}id") != dels[0].get(f"{{{W_NS}}}id")


def test_revisions_deleted_text_not_in_document_text(tmp_path):
    out = tmp_path / "revisions_text.docx"
    assert md_build_v2(MD_REVISIONS, str(out), "pytest")["ok"]
    doc = Document(str(out))
    p = next(p for p in doc.paragraphs if "Стоимость" in p.text)
    # delText и текст внутри w:ins не входят в видимый текст параграфа
    # (python-docx считает только прямые run'ы) — вставку проверяем по XML
    assert "не изменилась" not in p.text
    ins = p._p.find(qn("w:ins"))
    assert ins is not None
    ins_text = "".join(t.text or "" for t in ins.iter(qn("w:t")))
    assert ins_text == "составляет 90 000"


def test_add_revision_direct_api():
    doc = Document()
    p = doc.add_paragraph("Итого: ")
    add_revision(p, inserted_text="новое", deleted_text="старое",
                 author="colleague", date="2026-01-01T00:00:00Z",
                 font="Times New Roman", size_pt=12)
    doc_xml = p._p.xml
    assert 'w:author="colleague"' in doc_xml
    assert 'w:date="2026-01-01T00:00:00Z"' in doc_xml
    assert '<w:delText xml:space="preserve">старое</w:delText>' in doc_xml
    assert 'w:ascii="Times New Roman"' in doc_xml
    assert '<w:t xml:space="preserve">новое</w:t>' in doc_xml


# ------------------------------------------------------------------------------
# Д. w:trackRevisions в word/settings.xml
# ------------------------------------------------------------------------------

def test_track_revisions_in_settings(tmp_path):
    doc = Document()
    ensure_track_revisions(doc)
    ensure_track_revisions(doc)  # идемпотентно: второй вызов не дублирует
    out = tmp_path / "track.docx"
    doc.save(str(out))
    settings_xml = _zip_part(out, "word/settings.xml")
    assert "<w:trackRevisions" in settings_xml
    assert settings_xml.count("<w:trackRevisions") == 1
    Document(str(out))  # документ открывается корректно


# ------------------------------------------------------------------------------
# Е. Крит-markup сплиттер (md_engine) и смешанный абзац
# ------------------------------------------------------------------------------

def test_split_review_segments_parsing():
    segs = split_review_segments("Цена {--старая--} {++новая++}, [сумма]{comment: \"проверить\"}.")
    assert [(s.kind, s.text) for s in segs] == [
        ("text", "Цена "), ("del", "старая"), ("text", " "),
        ("ins", "новая"), ("text", ", "),
        ("comment", "сумма"), ("text", "."),
    ]
    assert segs[5].comment_text == "проверить"
    assert [s.kind for s in split_review_segments("текст {++ без конца")] == ["text"]


def test_mixed_paragraph_comments_plus_revisions(tmp_path):
    out = tmp_path / "mixed.docx"
    res = md_build_v2(
        "# Документ\n\n"
        "Ставка [18%]{comment: \"Рассмотреть НДС\"} при поставке {--была--} {++останется++}.\n",
        str(out), "reviewer-x")
    assert res["ok"] is True, res
    doc_xml = _zip_part(out, "word/document.xml")
    assert '<w:commentRangeStart w:id="1"/>' in doc_xml
    assert "<w:ins" in doc_xml and "<w:del" in doc_xml
    assert "Рассмотреть НДС" in _zip_part(out, "word/comments.xml")
    assert 'w:author="colleague"' in _zip_part(out, "word/comments.xml")
