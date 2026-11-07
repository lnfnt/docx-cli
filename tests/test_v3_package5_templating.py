# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Пакет 5 (TASK_PACKAGE5_TEMPLATING.md): движок шаблонизации.

Проверки: дробленые run'ы (run-reconciliation с сохранением Bold),
вложенные ключи, табличные циклы {{#items}}, условные блоки {{#if}},
колонтитулы, сноски, strict fail-closed, CLI docx-cli render.
"""
from __future__ import annotations

import json
import re
import zipfile

import pytest
from docx import Document

from conftest import run_cli
from docx_cli.errors import TemplateError
from docx_cli.templating import render_template


def _text(doc):
    return "\n".join(p.text for p in doc.paragraphs)


def _cell_texts(doc, table_idx=0):
    return [c.text for row in doc.tables[table_idx].rows for c in row.cells]


# ------------------------------------------------------------------------------
# 3.1 Дробленые run'ы: {{company_name}} в 3 run'ах, Bold сохраняется
# ------------------------------------------------------------------------------

def _make_template(path, runs, bold=False):
    doc = Document()
    p = doc.add_paragraph()
    for chunk in runs:
        r = p.add_run(chunk)
        r.bold = bold
    doc.save(str(path))
    return path


def test_split_runs_reconciliation_keeps_bold(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["{{", "company_", "name}}"], bold=True)
    out = tmp_path / "out.docx"
    res = render_template(str(tmpl), {"company_name": "ООО Ромашка"}, str(out))
    assert res["ok"] is True and res["status"] == "ok", res

    doc2 = Document(str(out))
    p2 = doc2.paragraphs[0]
    assert p2.text == "ООО Ромашка"
    assert all(r.bold for r in p2.runs if r.text)
    assert "{{" not in _text(doc2) and "}}" not in _text(doc2)
    assert res["replaced"].get("company_name") == 1


def test_split_runs_fragmented_by_empty_and_styled_runs(tmp_path):
    doc = Document()
    p = doc.add_paragraph()
    r1 = p.add_run("Контрагент: {{")
    r1.bold = True
    p.add_run("")  # пустой run в середине
    p.add_run("cli")
    r4 = p.add_run("ent")
    r4.italic = True
    r5 = p.add_run("_name}} Договор")
    r5.italic = True
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    out = tmp_path / "out.docx"
    render_template(str(tmpl), {"client_name": "ООО Ромашка"}, str(out))
    doc2 = Document(str(out))
    assert doc2.paragraphs[0].text == "Контрагент: ООО Ромашка Договор"
    assert "{{" not in _text(doc2)


# ------------------------------------------------------------------------------
# 3.2 Вложенные объекты
# ------------------------------------------------------------------------------

def test_nested_keys(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx",
                          ["Город: {{order.customer.address.city}}, ИНН {{client.inn}}"])
    out = tmp_path / "out.docx"
    render_template(
        str(tmpl),
        {"order": {"customer": {"address": {"city": "Тверь"}}},
         "client": {"inn": "7801123456"}},
        str(out))
    assert "Город: Тверь, ИНН 7801123456" in _text(Document(str(out)))


# ------------------------------------------------------------------------------
# 3.3 Табличные циклы
# ------------------------------------------------------------------------------

def test_table_loop_clones_rows(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=4)
    headers = ["Товар", "Кол-во", "Цена", "Сумма"]
    for c, h in zip(table.rows[0].cells, headers):
        c.text = h
    markers = ["{{#items}}", "{{name}} × {{qty}}", "{{price}}", "{{/items}}"]
    for c, m in zip(table.rows[1].cells, markers):
        c.text = m
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    items = [{"name": "Товар А", "qty": 2, "price": "100"},
             {"name": "Товар Б", "qty": 1, "price": "250"},
             {"name": "Товар В", "qty": 5, "price": "10"}]
    out = tmp_path / "out.docx"
    res = render_template(str(tmpl), {"items": items}, str(out))
    assert res["rows_cloned"] == 3, res

    doc2 = Document(str(out))
    assert len(doc2.tables[0].rows) == 4  # шапка + 3 строки товаров
    texts = _cell_texts(doc2)
    assert "Товар А × 2" in texts
    assert "Товар Б × 1" in texts
    assert "Товар В × 5" in texts
    assert not any("{{" in t for t in texts), texts


def test_table_loop_empty_list_removes_template_row(tmp_path):
    doc = Document()
    table = doc.add_table(rows=2, cols=2)
    table.rows[0].cells[0].text = "Счет"
    table.rows[1].cells[0].text = "{{#items}}"
    table.rows[1].cells[1].text = "{{name}}{{/items}}"
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    out = tmp_path / "out.docx"
    render_template(str(tmpl), {"items": []}, str(out))
    assert len(Document(str(out)).tables[0].rows) == 1


# ------------------------------------------------------------------------------
# 3.5 (частично) Условные блоки
# ------------------------------------------------------------------------------

def test_if_block_inline_truthy_and_falsy(tmp_path):
    tmpl = _make_template(
        tmp_path / "t.docx",
        ["Скидка: {{#if has_discount}}есть, {{discount}}%{{/if}}. Конец."])
    out_t = tmp_path / "true.docx"
    render_template(str(tmpl), {"has_discount": True, "discount": 15}, str(out_t))
    t = _text(Document(str(out_t)))
    assert "Скидка: есть, 15%. Конец." == t

    out_f = tmp_path / "false.docx"
    render_template(str(tmpl), {"has_discount": False}, str(out_f))
    assert _text(Document(str(out_f))) == "Скидка: . Конец."


def test_if_block_across_paragraphs_removed_when_falsy(tmp_path):
    doc = Document()
    doc.add_paragraph("До {{#if flag}}")
    doc.add_paragraph("внутри блока {{flag_val}}")
    doc.add_paragraph("{{/if}} После")
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    out_f = tmp_path / "false.docx"
    render_template(str(tmpl), {"flag": False}, str(out_f))
    doc_f = Document(str(out_f))
    texts = [p.text for p in doc_f.paragraphs if p.text.strip()]
    assert texts == ["До ", " После"], texts
    assert "{{" not in "\n".join(texts)

    out_t = tmp_path / "true.docx"
    render_template(str(tmpl), {"flag": True, "flag_val": "значение"}, str(out_t))
    doc_t = Document(str(out_t))
    texts_t = [p.text for p in doc_t.paragraphs if p.text.strip()]
    assert texts_t == ["До ", "внутри блока значение", " После"], texts_t


# ------------------------------------------------------------------------------
# 3.4 Колонтитулы и сноски
# ------------------------------------------------------------------------------

def test_header_footer_replacement(tmp_path):
    doc = Document()
    section = doc.sections[0]
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    section.header.paragraphs[0].add_run("Договор № {{doc_number}}")
    section.footer.paragraphs[0].add_run("стр. {{client.name}}")
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    out = tmp_path / "out.docx"
    res = render_template(
        str(tmpl), {"doc_number": "Д-45/2026", "client": {"name": "Ромашка"}}, str(out))

    doc2 = Document(str(out))
    h2 = doc2.sections[0].header
    f2 = doc2.sections[0].footer
    assert not h2.is_linked_to_previous and not f2.is_linked_to_previous
    assert h2.paragraphs[0].text == "Договор № Д-45/2026"
    assert f2.paragraphs[0].text == "стр. Ромашка"
    assert res["replaced"].get("doc_number") == 1


def test_footnotes_replacement(tmp_path):
    """Сноска с плейсхолдером, разрезанным на 2 run'а, правится внутри пакета
    (python-docx footnotes не раскрывает — patch на уровне ZIP)."""
    doc = Document()
    doc.add_paragraph("Текст со сноской.")
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))

    W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    ft = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes xmlns:w="{W}">'
        '<w:footnote w:type="separator" w:id="-1">'
        '<w:p><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0">'
        '<w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>'
        '<w:footnote w:id="1"><w:p>'
        '<w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>'
        '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>'
        '<w:r><w:t xml:space="preserve"> Сумма: {{doc_</w:t></w:r>'
        '<w:r><w:t>number}} руб.</w:t></w:r>'
        '</w:p></w:footnote>'
        '</w:footnotes>'
    )
    with zipfile.ZipFile(str(tmpl)) as src:
        items = [(i, src.read(i.filename)) for i in src.infolist()]
    RECOG = "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"
    with zipfile.ZipFile(str(tmpl), "w") as zf:
        for info, blob in items:
            if info.filename == "word/_rels/document.xml.rels":
                blob = blob.replace(
                    b"</Relationships>",
                    (b'<Relationship Id="rIdFootnotes1" '
                     b'Type="http://schemas.openxmlformats.org/officeDocument/2006/'
                     b'relationships/footnotes" Target="footnotes.xml"/>'
                     b"</Relationships>"))
            if info.filename == "[Content_Types].xml":
                blob = blob.replace(
                    b"</Types>",
                    f'<Override PartName="/word/footnotes.xml" ContentType="{RECOG}"/>'
                    .encode() + b"</Types>")
            zf.writestr(info, blob)
        zf.writestr("word/footnotes.xml", ft)

    out = tmp_path / "out.docx"
    render_template(str(tmpl), {"doc_number": "77"}, str(out))
    after = zipfile.ZipFile(str(out)).read("word/footnotes.xml").decode("utf-8")
    joined = "".join(re.findall(r"<w:t[^>]*>([^<]*)</w:t>", after))
    assert joined == " Сумма: 77 руб."
    assert "{{doc_number}}" not in after
    assert "{{" not in joined


# ------------------------------------------------------------------------------
# 3.5 Strict fail-closed
# ------------------------------------------------------------------------------

def test_strict_missing_variable_raises(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["ИНН {{client.inn}}"])
    with pytest.raises(TemplateError) as exc:
        render_template(str(tmpl), {"client": {}}, str(tmp_path / "out.docx"), strict=True)
    assert "missing variable 'client.inn'" in str(exc.value)


def test_non_strict_missing_keeps_placeholder_degraded(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["ИНН {{client.inn}}"])
    out = tmp_path / "out.docx"
    res = render_template(str(tmpl), {}, str(out))
    assert res["status"] == "degraded" and res["missing"] == ["client.inn"]
    assert _text(Document(str(out))).endswith("{{client.inn}}")


def test_strict_missing_loop_variable_raises(tmp_path):
    doc = Document()
    table = doc.add_table(rows=1, cols=1)
    table.rows[0].cells[0].text = "{{#items}}{{/items}}"
    tmpl = tmp_path / "t.docx"
    doc.save(str(tmpl))
    with pytest.raises(TemplateError) as exc:
        render_template(str(tmpl), {}, str(tmp_path / "out.docx"), strict=True)
    assert "missing variable 'items'" in str(exc.value)


def test_respects_output_written_only_on_strict_pass(tmp_path):
    """Строгий режим не пишет битый output: отсутствующая переменная -> ошибка до save."""
    tmpl = _make_template(tmp_path / "t.docx", ["X {{a}} Y {{b}}"])
    with pytest.raises(TemplateError):
        render_template(str(tmpl), {"a": "1"}, str(tmp_path / "out.docx"), strict=True)


# ------------------------------------------------------------------------------
# CLI: docx-cli render
# ------------------------------------------------------------------------------

def test_cli_render_happy_path(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["Сумма {{payment.amount}} руб."])
    data = tmp_path / "data.json"
    data.write_text(json.dumps({"payment": {"amount": 15000}}), encoding="utf-8")
    out = tmp_path / "result.docx"
    cp = run_cli("render", str(tmpl), str(data), "-o", str(out), "--json")
    assert cp.returncode == 0, cp.stderr
    env = json.loads(cp.stdout)
    assert env["ok"] is True and env["status"] == "ok"
    assert _text(Document(str(out))) == "Сумма 15000 руб."


def test_cli_render_strict_fails_closed(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["X {{missing_var}}"])
    data = tmp_path / "data.json"
    data.write_text("{}", encoding="utf-8")
    cp = run_cli("render", str(tmpl), str(data), "-o", str(tmp_path / "o.docx"), "--strict")
    assert cp.returncode == 1
    env = json.loads(cp.stdout)
    assert env["ok"] is False and "missing" in env["reason"]


def test_cli_render_non_strict_degraded(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["X {{missing_var}}"])
    data = tmp_path / "data.json"
    data.write_text("{}", encoding="utf-8")
    out = tmp_path / "o.docx"
    cp = run_cli("render", str(tmpl), str(data), "-o", str(out))
    # S2-M2: render degraded = exit 2 (едина со build)
    assert cp.returncode == 2
    env = json.loads(cp.stdout)
    assert env["status"] == "degraded" and env["missing"] == ["missing_var"]
    assert _text(Document(str(out))) == "X {{missing_var}}"


def test_cli_render_stdin_data(tmp_path):
    tmpl = _make_template(tmp_path / "t.docx", ["№ {{doc_number}}"])
    out = tmp_path / "o.docx"
    import subprocess
    import sys
    cp = subprocess.run(
        [sys.executable, "-m", "docx_cli", "render", str(tmpl), "-", "-o", str(out)],
        input=json.dumps({"doc_number": "А-1"}), capture_output=True, text=True)
    assert cp.returncode == 0, cp.stderr
    assert _text(Document(str(out))) == "№ А-1"
