# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""Тесты для Пакета 9: Защита документа, Многоколоночная верстка и Вложения (Домены 1, 11, 13)."""

import tempfile
import zipfile
from pathlib import Path
from docx import Document
from docx_cli.spec import DocSpec, OutputSpec, HeaderSpec, TitleSpec, SectionBreakBlock, ParagraphBlock
from docx_cli.builder import build_document
from docx_cli.md_engine import md_build_v2
from docx_cli.protection import set_document_protection
from docx_cli.attachments import embed_file_in_docx_zip, render_attachment_card


def test_document_protection_modes():
    """Проверка тега w:documentProtection в word/settings.xml."""
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "test_prot.docx"
        spec = DocSpec(
            output=OutputSpec(file_name=out.name, author="test"),
            header=HeaderSpec(lines=["Организация"]),
            title=TitleSpec(main="Тест защиты"),
            body=[ParagraphBlock(type="paragraph", text="Содержимое под защитой")],
            protection="read_only",
            protection_password="s3cret",
        )
        res = build_document(spec, str(out))
        assert res["ok"], res.get("reason")

        with zipfile.ZipFile(out) as z:
            set_xml = z.read("word/settings.xml").decode("utf-8")
            assert '<w:documentProtection' in set_xml
            # S2-H3: hash-атрибуты обязательны — иначе защита ≡ пустой пароль
            assert 'w:hashValue="' in set_xml
            assert 'w:salt="' in set_xml
            assert 'w:algorithmName="SHA-1"' in set_xml
            assert 'w:edit="readOnly"' in set_xml
            assert 'w:enforcement="1"' in set_xml


def test_multicolumn_section_layout():
    """Проверка двухколоночной верстки w:cols в w:sectPr."""
    with tempfile.TemporaryDirectory() as td:
        out = Path(td) / "test_cols.docx"
        md_text = """# Документ с колонками

Первая часть обычная одноколоночная.

<!-- section:portrait:continuous:columns:2 -->

Текст в левой колонке.
Текст в правой колонке.
"""
        res = md_build_v2(md_text, str(out), author="coordinator")
        assert res["status"] == "ok", res.get("reason")

        with zipfile.ZipFile(out) as z:
            doc_xml = z.read("word/document.xml").decode("utf-8")
            assert '<w:cols' in doc_xml
            assert 'w:num="2"' in doc_xml


def test_embed_attachment_file():
    """Проверка упаковки файла внутрь zip-структуры docx."""
    with tempfile.TemporaryDirectory() as td:
        docx_file = Path(td) / "test_attach.docx"
        data_file = Path(td) / "financial_audit.xlsx"
        data_file.write_bytes(b"PK\x03\x04fake_excel_binary_data_content")

        spec = DocSpec(
            output=OutputSpec(file_name=docx_file.name, author="test"),
            header=HeaderSpec(lines=["Организация"]),
            title=TitleSpec(main="Отчет с приложением"),
            body=[ParagraphBlock(type="paragraph", text="См. вложение ниже.")],
        )
        build_document(spec, str(docx_file))

        embed_res = embed_file_in_docx_zip(str(docx_file), str(data_file))
        assert embed_res["size_bytes"] > 0
        assert embed_res["path_in_zip"] == "word/attachments/financial_audit.xlsx"

        with zipfile.ZipFile(docx_file) as z:
            assert "word/attachments/financial_audit.xlsx" in z.namelist()
            content = z.read("word/attachments/financial_audit.xlsx")
            assert content == b"PK\x03\x04fake_excel_binary_data_content"
