# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S1-H3: белый текст гасится только если run реально внутри тёмной ячейки.

Раньше фильтр сверял подстроку глобально: слово «ШПИОН» белым по белому
в обычном абзаце пропускалось, если где-то в тёмной ячейке таблицы
встречалась та же подстрока.
"""
from __future__ import annotations

import io
import zipfile

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def make_docx(document_xml: str) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr("[Content_Types].xml", """<?xml version="1.0"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>""")
        z.writestr("_rels/.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>""")
        z.writestr("word/_rels/document.xml.rels", """<?xml version="1.0"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>""")
        z.writestr("word/document.xml", document_xml)
        z.writestr("word/styles.xml", f'<w:styles xmlns:w="{W}"/>')
    return buf.getvalue()


def test_white_text_in_light_para_not_suppressed(tmp_path):
    from docx_cli import inject_scan
    doc = (f'<w:document xmlns:w="{W}"><w:body>'
           f'<w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>ШПИОН</w:t></w:r></w:p>'
           f'<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="000000"/></w:tcPr>'
           f'<w:p><w:r><w:t>в тёмной ячейке слово ШПИОН легально</w:t></w:r></w:p>'
           f'</w:tc></w:tr></w:tbl>'
           f'</w:body></w:document>')
    p = tmp_path / "h3.docx"
    p.write_bytes(make_docx(doc))
    types = {f["type"] for f in inject_scan.scan_all(str(p))["findings"]}
    assert "white_text_candidate" in types or "white_text_on_light_bg" in types, types


def test_white_text_in_dark_cell_still_suppressed(tmp_path):
    from docx_cli import inject_scan
    doc = (f'<w:document xmlns:w="{W}"><w:body>'
           f'<w:tbl><w:tr><w:tc><w:tcPr><w:shd w:val="clear" w:fill="000000"/></w:tcPr>'
           f'<w:p><w:r><w:rPr><w:color w:val="FFFFFF"/></w:rPr><w:t>белый по тёмному</w:t></w:r></w:p>'
           f'</w:tc></w:tr></w:tbl>'
           f'</w:body></w:document>')
    p = tmp_path / "h3b.docx"
    p.write_bytes(make_docx(doc))
    types = {f["type"] for f in inject_scan.scan_all(str(p))["findings"]}
    assert "white_text_candidate" not in types, types
