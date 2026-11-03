# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""S1-C1: сканер обязан видеть инструкции, спрятанные через наследуемые
и специальные каналы форматирования:
- w:docDefaults/rPrDefault с vanish;
- rPr внутри стиля АБЗАЦА (pStyle), когда у run нет собственного rStyle;
- w:specVanish на run;
- w:position со значительным смещением (текст вне строки).
"""
from __future__ import annotations

import io
import zipfile

import pytest

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def make_docx(document_xml: str, styles_xml: str | None = None) -> bytes:
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
        if styles_xml:
            z.writestr("word/styles.xml", styles_xml)
    return buf.getvalue()


def scan(tmp_path, payload: bytes) -> set:
    from docx_cli import inject_scan
    p = tmp_path / "probe.docx"
    p.write_bytes(payload)
    return {f["type"] for f in inject_scan.scan_all(str(p)).get("findings", [])}


def test_docdefaults_vanish_detected(tmp_path):
    styles = (f'<w:styles xmlns:w="{W}">'
              f'<w:docDefaults><w:rPrDefault><w:rPr><w:vanish w:val="1"/></w:rPr>'
              f'</w:rPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
              f'<w:name w:val="Normal"/></w:style></w:styles>')
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>SECRET1</w:t></w:r>'
           f'</w:p></w:body></w:document>')
    assert "hidden_text_vanish" in scan(tmp_path, make_docx(doc, styles))


def test_pstyle_rpr_vanish_detected(tmp_path):
    styles = (f'<w:styles xmlns:w="{W}">'
              f'<w:style w:type="paragraph" w:styleId="HiddenP"><w:name w:val="hidden"/>'
              f'<w:rPr><w:vanish w:val="1"/></w:rPr></w:style></w:styles>')
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:pPr><w:pStyle w:val="HiddenP"/>'
           f'</w:pPr><w:r><w:t>SECRET2</w:t></w:r></w:p></w:body></w:document>')
    assert "hidden_text_vanish" in scan(tmp_path, make_docx(doc, styles))


def test_specvanish_detected(tmp_path):
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:rPr><w:specVanish/></w:rPr>'
           f'<w:t>SECRET3</w:t></w:r></w:p></w:body></w:document>')
    assert "hidden_text_vanish" in scan(tmp_path, make_docx(doc))


def test_large_position_detected(tmp_path):
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:rPr>'
           f'<w:position w:val="-64"/></w:rPr><w:t>SECRET4</w:t></w:r>'
           f'</w:p></w:body></w:document>')
    assert "offset_text" in scan(tmp_path, make_docx(doc))


def test_legitimate_superscript_not_flagged(tmp_path):
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:rPr>'
           f'<w:position w:val="6"/><w:vertAlign w:val="superscript"/></w:rPr>'
           f'<w:t>2</w:t></w:r></w:p></w:body></w:document>')
    assert "offset_text" not in scan(tmp_path, make_docx(doc))


def test_clean_document_stays_clean(tmp_path):
    doc = (f'<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>обычный текст</w:t>'
           f'</w:r></w:p></w:body></w:document>')
    assert scan(tmp_path, make_docx(doc)) == set()
