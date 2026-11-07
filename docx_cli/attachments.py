# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

"""ПАКЕТ 9: Внедрение вложений (Attachments / Embedded OLE — Домен 13).

Упаковывает внешние файлы внутрь docx-архива (word/attachments/{filename})
и рендерит визуальную карточку вложения с метаданными (имя, размер, SHA-256).
"""

import hashlib
import os
import shutil
import tempfile
import zipfile
from pathlib import Path
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx_cli.errors import BuildError


def embed_file_in_docx_zip(docx_path: str, file_path: str,
                           attachment_name: str | None = None) -> dict:
    """Упаковывает физический файл внутрь скомпилированного .docx архива.

    S2-M3/S2-H2: пакет должен остаться валидным OPC — добавляем Override в
    [Content_Types].xml и relationship из document part. Иначе Word показывает
    диалог восстановления (сирота без Content-Type/relationship).
    """
    src = Path(file_path)
    if not src.exists() or not src.is_file():
        raise BuildError(f"attachment file not found: {file_path}")

    target_name = attachment_name or src.name
    # S2-M3: имя вложения — только basename без path-разделителей
    import re as _re
    if not _re.fullmatch(r"[A-Za-z0-9А-Яа-я._+-]{1,128}", target_name):
        raise BuildError(f"invalid attachment name: {target_name!r}")
    data = src.read_bytes()
    sha256 = hashlib.sha256(data).hexdigest()
    size_bytes = len(data)

    target_zip_path = f"word/attachments/{target_name}"

    # Перепаковываем zip
    temp_fd, temp_path = tempfile.mkstemp(suffix=".docx")
    os.close(temp_fd)

    ct_override = (
        f'<Override PartName="/{target_zip_path}" '
        f'ContentType="application/octet-stream"/>'
    )
    with zipfile.ZipFile(docx_path, "r") as zin, \
         zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == target_zip_path:
                continue  # замена существующего вложения с тем же именем
            if item.filename == "[Content_Types].xml":
                ct = zin.read(item.filename).decode("utf-8")
                if f'PartName="/{target_zip_path}"' not in ct:
                    ct = ct.replace("</Types>", ct_override + "</Types>")
                zout.writestr(item, ct.encode("utf-8"))
                continue
            if item.filename == "word/_rels/document.xml.rels":
                rels = zin.read(item.filename).decode("utf-8")
                rid = f"rIdAttach{abs(hash(target_zip_path)) % 100000}"
                if f'Target="attachments/{target_name}"' not in rels:
                    rel = (f'<Relationship Id="{rid}" '
                           f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachment" '
                           f'Target="attachments/{target_name}"/>')
                    rels = rels.replace("</Relationships>", rel + "</Relationships>")
                zout.writestr(item, rels.encode("utf-8"))
                continue
            zout.writestr(item, zin.read(item.filename))
        zout.writestr(target_zip_path, data)

    shutil.move(temp_path, docx_path)

    return {
        "name": target_name,
        "path_in_zip": target_zip_path,
        "size_bytes": size_bytes,
        "sha256": sha256,
    }


def render_attachment_card(doc, file_path: str, title: str | None = None,
                           profile=None) -> None:
    """Рендерит карточку вложения в теле документа."""
    src = Path(file_path)
    if not src.exists():
        # Если файл пока не существует при черновой сборке
        filename = src.name
        size_str = "вложение"
        sha_str = ""
    else:
        filename = src.name
        size_bytes = src.stat().st_size
        size_str = f"{size_bytes / 1024:.1f} КБ" if size_bytes >= 1024 else f"{size_bytes} байт"
        sha_str = f"SHA-256: {hashlib.sha256(src.read_bytes()).hexdigest()[:16]}..."

    tbl = doc.add_table(rows=1, cols=1)
    tbl.autofit = False
    cell = tbl.cell(0, 0)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.left_indent = Cm(0.3)
    p.paragraph_format.first_line_indent = Cm(0)

    # Иконка скрепки / вложения
    r_icon = p.add_run("📎 ")
    r_icon.font.size = Pt(12)

    label = title or filename
    r_title = p.add_run(f"Вложение: {label}\n")
    r_title.bold = True
    r_title.font.size = Pt(10)

    details = f"Файл: {filename} ({size_str})"
    if sha_str:
        details += f" | {sha_str}"
    r_det = p.add_run(details)
    r_det.font.size = Pt(8.5)
    r_det.font.color.rgb = RGBColor(100, 100, 100)
