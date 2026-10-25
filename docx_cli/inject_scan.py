"""docx-cli scan — аудит входящих .docx на LLM-атаки (скрытые векторы).

v0.3.0 (2026-10-20): собственная clean-room реализация эвристик аудита
входящих документов. Дополняет guard-слои обработки текста: те изолируют
ПРОЧИТАННОЕ (spotlight), этот находит
СКРЫТОЕ ДО обработки:

- скрытый текст w:vanish;
- микрошрифт (<= 2pt в w:sz);
- белый текст на светлом/прозрачном фоне (с фильтром тёмных ячеек таблиц);
- текст внутри отслеживаемых правок w:ins / w:del;
- скрытые инструкции в word/comments.xml.

Fail-closed: битый файл/не-docx = findings не пуст и clean=False.
v0.6.0: стиль-наследование (rStyle/basedOn), near-white/themeColor,
все word/*.xml части (headers/footers/footnotes/endnotes), fail-closed
конверты для не-docx/битого XML (внешний аудит 21.10, векторы F1/F2/F3)."""
from __future__ import annotations

import copy
import re
import zipfile
from xml.etree import ElementTree as ET

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

def _extract_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(f"{{{W}}}t"))


def _has_dark_cell_shading(tc) -> bool:
    shd = tc.find(f".//{{{W}}}shd")
    if shd is None:
        return False
    fill = (shd.get(f"{{{W}}}fill") or "").lower()
    return fill not in ("", "auto", "ffffff")


def _styles_map(styles) -> dict:
    if styles is None:
        return {}
    if isinstance(styles, dict):
        return styles
    out = {}
    try:
        for s in styles.iter(f"{{{W}}}style"):
            sid = s.get(f"{{{W}}}styleId")
            if sid is not None:
                out[sid] = s
    except Exception:  # noqa: BLE001
        pass
    return out


def _style_chain(styles_map: dict, style_id) -> list:
    chain = []
    cur = style_id
    seen = set()
    while cur is not None and cur not in seen and len(chain) < 4:
        seen.add(cur)
        chain.append(cur)
        st = styles_map.get(cur)
        if st is None:
            break
        bo = st.find(f"{{{W}}}basedOn")
        cur = bo.get(f"{{{W}}}val") if bo is not None else None
    chain.reverse()
    return chain


def _effective_rpr(styles_map: dict, run_rpr, style_id) -> "ET.Element":
    erpr = ET.Element(f"{{{W}}}rPr")
    for sid in _style_chain(styles_map, style_id):
        st = styles_map.get(sid)
        if st is None:
            continue
        srpr = st.find(f"{{{W}}}rPr")
        if srpr is None:
            continue
        for child in srpr:
            existing = erpr.find(child.tag)
            if existing is not None:
                erpr.remove(existing)
            erpr.append(copy.deepcopy(child))
    if run_rpr is not None:
        for child in run_rpr:
            existing = erpr.find(child.tag)
            if existing is not None:
                erpr.remove(existing)
            erpr.append(copy.deepcopy(child))
    return erpr


def _scan_runs(root, *args, path="document.xml", styles=None, styles_map=None,
               findings=None, **kwargs) -> None:
    _findings = findings
    _styles = styles if styles is not None else styles_map
    _path = path
    for a in args:
        if isinstance(a, list):
            _findings = a
        elif isinstance(a, str):
            _path = a
        else:
            if _styles is None:
                _styles = a
            elif _findings is None:
                _findings = a
    if _findings is None:
        _findings = kwargs.get("findings", [])
    if _path is None:
        _path = kwargs.get("path", "document.xml")
    smap = _styles_map(_styles)
    if _findings is None:
        return

    for r in root.iter(f"{{{W}}}r"):
        rpr = r.find(f"{{{W}}}rPr")
        style_id = None
        rs = None
        if rpr is not None:
            rs = rpr.find(f"{{{W}}}rStyle")
        if rs is None:
            rs = r.find(f"{{{W}}}rStyle")
        if rs is not None:
            style_id = rs.get(f"{{{W}}}val")
        eff = _effective_rpr(smap, rpr, style_id)
        text = _extract_text(r)
        if not text:
            continue

        vanish = eff.find(f"{{{W}}}vanish")
        if vanish is not None and vanish.get(f"{{{W}}}val", "1") not in ("0", "false", "off"):
            _findings.append({"type": "hidden_text_vanish", "text": text, "path": _path})

        sz = eff.find(f"{{{W}}}sz")
        szcs = eff.find(f"{{{W}}}szCs")
        micro_val = None
        if sz is not None and (sz.get(f"{{{W}}}val") or "").isdigit() and int(sz.get(f"{{{W}}}val")) <= 4:
            micro_val = sz.get(f"{{{W}}}val")
        elif szcs is not None and (szcs.get(f"{{{W}}}val") or "").isdigit() and int(szcs.get(f"{{{W}}}val")) <= 4:
            micro_val = szcs.get(f"{{{W}}}val")
        if micro_val is None and szcs is not None and (szcs.get(f"{{{W}}}val") or "").isdigit() and int(szcs.get(f"{{{W}}}val")) <= 4:
            micro_val = szcs.get(f"{{{W}}}val")
        if micro_val is not None:
            _findings.append({"type": "microfont", "text": text,
                             "size_pt": int(micro_val) / 2, "path": _path})

        color = eff.find(f"{{{W}}}color")
        if color is not None:
            val = (color.get(f"{{{W}}}val") or "").lower()
            tc = (color.get(f"{{{W}}}themeColor") or "").lower()
            if val in ("ffffff", "white"):
                _findings.append({"type": "white_text_candidate", "text": text,
                                 "color": "ffffff", "path": _path})
            else:
                if _near_white(val):
                    _findings.append({"type": "near_white", "text": text,
                                     "color": val, "path": _path})
                if tc in ("light1", "light2", "background1", "background2") or \
                        (tc == "auto" and val in ("ffffff", "white", "auto")):
                    _findings.append({"type": "theme_white", "text": text,
                                     "themeColor": tc or val, "path": _path})


def _near_white(val: str) -> bool:
    if not re.fullmatch(r"[0-9a-f]{6}", val):
        return False
    try:
        r = int(val[0:2], 16)
        g = int(val[2:4], 16)
        b = int(val[4:6], 16)
    except ValueError:
        return False
    dr, dg, db = 255 - r, 255 - g, 255 - b
    if dr <= 8 and dg <= 8 and db <= 8:
        return True
    return (dr * dr + dg * dg + db * db) ** 0.5 < 16


def _filter_white_by_dark_cells(root, findings: list) -> list:
    dark_cells = [tc for tbl in root.iter(f"{{{W}}}tbl")
                  for tc in tbl.iter(f"{{{W}}}tc") if _has_dark_cell_shading(tc)]
    out = []
    for f in findings:
        if f.get("type") != "white_text_candidate":
            out.append(f)
            continue
        in_dark = any(f["text"] in _extract_text(tc) for tc in dark_cells)
        if not in_dark:
            f["type"] = "white_text_on_light_bg"
            f["note"] = "confirmed: white text on light/transparent background"
            out.append(f)
    return out


def _scan_part(zf: zipfile.ZipFile, part: str, findings: list, styles=None) -> None:
    root = ET.fromstring(zf.read(part))
    smap = _styles_map(styles)
    if not smap:
        try:
            if "word/styles.xml" in set(zf.namelist()):
                smap = _styles_map(ET.fromstring(zf.read("word/styles.xml")))
        except Exception:  # noqa: BLE001
            smap = {}
    _scan_runs(root, findings, path=part, styles=smap)
    findings[:] = _filter_white_by_dark_cells(root, findings) if False else findings
    # filter only newly added white candidates for this part without disturbing prior findings:
    # re-run filter on tail is complex; apply filter to whole findings is idempotent for non-candidates.
    filtered = _filter_white_by_dark_cells(root, findings)
    findings.clear()
    findings.extend(filtered)
    for tag in ("ins", "del"):
        for el in root.iter(f"{{{W}}}{tag}"):
            txt = _extract_text(el)
            if txt:
                findings.append({"type": "tracked_change", "text": txt, "path": part})


def _domain(url: str) -> str | None:
    if not url:
        return None
    u = url.strip().lower()
    if "://" in u:
        u = u.split("://", 1)[1]
    u = u.split("/", 1)[0].split("?", 1)[0].split("#", 1)[0]
    if ":" in u:
        u = u.split(":", 1)[0]
    return u or None


def _domain_from_text(text: str) -> str | None:
    import re
    if not text:
        return None
    m = re.search(r"(?:https?://)?(?:www\.)?([a-z0-9][a-z0-9.-]*\.[a-z]{2,})", text.lower())
    return m.group(1) if m else None


def _suspicious_metadata(text: str) -> bool:
    if not text or not text.strip():
        return False
    invisible = ("\u200b", "\u200c", "\u200d", "\u2060", "\ufeff", "\u202a", "\u202b", "\u202c", "\u202d", "\u202e")
    if any(ch in text for ch in invisible):
        return True
    low = text.lower()
    keys = ("инструкц", "ignore", "system", "prompt", "assistant", "disregard",
            "jailbreak", "<", "w:t", "http://", "https://")
    return any(k in low for k in keys)


def _check_field(instr: str, display: str, findings: list, path: str) -> None:
    if not instr:
        return
    iu = instr.upper()
    benign = ("PAGE" in iu and "HYPERLINK" not in iu) or ("NUMPAGES" in iu and "HYPERLINK" not in iu)
    if (not benign) and ("W:T" in iu or "<" in instr or "instrText" in instr):
        findings.append({"type": "field_code", "text": instr.strip(), "path": path})
    if "HYPERLINK" in iu:
        import re
        m = re.search(r'HYPERLINK\s+"([^"]+)"', instr, re.I)
        url = m.group(1) if m else None
        td = _domain(url) if url else None
        dd = _domain_from_text(display)
        if td and dd and td != dd:
            findings.append({"type": "link_mismatch", "text": display, "target": url, "path": path})


def _scan_hyperlinks_and_fields(root, findings: list, path: str, rels: dict) -> None:
    for hp in root.iter(f"{{{W}}}hyperlink"):
        rid = hp.get(f"{{{R}}}id")
        target = rels.get(rid) if rid else None
        if target is None:
            target = hp.get(f"{{{W}}}anchor")
        display = "".join(t.text or "" for t in hp.iter(f"{{{W}}}t"))
        td = _domain(target) if target and ("://" in (target or "") or "." in (target or "")) else None
        # only treat as URL domain if target looks like url
        if target and "://" not in (target or "") and not (target or "").startswith("http"):
            td = None
        dd = _domain_from_text(display)
        if td and dd and td != dd:
            findings.append({"type": "link_mismatch", "text": display, "target": target, "path": path})

    for fs in root.iter(f"{{{W}}}fldSimple"):
        instr = fs.get(f"{{{W}}}instr", "")
        display = "".join(t.text or "" for t in fs.iter(f"{{{W}}}t"))
        _check_field(instr, display, findings, path)

    for p in root.iter(f"{{{W}}}p"):
        instrs = [it.text or "" for it in p.iter(f"{{{W}}}instrText")]
        if not instrs:
            continue
        display = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
        _check_field(" ".join(instrs), display, findings, path)


def scan_all(filepath: str) -> dict:
    findings: list = []
    path = str(filepath)
    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
            if "word/document.xml" not in names:
                return {"file": path, "clean": False,
                        "findings": [{"type": "not_a_docx", "path": "word/document.xml"}]}

            smap = {}
            if "word/styles.xml" in names:
                try:
                    smap = _styles_map(ET.fromstring(zf.read("word/styles.xml")))
                except Exception:  # noqa: BLE001
                    findings.append({"type": "xml_parse_error", "path": "word/styles.xml"})

            parts = sorted(n for n in names if (n.startswith("word/") or n.startswith("docProps/")) and n.endswith(".xml"))
            for part in parts:
                try:
                    root = ET.fromstring(zf.read(part))
                except ET.ParseError:
                    findings.append({"type": "xml_parse_error", "path": part})
                    continue

                if part.startswith("docProps/"):
                    text = " ".join(root.itertext())
                    if _suspicious_metadata(text):
                        findings.append({"type": "metadata_injection", "text": text.strip(), "path": part})
                    continue

                if not list(root.iter(f"{{{W}}}t")) and part not in ("word/document.xml",):
                    # metadata-ish word parts may have no t; still scan fields
                    pass

                part_findings = []
                _scan_runs(root, part_findings, path=part, styles=smap)
                part_findings = _filter_white_by_dark_cells(root, part_findings)
                for tag in ("ins", "del"):
                    for el in root.iter(f"{{{W}}}{tag}"):
                        txt = _extract_text(el)
                        if txt:
                            part_findings.append({"type": "tracked_change", "text": txt, "path": part})

                rels = {}
                rels_name = part.replace("word/", "word/_rels/", 1) + ".rels" if part.startswith("word/") else None
                if rels_name and rels_name in names:
                    try:
                        rroot = ET.fromstring(zf.read(rels_name))
                        for rel in rroot.iter(f"{{{REL_NS}}}Relationship"):
                            rid = rel.get("Id")
                            target = rel.get("Target")
                            if rid and target:
                                rels[rid] = target
                    except ET.ParseError:
                        pass
                _scan_hyperlinks_and_fields(root, part_findings, part, rels)
                findings.extend(part_findings)
    except FileNotFoundError:
        return {"file": path, "clean": False,
                "findings": [{"type": "io_error", "path": path, "reason": "file not found"}]}
    except zipfile.BadZipFile:
        return {"file": path, "clean": False,
                "findings": [{"type": "invalid_zip", "path": path}]}
    except OSError as exc:
        return {"file": path, "clean": False,
                "findings": [{"type": "io_error", "path": path, "reason": str(exc)}]}
    except Exception as exc:  # noqa: BLE001
        return {"file": path, "clean": False,
                "findings": [{"type": "io_error", "path": path, "reason": str(exc)}]}
    return {"file": path, "clean": not findings, "findings": findings}
