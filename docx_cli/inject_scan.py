# SPDX-License-Identifier: Apache-2.0
# Copyright (c) 2026 lnfnt <lnfnt@proton.me>

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
        # S1-C1: docDefaults — базовый слой наследования для ВСЕХ run.
        dd = styles.find(f"{{{W}}}docDefaults")
        if dd is not None:
            rprd = dd.find(f"{{{W}}}rPrDefault")
            if rprd is not None:
                rpr = rprd.find(f"{{{W}}}rPr")
                if rpr is not None:
                    out["__docDefaults__"] = rpr
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
    # F-04 (аудит w2): жёсткий лимит 8 звеньев пропускал vanish на глубине 9+.
    # Ограничиваем только циклом seen — легитимные цепочки глубже 8 редки,
    # а злоумышленник лимит не уважает.
    while cur is not None and cur not in seen:
        seen.add(cur)
        chain.append(cur)
        st = styles_map.get(cur)
        if st is None:
            break
        bo = st.find(f"{{{W}}}basedOn")
        cur = bo.get(f"{{{W}}}val") if bo is not None else None
    chain.reverse()
    return chain


def _merge_rpr_layer(erpr, layer) -> None:
    if layer is None:
        return
    for child in layer:
        existing = erpr.find(child.tag)
        if existing is not None:
            erpr.remove(existing)
        erpr.append(copy.deepcopy(child))


def _effective_rpr(styles_map: dict, run_rpr, style_id, para_style_id=None) -> "ET.Element":
    erpr = ET.Element(f"{{{W}}}rPr")
    # S1-C1: слои наследования — docDefaults < стиль абзаца < стиль символа < run.
    _merge_rpr_layer(erpr, styles_map.get("__docDefaults__"))
    for sid in _style_chain(styles_map, para_style_id):
        st = styles_map.get(sid)
        if st is None:
            continue
        _merge_rpr_layer(erpr, st.find(f"{{{W}}}rPr"))
    for sid in _style_chain(styles_map, style_id):
        st = styles_map.get(sid)
        if st is None:
            continue
        _merge_rpr_layer(erpr, st.find(f"{{{W}}}rPr"))
    _merge_rpr_layer(erpr, run_rpr)
    return erpr


def _para_style_of(p, smap) -> str | None:
    ppr = p.find(f"{{{W}}}pPr")
    if ppr is None:
        return None
    ps = ppr.find(f"{{{W}}}pStyle")
    return ps.get(f"{{{W}}}val") if ps is not None else None


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

    # S1-C1: сканируем по абзацам — pStyle абзаца наследуется его run'ами.
    # Run'ы вне абзацев (sdt/txbx и т.п.) обходятся отдельно без pStyle.
    def scan_run(r, para_style_id):
        rpr = r.find(f"{{{W}}}rPr")
        style_id = None
        rs = None
        if rpr is not None:
            rs = rpr.find(f"{{{W}}}rStyle")
        if rs is None:
            rs = r.find(f"{{{W}}}rStyle")
        if rs is not None:
            style_id = rs.get(f"{{{W}}}val")
        eff = _effective_rpr(smap, rpr, style_id, para_style_id)
        text = _extract_text(r)
        if not text:
            return

        # S1-C1: vanish и specVanish — оба маркера скрытого текста.
        hidden = False
        for tag in ("vanish", "specVanish"):
            el = eff.find(f"{{{W}}}{tag}")
            if el is not None and el.get(f"{{{W}}}val", "1") not in ("0", "false", "off"):
                hidden = True
                break
        if hidden:
            _findings.append({"type": "hidden_text_vanish", "text": text, "path": _path})

        # S1-C1: значительное вертикальное смещение — текст рендерится вне строки.
        # Порог снижен до 8pt (16 half-points): сдвиг 10pt (обход прежних 12pt)
        # визуально «мелкий», но уже вне базовой линии — носитель скрытой
        # инструкции. Дробные значения парсятся float'ом (прежний int()
        # глотал ValueError на «36.0» и игнорировал находку); непарсящееся
        # значение — само по себе подозрительно, fail-closed флаг.
        pos = eff.find(f"{{{W}}}position")
        if pos is not None:
            raw_pv = pos.get(f"{{{W}}}val") or "0"
            try:
                pv = float(raw_pv)
                if abs(pv) >= 16:  # >= 8pt — не бывает у легитимного верхнего/нижнего индекса
                    _findings.append({"type": "offset_text", "text": text,
                                      "position_halfpt": pv, "path": _path})
            except ValueError:
                _findings.append({"type": "offset_text_unparseable", "text": text,
                                  "raw_position": raw_pv, "path": _path})

        sz = eff.find(f"{{{W}}}sz")
        szcs = eff.find(f"{{{W}}}szCs")
        micro_val = None
        # S1-C1: порог микрошрифта — до 5pt (10 half-points); прежние ≤4pt
        # пропускали 5pt-носителей. Дублирующая ветка szcs удалена.
        if sz is not None and (sz.get(f"{{{W}}}val") or "").isdigit() and int(sz.get(f"{{{W}}}val")) <= 10:
            micro_val = sz.get(f"{{{W}}}val")
        elif szcs is not None and (szcs.get(f"{{{W}}}val") or "").isdigit() and int(szcs.get(f"{{{W}}}val")) <= 10:
            micro_val = szcs.get(f"{{{W}}}val")
        if micro_val is not None:
            _findings.append({"type": "microfont", "text": text,
                             "size_pt": int(micro_val) / 2, "path": _path})

        color = eff.find(f"{{{W}}}color")
        val = ""
        if color is not None:
            val = (color.get(f"{{{W}}}val") or "").lower()
            tc = (color.get(f"{{{W}}}themeColor") or "").lower()
            if val in ("ffffff", "white"):
                _findings.append({"type": "white_text_candidate", "text": text,
                                 "color": "ffffff", "path": _path,
                                 "in_dark_cell": id(r) in dark_run_ids})
            else:
                if _near_white(val):
                    _findings.append({"type": "near_white", "text": text,
                                     "color": val, "path": _path})
                if tc in ("light1", "light2", "background1", "background2") or \
                        (tc == "auto" and val in ("ffffff", "white", "auto")):
                    _findings.append({"type": "theme_white", "text": text,
                                     "themeColor": tc or val, "path": _path})
            # F-03 (аудит w2): themeTint >= E0 осветляет ЛЮБОЙ базовый цвет темы
            # (включая чёрный text1/dark1) до белого — визуально невидимый текст.
            tint = (color.get(f"{{{W}}}themeTint") or "").upper()
            if tint:
                try:
                    if int(tint, 16) >= 0xE0 and val not in ("ffffff", "white"):
                        _findings.append({"type": "theme_tint_invisible", "text": text,
                                         "themeColor": tc, "themeTint": tint, "path": _path})
                except ValueError:
                    pass
        # F-05 (аудит w2): текст, окрашенный В цвет маркера (black on black,
        # yellow on yellow) — человеку «цензурированный блок», машине — пейлоад.
        hl = eff.find(f"{{{W}}}highlight")
        if hl is not None:
            hl_val = (hl.get(f"{{{W}}}val") or "").lower()
            hl_rgb = _HL_RGB.get(hl_val)
            if hl_rgb and (val or "000000") == hl_rgb:
                _findings.append({"type": "highlight_color_match", "text": text,
                                 "color": val or "auto-black", "highlight": hl_val,
                                 "path": _path})

    # S1-H3: позиционное подавление белого текста — только run'ы, реально
    # лежащие внутри тёмной ячейки. Глобальная сверка подстрок гасила находки
    # в светлых абзацах лишь по совпадению слова с тёмной ячейкой.
    dark_run_ids: set = set()
    for tbl in root.iter(f"{{{W}}}tbl"):
        for tc in tbl.iter(f"{{{W}}}tc"):
            if _has_dark_cell_shading(tc):
                for r in tc.iter(f"{{{W}}}r"):
                    dark_run_ids.add(id(r))

    in_paras: set = set()
    for p in root.iter(f"{{{W}}}p"):
        psid = _para_style_of(p, smap)
        for r in p.iter(f"{{{W}}}r"):
            in_paras.add(id(r))
            scan_run(r, psid)
    for r in root.iter(f"{{{W}}}r"):
        if id(r) not in in_paras:
            scan_run(r, None)


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


# F-05 (аудит w2): OOXML-имена маркеров highlight -> RGB (ST_HighlightColor).
# Для сравнения с цветом шрифта — совпадение = визуально «закрашенный» текст.
_HL_RGB = {
    "black": "000000", "blue": "0000ff", "cyan": "00ffff", "green": "00ff00",
    "magenta": "ff00ff", "red": "ff0000", "yellow": "ffff00", "white": "ffffff",
    "darkblue": "000080", "darkcyan": "008080", "darkgreen": "008000",
    "darkmagenta": "800080", "darkred": "800000", "darkyellow": "808000",
    "darkgray": "808080", "lightgray": "c0c0c0",
}


def _filter_white_by_dark_cells(root, findings: list) -> list:
    # S1-H3: подавление строго позиционное — только run внутри тёмной ячейки.
    # Находки без флага in_dark_cell (старые вызовы) НЕ гасим: белый текст —
    # угроза, ложное подавление хуже ложной тревоги.
    out = []
    for f in findings:
        if f.get("type") != "white_text_candidate":
            out.append(f)
            continue
        if f.get("in_dark_cell"):
            continue  # легитимный белый по тёмному — не находка
        f = dict(f)
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
    # S1-C2: инструкция поля может быть разбита по нескольким instrText
    # (« MACRO» + «BUTTON m») — прежняя склейка через пробел давала
    # « MACRO BUTTON m», и startswith("MACROBUTTON") не срабатывал.
    # Схлопываем внутренние пробельные серии перед всеми сигнатурами.
    iu = re.sub(r"\s+", "", instr).upper()
    iu_stripped = iu.strip()
    if iu_stripped.startswith("DDEAUTO") or iu_stripped.startswith("DDE"):
        findings.append({"type": "dde_auto", "text": instr.strip(), "path": path})
    # S1-H2: исполняемые и внешние поля — MACROBUTTON запускает макрос,
    # INCLUDETEXT с http(s)/file-целью тянет контент с удалённого/локального
    # ресурса. Обе разновидности дают clean=True на прежнем детекте (только
    # DDE) — флагам подлежат безусловно.
    if iu_stripped.startswith("MACROBUTTON"):
        findings.append({"type": "macrobutton_field", "text": instr.strip(), "path": path})
    # F-01 (аудит w2): путь/URL может быть БЕЗ кавычек (OOXML-допустимо при
    # отсутствии пробелов): INCLUDETEXT http://evil/x.docx — ловим обе формы.
    m_inc = re.search(r'INCLUDETEXT\s+(?:"([^"]+)"|([^\s"]+))', instr, re.IGNORECASE)
    if m_inc:
        target = (m_inc.group(1) or m_inc.group(2) or "").strip()
        tl = target.lower()
        if (tl.startswith(("http://", "https://", "file://", "file:", "\\\\"))
                or re.match(r"^[a-z]:[\\/]", tl)):
            findings.append({"type": "include_text_external", "text": instr.strip(),
                             "target": target, "path": path})
    benign = ("PAGE" in iu and "HYPERLINK" not in iu) or ("NUMPAGES" in iu and "HYPERLINK" not in iu)
    if (not benign) and ("W:T" in iu or "<" in instr or "instrText" in instr):
        findings.append({"type": "field_code", "text": instr.strip(), "path": path})
    if "HYPERLINK" in iu:
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

    seen_instr: set = set()
    for p in root.iter(f"{{{W}}}p"):
        instrs = [it.text or "" for it in p.iter(f"{{{W}}}instrText")]
        if not instrs:
            continue
        for it in p.iter(f"{{{W}}}instrText"):
            seen_instr.add(id(it))
        display = "".join(t.text or "" for t in p.iter(f"{{{W}}}t"))
        _check_field(" ".join(instrs), display, findings, path)
    # F-02 (аудит w2): instrText вне <w:p> (блочные sdtContent/customXml и пр.)
    # прежним циклом по абзацам не виден — DDEAUTO/INCLUDETEXT там обходили
    # сканер. Обрабатываем глобально каждый неучтённый instrText.
    for it in root.iter(f"{{{W}}}instrText"):
        if id(it) in seen_instr:
            continue
        txt = it.text or ""
        if txt.strip():
            _check_field(txt, "", findings, path)


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
