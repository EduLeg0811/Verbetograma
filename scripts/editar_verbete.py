"""Gera copias DOCX marcadas ou corrigidas a partir dos achados formais."""

from __future__ import annotations

import copy
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from .analisar_verbete import NS, clean
except ImportError:
    try:
        from scripts.analisar_verbete import NS, clean
    except ImportError:
        from analisar_verbete import NS, clean

# Registry of standard namespaces to prevent corrupt element prefixes in DOCX (like ns0:, ns1:)
NAMESPACES_TO_REGISTER = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "ve": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "o": "urn:schemas-microsoft-com:office:office",
    "v": "urn:schemas-microsoft-com:vml",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/15/chartex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
}

for prefix, uri in NAMESPACES_TO_REGISTER.items():
    ET.register_namespace(prefix, uri)

# Canonical ordering of w:rPr elements inside a text run (strict order required by OpenXML schema)
RPR_ORDER = [
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
    "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u",
    "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath", "rPrChange"
]
RPR_ORDER_MAP = {name: idx for idx, name in enumerate(RPR_ORDER)}


def _local_name(tag: str) -> str:
    if tag.startswith("{"):
        return tag.split("}", 1)[1]
    return tag.split(":", 1)[-1]


def _sort_rpr_children(rpr: ET.Element) -> None:
    children = list(rpr)
    children.sort(key=lambda child: RPR_ORDER_MAP.get(_local_name(child.tag), 999))
    for child in list(rpr):
        rpr.remove(child)
    for child in children:
        rpr.append(child)


def _qn(name: str) -> str:
    prefix, tag = name.split(":")
    return f"{{{NS[prefix]}}}{tag}"


def _paragraph_text(p: ET.Element) -> str:
    return clean("".join(t.text or "" for t in p.findall(".//w:t", NS)))


def _run_text(run: ET.Element) -> str:
    return "".join(t.text or "" for t in run.findall(".//w:t", NS))


def _set_run_text(run: ET.Element, text: str) -> None:
    texts = run.findall(".//w:t", NS)
    if not texts:
        t = ET.SubElement(run, _qn("w:t"))
    else:
        t = texts[0]
        for extra in texts[1:]:
            parent = run
            parent.remove(extra)
    t.text = text
    if text.startswith(" ") or text.endswith(" "):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _ensure_rpr(run: ET.Element) -> ET.Element:
    rpr = run.find("w:rPr", NS)
    if rpr is None:
        rpr = ET.Element(_qn("w:rPr"))
        run.insert(0, rpr)
    return rpr


def _set_bold(run: ET.Element) -> None:
    rpr = _ensure_rpr(run)
    if rpr.find("w:b", NS) is None:
        rpr.append(ET.Element(_qn("w:b")))
    _sort_rpr_children(rpr)


def _set_highlight(run: ET.Element) -> None:
    rpr = _ensure_rpr(run)
    highlight = rpr.find("w:highlight", NS)
    if highlight is None:
        highlight = ET.Element(_qn("w:highlight"))
        rpr.append(highlight)
    highlight.set(_qn("w:val"), "yellow")
    _sort_rpr_children(rpr)


def _split_run_for_char(p: ET.Element, char_idx: int) -> ET.Element | None:
    cursor = 0
    for run in list(p.findall("./w:r", NS)):
        texts = run.findall("./w:t", NS)
        if not texts:
            text = ""
        else:
            text = "".join(t.text or "" for t in texts)
            
        end = cursor + len(text)
        if cursor <= char_idx < end:
            local = char_idx - cursor
            before, char, after = text[:local], text[local : local + 1], text[local + 1 :]
            pos = list(p).index(run)
            
            def make_run(txt: str, keep_extra: bool) -> ET.Element:
                new_run = copy.deepcopy(run)
                texts = new_run.findall("./w:t", NS)
                if texts:
                    t = texts[0]
                    t.text = txt
                    if txt.startswith(" ") or txt.endswith(" "):
                        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    else:
                        t.attrib.pop("{http://www.w3.org/XML/1998/namespace}space", None)
                    for extra_t in texts[1:]:
                        new_run.remove(extra_t)
                else:
                    t = ET.SubElement(new_run, _qn("w:t"))
                    t.text = txt
                    
                if not keep_extra:
                    for child in list(new_run):
                        if child.tag not in {_qn("w:rPr"), _qn("w:t")}:
                            new_run.remove(child)
                return new_run

            new_runs = []
            if before:
                new_runs.append(make_run(before, keep_extra=True))
                char_run = make_run(char, keep_extra=False)
                new_runs.append(char_run)
            else:
                char_run = make_run(char, keep_extra=True)
                new_runs.append(char_run)
                
            if after:
                new_runs.append(make_run(after, keep_extra=False))
                
            p.remove(run)
            for offset, new_run in enumerate(new_runs):
                p.insert(pos + offset, new_run)
            return char_run
        cursor = end
    return None


def _last_period_index(text: str) -> int | None:
    for idx in range(len(text) - 1, -1, -1):
        if text[idx].isspace():
            continue
        return idx if text[idx] == "." else None
    return None


def _target_index(rule: str, text: str) -> int | None:
    if "Dois-pontos" in rule:
        idx = text.rfind(":")
        return idx if idx >= 0 else None
    if "Ponto final" in rule:
        return _last_period_index(text)
    if "Epigrafe" in rule:
        match_dot = text.find(".")
        match_colon = text.find(":")
        candidates = [idx for idx in [match_dot, match_colon] if idx >= 0]
        return min(candidates) if candidates else None
    return None


def _find_paragraph(paragraphs: list[ET.Element], evidence: str) -> ET.Element | None:
    target = clean(evidence)
    if not target:
        return None
    target_prefix = target[:180]
    for p in paragraphs:
        text = _paragraph_text(p)
        if text == target or text.startswith(target_prefix) or target.startswith(text[:180]):
            return p
    return None


def _write_docx_from_tree(source: Path, document_xml: bytes) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        out_path = Path(tmp.name)
    try:
        with zipfile.ZipFile(source, "r") as zin, zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == "word/document.xml":
                    zout.writestr(item, document_xml)
                else:
                    zout.writestr(item, zin.read(item.filename))
        data = out_path.read_bytes()
    finally:
        out_path.unlink(missing_ok=True)
    return data


def _edit_docx(source: Path, result: dict, mode: str) -> bytes:
    source = Path(source)
    with zipfile.ZipFile(source) as zf:
        document_xml = zf.read("word/document.xml")

    # Dynamically extract and register every namespace defined in the document.xml
    try:
        from io import BytesIO
        context = ET.iterparse(BytesIO(document_xml), events=("start-ns",))
        for event, elem in context:
            prefix, uri = elem
            ET.register_namespace(prefix, uri)
    except Exception:
        pass

    root = ET.fromstring(document_xml)
    paragraphs = [p for p in root.findall(".//w:p", NS) if _paragraph_text(p)]

    for item in result.get("formatacao", {}).get("ajustes", []):
        paragraph = _find_paragraph(paragraphs, item.get("evidencia", ""))
        if paragraph is None:
            continue
        text = "".join(_run_text(run) for run in paragraph.findall("./w:r", NS))
        idx = _target_index(item.get("regra", ""), text)
        if idx is None:
            continue
        run = _split_run_for_char(paragraph, idx)
        if run is None:
            continue
        if mode == "mark":
            _set_highlight(run)
        elif mode == "correct":
            if "Dois-pontos" in item.get("regra", "") or "Ponto final" in item.get("regra", "") or "Epigrafe" in item.get("regra", ""):
                _set_bold(run)

    # Use standard double-quoted xml declaration instead of default python single-quotes version
    xml_bytes = ET.tostring(root, encoding="utf-8")
    
    # Restore the original <w:document ...> start tag to preserve all namespace declarations (like w15, w16se)
    # listed in mc:Ignorable that ElementTree discards because they are not directly used in tags.
    import re
    orig_doc_start = re.search(rb"<\w+:document\s+[^>]*>", document_xml)
    new_doc_start = re.search(rb"<\w+:document\s+[^>]*>", xml_bytes)
    if orig_doc_start and new_doc_start:
        xml_bytes = xml_bytes[:new_doc_start.start()] + orig_doc_start.group(0) + xml_bytes[new_doc_start.end():]

    xml_declaration = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
    return _write_docx_from_tree(source, xml_declaration + xml_bytes)


def gerar_verbete_marcado(source: str | Path, result: dict) -> bytes:
    """Retorna DOCX com highlight amarelo nos pontos de correcao automatizados."""
    return _edit_docx(Path(source), result, "mark")


def gerar_verbete_corrigido(source: str | Path, result: dict) -> bytes:
    """Retorna DOCX com correcoes automaticas de negrito/pontuacao aplicadas."""
    return _edit_docx(Path(source), result, "correct")


def gerar_pdf_marcado(source: str | Path, result: dict) -> bytes:
    """Retorna PDF com highlight amarelo nos trechos de evidencia encontrados."""
    import fitz

    doc = fitz.open(str(source))
    evidencias = [item.get("evidencia", "") for item in result.get("formatacao", {}).get("ajustes", [])]
    if not evidencias:
        evidencias = [item.get("evidencia", "") for item in result.get("achados", [])]

    for evidence in evidencias:
        evidence = clean(evidence)
        if not evidence:
            continue
        candidates = [
            evidence,
            evidence[:240],
            evidence[:160],
            evidence.split(". ")[0],
        ]
        for page in doc:
            found = False
            for candidate in candidates:
                candidate = candidate.strip()
                if len(candidate) < 12:
                    continue
                rects = page.search_for(candidate)
                if not rects:
                    continue
                for rect in rects:
                    annot = page.add_highlight_annot(rect)
                    annot.set_colors(stroke=(1, 0.86, 0.12))
                    annot.update()
                found = True
                break
            if found:
                break

    data = doc.tobytes(garbage=4, deflate=True)
    doc.close()
    return data
