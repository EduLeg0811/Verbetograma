"""Gera copias DOCX marcadas ou corrigidas a partir dos achados formais."""

from __future__ import annotations

import copy
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from .analisar_verbete import NS, clean

ET.register_namespace("w", NS["w"])


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


def _set_highlight(run: ET.Element) -> None:
    rpr = _ensure_rpr(run)
    highlight = rpr.find("w:highlight", NS)
    if highlight is None:
        highlight = ET.Element(_qn("w:highlight"))
        rpr.append(highlight)
    highlight.set(_qn("w:val"), "yellow")


def _split_run_for_char(p: ET.Element, char_idx: int) -> ET.Element | None:
    cursor = 0
    for run in list(p.findall("./w:r", NS)):
        text = _run_text(run)
        end = cursor + len(text)
        if cursor <= char_idx < end:
            local = char_idx - cursor
            before, char, after = text[:local], text[local : local + 1], text[local + 1 :]
            pos = list(p).index(run)
            new_runs = []
            if before:
                before_run = copy.deepcopy(run)
                _set_run_text(before_run, before)
                new_runs.append(before_run)
            char_run = copy.deepcopy(run)
            _set_run_text(char_run, char)
            new_runs.append(char_run)
            if after:
                after_run = copy.deepcopy(run)
                _set_run_text(after_run, after)
                new_runs.append(after_run)
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

    return _write_docx_from_tree(source, ET.tostring(root, encoding="utf-8", xml_declaration=True))


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
