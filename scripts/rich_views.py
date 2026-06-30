from __future__ import annotations

from copy import deepcopy
from html import escape
from pathlib import Path
import re

from scripts.analisar_verbete import (
    clean,
    clean_sec,
    detect_division,
    detect_independent_epigraph,
    get_effective_section,
    logia_terms,
    read_docx_rich,
    section_body,
    starts_with_argumentologia_epigraph,
)
from scripts.catalogo import SECOES


def rich_runs_to_html(runs: list[dict]) -> str:
    parts = []
    for run in runs:
        text = escape(str(run.get("text", "")))
        if not text:
            continue
        if run.get("highlight"):
            text = f'<mark class="docx-highlight">{text}</mark>'
        if run.get("bold"):
            text = f"<strong>{text}</strong>"
        if run.get("italic"):
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts)


def find_rich_paragraph(paragraphs: list[dict], evidence: str) -> dict | None:
    target = clean(evidence)
    if not target:
        return None
    prefix = target[:180]
    for paragraph in paragraphs:
        text = clean(paragraph["text"])
        if text == target or text.startswith(prefix) or target.startswith(text[:180]):
            return paragraph
    return next((p for p in paragraphs if target in clean(p["text"])), None)


def highlight_runs_span(paragraph: dict, start: int, end: int) -> None:
    if start < 0 or start >= end:
        return
    runs, cursor = [], 0
    for run in paragraph["runs"]:
        text = run["text"]
        run_end = cursor + len(text)
        overlap_start, overlap_end = max(cursor, start), min(run_end, end)
        if overlap_start < overlap_end:
            local_start, local_end = overlap_start - cursor, overlap_end - cursor
            if local_start:
                before = deepcopy(run); before["text"] = text[:local_start]; runs.append(before)
            marked = deepcopy(run); marked["text"] = text[local_start:local_end]; marked["highlight"] = True; runs.append(marked)
            if local_end < len(text):
                after = deepcopy(run); after["text"] = text[local_end:]; runs.append(after)
        else:
            runs.append(run)
        cursor = run_end
    paragraph["runs"] = runs


def target_spans(rule: str, text: str, suggestion: str, evidence: str) -> list[tuple[int, int]]:
    rule = rule.lower()
    spans: list[tuple[int, int]] = []
    if "ponto final" in rule:
        index = text.rfind("."); spans = [(index, index + 1)] if index >= 0 else []
    elif "dois-pontos" in rule or "dois pontos" in rule:
        index = text.rfind(":"); spans = [(index, index + 1)] if index >= 0 else []
    elif "epígrafe" in rule or "epigrafe" in rule:
        candidates = [i for i in (text.find("."), text.find(":")) if i >= 0]
        spans = [(0, min(candidates) + 1)] if candidates else []
    elif "expressão composta" in rule or "expressao composta" in rule:
        index = text.find(evidence); spans = [(index, index + len(evidence))] if index >= 0 else []
    elif "artigo sem itálico" in rule:
        spans = [(m.start(), m.end()) for m in re.finditer(r"\b(o|a|os|as)\b", text)]
    elif "separador sem itálico" in rule:
        spans = [(i, i + 1) for i, char in enumerate(text) if char == ";"]
    elif "título em negrito" in rule or "capitalização do título" in rule:
        match = re.match(r"^\s*(\d{1,3})\.?\s+(.+)$", text)
        if match:
            start, offset = match.end(), text[match.end():].find(":")
            spans = [(start, start + offset)] if offset >= 0 else []
    elif "capitalização dos itens secundários" in rule or "estilo dos itens secundários" in rule:
        match = re.search(r"O item(?: secundário)? '([^']+)'", suggestion)
        if match:
            index = text.find(match.group(1)); spans = [(index, index + len(match.group(1)))] if index >= 0 else []
    elif "espaçamento duplo" in rule:
        spans = [(m.start(), m.end()) for m in re.finditer(r"(?<=\S) (?=\S)", text)]
    elif "dois espaços após número" in rule:
        match = re.search(r"(\d{1,2}\.) (\S)", text)
        spans = [(match.start(2) - 1, match.start(2))] if match else []
    elif "padrão zero" in rule:
        match = re.match(r"^\s*\d{1,3}\.?\s*", text); spans = [(match.start(), match.end())] if match else []
    elif "sufixo repetido" in rule:
        match = re.search(r"sufixo '([^']+)'", suggestion)
        if match:
            spans = [(m.start(), m.end()) for m in re.finditer(r"\b\w*" + re.escape(match.group(1)) + r"\b", text)]
    elif "ortopensatologia" in rule or "fórmula da ortopensatologia" in rule:
        spans = [(i, i + 1) for i, char in enumerate(text) if char in "“”\""]
    elif "interlocução" in rule:
        spans = [(m.start(), m.end()) for m in re.finditer(r"leitor ou leitora", text, re.I)]
    elif "parasitas" in rule:
        for word in (item.strip() for item in evidence.split(",")):
            if word:
                spans.extend((m.start(), m.end()) for m in re.finditer(r"\b" + re.escape(word) + r"\b", text, re.I))
    return spans


def build_rich_docx_views(path: Path, result: dict) -> dict[str, str]:
    try:
        paragraphs = read_docx_rich(path)
    except Exception:
        return {}

    for item in result.get("formatacao", {}).get("ajustes", []) + result.get("achados", []):
        paragraph = find_rich_paragraph(paragraphs, item.get("evidencia", ""))
        if paragraph:
            raw = paragraph.get("raw_text", paragraph["text"])
            for start, end in target_spans(item.get("regra", ""), raw, item.get("sugestao", ""), item.get("evidencia", "")):
                highlight_runs_span(paragraph, start, end)

    current_div = current_sec = None
    sections: dict[str, list[str]] = {}
    emphatic_found = False
    for index, paragraph in enumerate(paragraphs):
        text, html = paragraph["text"], rich_runs_to_html(paragraph["runs"])
        division = detect_division(text)
        if division:
            current_div, current_sec = division, None
            continue
        if current_div == "Argumentologia":
            if starts_with_argumentologia_epigraph(text) or not current_sec:
                number = len([key for key in sections if key.startswith("Argumentologia_")])
                current_sec = f"Argumentologia_{number}"
            sections.setdefault(current_sec, []).append(html)
            continue
        section = detect_independent_epigraph(text)
        if section:
            current_sec = section
            sections.setdefault(section, []).append(html)
            continue
        next_section = detect_independent_epigraph(paragraphs[index + 1]["text"]) if index + 1 < len(paragraphs) else None
        if current_sec == "Remissiologia" and next_section == "Questionologia" and len(text) > 40 and not emphatic_found:
            current_sec, emphatic_found = "Frase Enfatica", True
        if current_sec:
            sections.setdefault(current_sec, []).append(html)
    return {section: "<br><br>".join(parts) for section, parts in sections.items()}


def plain_section_html(text: str, findings: list[dict]) -> str:
    spans = []
    for item in findings:
        spans.extend(target_spans(item["regra"], text, item["sugestao"], item["evidencia"]))
    merged = []
    for start, end in sorted(set(spans)):
        if merged and start <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    parts, cursor = [], 0
    for start, end in merged:
        parts.extend((escape(text[cursor:start]), f'<mark class="docx-highlight">{escape(text[start:end])}</mark>'))
        cursor = end
    parts.append(escape(text[cursor:]))
    return "".join(parts)


def build_section_summary(result: dict, rich_sections: dict[str, str]) -> list[dict]:
    rows = []
    for section in result.get("secoes_presentes", []):
        text = result.get("secoes_texto", {}).get(section, "")
        findings = [item for item in result.get("achados", []) if item["secao"] == section]
        effective = get_effective_section(section, text)
        maximum = SECOES.get(effective, {}).get("max")
        reached = any(item["secao"] == section for item in result.get("maximos", []))
        terms = []
        if effective not in {"Interdisciplinologia", "Frase Enfatica"} and re.search(r"logia$", effective, re.I):
            terms.append(effective)
        if effective not in {"Interdisciplinologia", "Remissiologia", "Frase Enfatica"}:
            terms.extend(logia_terms(section_body(text)))
        rows.append({
            "section": clean_sec(section),
            "key": section,
            "html": rich_sections.get(section) or plain_section_html(text, findings),
            "count": result.get("contagens", {}).get(section),
            "maximum": f">= {maximum}" if maximum and reached else f"< {maximum}" if maximum else "",
            "logias": len(terms),
            "findings": findings,
        })
    return rows
