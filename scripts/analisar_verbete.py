"""Motor formal leve para analisar verbetes e gerar achados em JSON."""

from __future__ import annotations

import argparse
import json
import re
import shutil
import subprocess
import sys
import tempfile
import unicodedata
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

try:
    from .catalogo import DIVISAO_ALIASES, DIVISOES, MARCA_EXCELENCIA, PARASITAS, SECOES, SECOES_FIXAS
except ImportError:
    from catalogo import DIVISAO_ALIASES, DIVISOES, MARCA_EXCELENCIA, PARASITAS, SECOES, SECOES_FIXAS

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def _strip_accents(text: str) -> str:
    return "".join(ch for ch in unicodedata.normalize("NFD", text) if unicodedata.category(ch) != "Mn")


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", _strip_accents(text).lower()).strip()


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text.replace("\xa0", " ")).strip()


def read_docx(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    paragraphs = []
    for p in root.findall(".//w:p", NS):
        parts = [t.text or "" for t in p.findall(".//w:t", NS)]
        text = clean("".join(parts))
        if text:
            paragraphs.append(text)
    return paragraphs


def read_docx_rich(path: Path) -> list[dict]:
    with zipfile.ZipFile(path) as zf:
        xml = zf.read("word/document.xml")
    root = ET.fromstring(xml)
    paragraphs = []
    for p in root.findall(".//w:p", NS):
        ppr = p.find("w:pPr", NS)
        align = None
        if ppr is not None:
            jc = ppr.find("w:jc", NS)
            align = jc.get(f"{{{NS['w']}}}val") if jc is not None else None
            spacing = ppr.find("w:spacing", NS)
            line_spacing = spacing.get(f"{{{NS['w']}}}line") if spacing is not None else None
            borders = ppr.find("w:pBdr", NS)
            border_vals = []
            if borders is not None:
                for border in list(borders):
                    val = border.get(f"{{{NS['w']}}}val")
                    if val:
                        border_vals.append(val)
        else:
            line_spacing = None
            border_vals = []
        runs = []
        for r in p.findall("./w:r", NS):
            text = "".join(t.text or "" for t in r.findall(".//w:t", NS)).replace("\xa0", " ")
            if not text:
                continue
            rpr = r.find("w:rPr", NS)
            bold = rpr is not None and rpr.find("w:b", NS) is not None
            italic = rpr is not None and rpr.find("w:i", NS) is not None
            small_caps = rpr is not None and rpr.find("w:smallCaps", NS) is not None
            underline = rpr is not None and rpr.find("w:u", NS) is not None
            font = ""
            size = None
            if rpr is not None:
                rfonts = rpr.find("w:rFonts", NS)
                if rfonts is not None:
                    font = rfonts.get(f"{{{NS['w']}}}ascii") or rfonts.get(f"{{{NS['w']}}}hAnsi") or ""
                sz = rpr.find("w:sz", NS)
                if sz is not None and sz.get(f"{{{NS['w']}}}val"):
                    try:
                        size = int(sz.get(f"{{{NS['w']}}}val")) / 2
                    except (TypeError, ValueError):
                        size = None
                spacing = rpr.find("w:spacing", NS)
                char_spacing = None
                if spacing is not None and spacing.get(f"{{{NS['w']}}}val"):
                    try:
                        char_spacing = int(spacing.get(f"{{{NS['w']}}}val")) / 20
                    except (TypeError, ValueError):
                        char_spacing = None
                else:
                    char_spacing = None
            else:
                char_spacing = None
            runs.append(
                {
                    "text": text,
                    "bold": bold,
                    "italic": italic,
                    "small_caps": small_caps,
                    "underline": underline,
                    "font": font,
                    "size": size,
                    "char_spacing": char_spacing,
                }
            )
        text = "".join(run["text"] for run in runs).strip()
        if text:
            paragraphs.append({"text": clean(text), "raw_text": text, "runs": runs, "align": align, "line_spacing": line_spacing, "borders": border_vals})
    return paragraphs


def read_docx_page_count(path: Path) -> int | None:
    try:
        with zipfile.ZipFile(path) as zf:
            xml = zf.read("docProps/app.xml")
        root = ET.fromstring(xml)
    except Exception:
        return None
    for elem in root.iter():
        if elem.tag.endswith("Pages") and elem.text and elem.text.isdigit():
            return int(elem.text)
    return None


def read_docx_layout(path: Path) -> dict:
    data: dict = {"sections": [], "headers": []}
    try:
        with zipfile.ZipFile(path) as zf:
            document_xml = zf.read("word/document.xml")
            root = ET.fromstring(document_xml)
            for sect in root.findall(".//w:sectPr", NS):
                pg_sz = sect.find("w:pgSz", NS)
                pg_mar = sect.find("w:pgMar", NS)
                section = {}
                if pg_sz is not None:
                    section["w"] = int(pg_sz.get(f"{{{NS['w']}}}w", "0") or 0)
                    section["h"] = int(pg_sz.get(f"{{{NS['w']}}}h", "0") or 0)
                if pg_mar is not None:
                    section["margins"] = {
                        key: int(pg_mar.get(f"{{{NS['w']}}}{key}", "0") or 0)
                        for key in ["top", "bottom", "left", "right", "gutter"]
                    }
                if section:
                    data["sections"].append(section)
            for name in zf.namelist():
                if not re.fullmatch(r"word/header\d+\.xml", name):
                    continue
                header_root = ET.fromstring(zf.read(name))
                runs = []
                text_parts = []
                for r in header_root.findall(".//w:r", NS):
                    text = "".join(t.text or "" for t in r.findall(".//w:t", NS))
                    if text:
                        text_parts.append(text)
                    rpr = r.find("w:rPr", NS)
                    font = ""
                    size = None
                    italic = False
                    if rpr is not None:
                        italic = rpr.find("w:i", NS) is not None
                        rfonts = rpr.find("w:rFonts", NS)
                        if rfonts is not None:
                            font = rfonts.get(f"{{{NS['w']}}}ascii") or rfonts.get(f"{{{NS['w']}}}hAnsi") or ""
                        sz = rpr.find("w:sz", NS)
                        if sz is not None and sz.get(f"{{{NS['w']}}}val"):
                            try:
                                size = int(sz.get(f"{{{NS['w']}}}val")) / 2
                            except (TypeError, ValueError):
                                size = None
                    if text:
                        runs.append({"text": text, "font": font, "size": size, "italic": italic})
                fields = [instr.text or "" for instr in header_root.findall(".//w:instrText", NS)]
                data["headers"].append({"text": clean("".join(text_parts)), "runs": runs, "fields": fields})
    except Exception:
        return data
    return data


def read_pdf(path: Path) -> list[str]:
    try:
        from pypdf import PdfReader
    except Exception as exc:
        raise RuntimeError("Instale pypdf para analisar PDF: pip install pypdf") from exc
    reader = PdfReader(str(path))
    lines: list[str] = []
    for page in reader.pages:
        text = page.extract_text() or ""
        lines.extend(clean(line) for line in text.splitlines() if clean(line))
    return lines


def read_plain(path: Path) -> list[str]:
    return [clean(line) for line in path.read_text(encoding="utf-8", errors="ignore").splitlines() if clean(line)]


def convert_doc_to_docx(path: Path) -> Path | None:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if soffice:
        out_dir = Path(tempfile.mkdtemp(prefix="verb_doc_"))
        subprocess.run(
            [soffice, "--headless", "--convert-to", "docx", "--outdir", str(out_dir), str(path)],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        converted = out_dir / f"{path.stem}.docx"
        if converted.exists():
            return converted
    try:
        import win32com.client  # type: ignore

        out_dir = Path(tempfile.mkdtemp(prefix="verb_doc_"))
        converted = out_dir / f"{path.stem}.docx"
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(path.resolve()))
        doc.SaveAs(str(converted.resolve()), FileFormat=16)
        doc.Close(False)
        word.Quit()
        if converted.exists():
            return converted
    except Exception:
        return None
    return None


def read_document(path: Path) -> tuple[list[str], list[str]]:
    warnings = []
    suffix = path.suffix.lower()
    if suffix == ".docx":
        return read_docx(path), warnings
    if suffix == ".pdf":
        warnings.append("PDF analisado em modo melhor esforco; confirme achados de formatacao no DOC/DOCX.")
        return read_pdf(path), warnings
    if suffix in {".txt", ".md"}:
        return read_plain(path), warnings
    if suffix == ".doc":
        converted = convert_doc_to_docx(path)
        if converted:
            warnings.append("DOC convertido automaticamente para DOCX antes da analise.")
            return read_docx(converted), warnings
        warnings.append("DOC binario nao pode ser convertido neste ambiente. Instale LibreOffice ou use Microsoft Word/COM, ou envie DOCX.")
        return [], warnings
    raise ValueError(f"Formato nao suportado: {suffix}")


def detect_division(text: str) -> str | None:
    m = re.match(r"^(?:[IVX]+\.\s*)?(.+?)\s*$", text, re.I)
    if not m:
        return None
    key = norm(m.group(1)).replace(".", "")
    return DIVISAO_ALIASES.get(key)


def detect_section(text: str) -> str | None:
    head = re.split(r"[:.]", text, maxsplit=1)[0].strip()
    head_norm = norm(head)
    for sec in SECOES:
        if head_norm == norm(sec):
            return sec
    if head_norm in {"frase enfatica", "frase enfática"}:
        return "Frase Enfatica"
    return None


def detect_independent_epigraph(text: str) -> str | None:
    known = detect_section(text)
    if known:
        return known
    match = re.match(r"^([A-ZÁÉÍÓÚÂÊÔÃÕÇ][A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç -]{3,60}?)(?:[:.])(?:\s+|$)", text)
    if not match:
        return None
    head = clean(match.group(1))
    if re.search(r"logia$", head, re.I):
        return head
    return None


def extract_meta(paragraphs: list[str], sections: dict[str, str]) -> dict:
    title = ""
    specialty = ""
    for p in paragraphs[:12]:
        if not specialty and re.fullmatch(r"\(.+?logia\)", p, re.I):
            specialty = p.strip("() ")
            break
        if not title and not detect_division(p) and not detect_section(p):
            title = p
    tematologia = sections.get("Tematologia", "")
    tema = ""
    m = re.search(r"\b(homeost[aá]tico|neutro|nosogr[aá]fico)\b", tematologia, re.I)
    if m:
        tema = m.group(1).lower()
    return {"titulo": title, "especialidade": specialty, "tema_central": tema, "paginas": None}


def section_body(text: str) -> str:
    lines = text.splitlines()
    if lines and detect_section(lines[0]) and len(lines) > 1:
        return "\n".join(lines[1:]).strip()
    return re.sub(r"^[^:.]{2,80}[:.]\s*", "", text).strip()


def count_items(section: str, text: str) -> int:
    body = section_body(text)
    rule = SECOES.get(section, {})
    enum = rule.get("enum")
    if not body:
        return 0
    if section == "Enumerologia":
        numbered = re.findall(r"(?:^|\s)\d{1,3}\.\s+", body)
        if numbered:
            return len(numbered)
        parts = [p for p in re.split(r";|\n", body) if clean(p.strip(" .;"))]
        return len(parts) if parts else 1
    if section == "Etimologia":
        return 1
    declared = re.search(r"\b(?:as|os)\s+(\d{1,3})\s+express", body, re.I)
    if section == "Neologia" and declared:
        return int(declared.group(1))
    if enum == "horizontal_numerada":
        first_line = body.splitlines()[0] if body.splitlines() else body
        return len(re.findall(r"(?:^|\s)\d{1,3}\.\s+", first_line))
    if enum == "vertical_numerada":
        return len(re.findall(r"(?m)(?:^|\n)\s*\d{1,3}\.\s+", body))
    if enum == "linhas":
        return max(1, len([p for p in re.split(r"[.;]\s+", body) if p.strip()]))
    if ";" in body:
        return len([p for p in body.split(";") if p.strip()])
    if "," in body and section in {"Definologia", "Etimologia", "Cognatologia", "Neologia"}:
        return len([p for p in body.split(",") if p.strip()])
    return 1


def numbered_tokens(text: str) -> list[str]:
    return re.findall(r"(?:^|\s)(\d{1,3})\.\s+", text)


def vertical_numbered_tokens(text: str) -> list[str]:
    return re.findall(r"(?m)(?:^|\n)\s*(\d{1,3})\.\s+", text)


def check_zero_pattern(section: str, text: str) -> tuple[str, str] | None:
    rule = SECOES.get(section, {})
    if rule.get("enum") not in {"horizontal_numerada", "vertical_numerada"}:
        return None
    body = section_body(text)
    nums = vertical_numbered_tokens(body) if rule.get("enum") == "vertical_numerada" else numbered_tokens(body)
    if not nums:
        return None
    total = len(nums)
    first_nine = [num for num in nums if int(num) <= 9]
    if total >= 10 and any(len(num) == 1 for num in first_nine):
        return "Padrão zero ausente", "Em enumeração numerada com 10 ou mais itens, usar 01 a 09 nos nove primeiros itens."
    if total <= 9 and any(len(num) > 1 and num.startswith("0") for num in first_nine):
        return "Padrão zero indevido", "Em enumeração numerada com até 9 itens, usar 1 a 9 sem zero à esquerda."
    return None


def item_texts(section: str, text: str) -> list[str]:
    body = section_body(text)
    if not body:
        return []
    enum = SECOES.get(section, {}).get("enum")
    if enum in {"horizontal_numerada", "vertical_numerada"}:
        pattern = r"(?m)(?:^|\n)\s*(\d{1,3})\.\s+" if enum == "vertical_numerada" else r"(?:^|\s)(\d{1,3})\.\s+"
        source = body if enum == "vertical_numerada" else (body.splitlines()[0] if body.splitlines() else body)
        matches = list(re.finditer(pattern, source))
        items = []
        for idx, match in enumerate(matches):
            start = match.end()
            end = matches[idx + 1].start() if idx + 1 < len(matches) else len(source)
            item = clean(source[start:end].strip(" ;."))
            if item:
                items.append(item)
        return items
    if ";" in body:
        return [clean(part.strip(" .")) for part in body.split(";") if clean(part.strip(" ."))]
    if "," in body and section in {"Cognatologia", "Neologia"}:
        return [clean(part.strip(" .")) for part in body.split(",") if clean(part.strip(" ."))]
    return [clean(body.strip(" ."))] if clean(body.strip(" .")) else []


def alphabetic_key(text: str) -> str:
    text = re.sub(r"^\d{1,3}\.\s*", "", text)
    text = re.sub(r"^(?:a|as|o|os|um|uma|uns|umas)\s+", "", text, flags=re.I)
    text = re.sub(r"^[\"'“”‘’*_\s]+", "", text)
    words = re.findall(r"[\wÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç-]+", text)
    return norm(" ".join(words))


def alphabetical_issue(section: str, text: str) -> tuple[str, str] | None:
    if section not in {
        "Cognatologia",
        "Sinonimologia",
        "Antonimologia",
        "Taxologia",
        "Caracterologia",
        "Tipologia",
        "Tabelologia",
        "Remissiologia",
    }:
        return None
    items = item_texts(section, text)
    if len(items) < 3:
        return None
    keys = [alphabetic_key(item) for item in items]
    keys = [key for key in keys if key]
    if len(keys) >= 3 and keys != sorted(keys):
        return "Ordem alfabetica", "Reordenar os itens em ordem alfabetica."
    return None


def numbered_spacing_issue(section: str, text: str) -> tuple[str, str] | None:
    if section not in {"Sinonimologia", "Antonimologia"}:
        return None
    body = section_body(text)
    bad = re.search(r"(?:^|\s)\d{1,2}\.(?!  )\S?", body)
    if bad:
        return "Dois espacos apos numero", "Usar dois espacos entre o numero da acepcao e o texto em Sinonimologia/Antonimologia."
    return None


def sinon_anton_separator_issue(section: str, text: str) -> tuple[str, str] | None:
    if section not in {"Sinonimologia", "Antonimologia"}:
        return None
    body = section_body(text)
    if re.search(r"\d{1,2}\.\s+[^.]*;\s+\d{1,2}\.", body):
        return "Separacao de acepcoes", "Separar acepcoes principais por ponto; reservar ponto e virgula para subitens internos."
    return None


def quote_issue(section: str, text: str) -> tuple[str, str] | None:
    if section not in {"Etimologia", "Citaciologia", "Ortopensatologia"}:
        return None
    pairs = [("“", "”"), ('"', '"'), ("‘", "’")]
    for open_q, close_q in pairs:
        if open_q == close_q:
            if text.count(open_q) % 2:
                return "Aspas abre-fecha", "Conferir fechamento das aspas nesta secao."
        elif text.count(open_q) != text.count(close_q):
            return "Aspas abre-fecha", "Conferir abertura e fechamento das aspas nesta secao."
    return None


def megapensene_issue(text: str) -> tuple[str, str] | None:
    body = section_body(text)
    if not body:
        return None

    intro = re.match(r"^(.{0,160}?\bmegapensene\w*\b.{0,160}?:)\s*(.+)$", body, re.I | re.S)
    if intro:
        body = intro.group(2)
    body = re.sub(r"^[\s\-–—:]+", "", body).strip()

    sentences = [clean(part.strip(" .;:–—-")) for part in re.split(r"\s*[.;]\s*", body) if clean(part.strip(" .;:–—-"))]
    bad = []
    for sentence in sentences:
        words = re.findall(r"[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç]+(?:-[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç]+)?", sentence)
        if words and len(words) != 3:
            bad.append(sentence)
    if bad:
        return "Megapensene trivocabular", "Cada megapensene deve ter exatamente 3 palavras."
    for sentence in sentences:
        if ":" in sentence:
            continue
        words = re.findall(r"[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç]+(?:-[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç]+)?", sentence)
        if not words:
            continue
        has_verb = any(
            norm(word) in {
                "e", "sao", "foi", "foram", "sera", "serao", "seria", "seriam", "esta", "estao",
                "gera", "geram", "cria", "criam", "qualifica", "qualificam", "sintetiza", "sintetizam",
                "evolui", "evoluem", "recicla", "reciclam", "expande", "expandem", "esclarece", "esclarecem",
            }
            or re.search(r"(ar|er|ir|ou|am|em|ava|iam|arao|erao|irao)$", norm(word))
            for word in words
        )
        if not has_verb:
            return "Megapensene sem verbo ou elipse", "Usar 1 verbo explicito ou elipse por dois-pontos no megapensene trivocabular."
    return None


def ortopensatologia_issue(text: str) -> tuple[str, str] | None:
    items = item_texts("Ortopensatologia", text)
    body = section_body(text)
    if not items and not body:
        return None
    if len(items) > 3:
        return "Limite de ortopensatas", "Usar no maximo 3 ortopensatas."
    keys = [alphabetic_key(item) for item in items]
    if len(keys) >= 3 and keys != sorted(keys):
        return "Ortopensatas em ordem alfabetica", "Ordenar as ortopensatas alfabeticamente."
    if items:
        if not all(("“" in item and "”" in item) or ('"' in item) for item in items):
            return "Formula da Ortopensatologia", "Em lista numerada, manter as ortopensatas entre aspas e preservar o subtitulo."
    elif body and not re.search(r"[:]\s*[–-]\s*[“\"]", body):
        return "Formula da Ortopensatologia", "Para ortopensata unica, usar dois-pontos, travessao e aspas antes da transcricao."
    return None


def enumerologia_issue(text: str) -> tuple[str, str] | None:
    qtd = count_items("Enumerologia", text)
    if qtd != 7:
        return "Enumerologia com total invalido", "A Enumerologia deve conter exatamente 7 itens."
    return None


def exemplologia_article_issue(text: str) -> tuple[str, str] | None:
    body = section_body(text)
    for part in [p for p in body.split(";") if p.strip()]:
        before_eq = part.split("=", 1)[0]
        if re.search(r"\b(?:o|a|os|as|um|uma|uns|umas)\b", before_eq, re.I):
            return "Artigo antes do sinal igual", "Na Exemplologia, usar artigos apenas apos o sinal '='."
    return None


def remissiologia_spacing_issue(text: str) -> tuple[str, str] | None:
    body = section_body(text)
    if count_items("Remissiologia", text) < 2:
        return None
    compact_items = re.findall(r"(?m)(?:^|\n)\s*\d{1,3}\.\s+[^\n]+", body)
    if len(compact_items) >= 2 and "\n\n" not in body:
        return "Espacamento da Remissiologia", "Usar espaçamento duplo entre os itens da Remissiologia."
    return None


def bibliography_issue(text: str) -> tuple[str, str] | None:
    items = item_texts("Bibliografia Especifica", text)
    if not items:
        return None
    bad = []
    for item in items:
        has_author = bool(re.match(r"(?:[A-ZÁÉÍÓÚÂÊÔÃÕÇ][\wÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç .,\'-]{2,}|Idem);", item))
        has_year = bool(re.search(r"\b(?:19|20)\d{2}\b", item))
        has_pages = bool(re.search(r"\b(?:p\.|págs?\.?)\s*\d+", item, re.I))
        if not (has_author and has_year and has_pages):
            bad.append(item)
    if bad:
        return "Padrao bibliografico", "Conferir autor em caixa alta, ano e paginacao no padrao bibliografico conscienciologico."
    return None


def initials_issue(paragraphs: list[str]) -> tuple[str, str] | None:
    meaningful = [p for p in paragraphs if not detect_division(p)]
    if not meaningful:
        return None
    tail = meaningful[-1].strip()
    compact_tail = re.sub(r"[\s.]+", "", tail)
    if re.fullmatch(r"[A-ZÁÉÍÓÚÂÊÔÃÕÇ]{2,3}", compact_tail):
        return None
    return "Iniciais do verbetografo", "Inserir ao final as iniciais do verbetografo em maiusculas, com 2 letras e no maximo 3."


def gender_pair_issue(sections: dict[str, str]) -> tuple[str, str] | None:
    if "Masculinologia" not in sections or "Femininologia" not in sections:
        return None
    masc = item_texts("Masculinologia", sections["Masculinologia"])
    fem = item_texts("Femininologia", sections["Femininologia"])
    if len(masc) != len(fem):
        return "Espelhamento Masculinologia x Femininologia", "Manter quantidade equivalente de itens nas secoes Masculinologia e Femininologia."
    return None


def logia_terms(text: str) -> list[str]:
    return re.findall(r"\b[A-ZÁÉÍÓÚÂÊÔÃÕÇ][\wÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç-]*logia\b", text)


def compute_logias(sections: dict[str, str]) -> tuple[int, list[str]]:
    terms: list[str] = []
    skip_body = {"Interdisciplinologia", "Remissiologia", "Frase Enfatica"}
    skip_title = {"Interdisciplinologia", "Frase Enfatica"}
    for sec, text in sections.items():
        if sec not in skip_title and re.search(r"logia$", sec, re.I):
            terms.append(sec)
        if sec in skip_body:
            continue
        terms.extend(logia_terms(section_body(text)))
    unique = sorted(set(terms), key=norm)
    return len(terms), unique


def add(achados: list[dict], severidade: str, secao: str, regra: str, evidencia: str, sugestao: str) -> None:
    achados.append(
        {
            "severidade": severidade,
            "secao": secao,
            "regra": regra,
            "evidencia": evidencia[:500],
            "sugestao": sugestao,
        }
    )


def build_sections(paragraphs: list[str]) -> tuple[dict[str, str], list[tuple[str, str | None]], list[str], bool]:
    current_div = None
    current_sec = None
    sections: dict[str, list[str]] = {}
    section_order: list[tuple[str, str | None]] = []
    divisions_seen: list[str] = []
    frase_enfatica_detectada = False

    for i, p in enumerate(paragraphs):
        div = detect_division(p)
        if div:
            current_div = div
            current_sec = None
            divisions_seen.append(div)
            continue

        sec = detect_independent_epigraph(p)
        if sec:
            current_sec = sec
            sections.setdefault(sec, []).append(p)
            section_order.append((sec, current_div))
            continue

        next_sec = detect_independent_epigraph(paragraphs[i + 1]) if i + 1 < len(paragraphs) else None
        if current_sec == "Remissiologia" and next_sec == "Questionologia" and len(p) > 40:
            frase_enfatica_detectada = True
            current_sec = "Frase Enfatica"
            sections.setdefault(current_sec, []).append(p)
            section_order.append((current_sec, current_div))
            continue

        if current_sec:
            sections.setdefault(current_sec, []).append(p)

    return {sec: "\n".join(parts) for sec, parts in sections.items()}, section_order, divisions_seen, frase_enfatica_detectada


def title_tokens(title: str) -> set[str]:
    stop = {"a", "o", "as", "os", "de", "do", "da", "dos", "das", "e"}
    return {token for token in re.findall(r"[a-z0-9áéíóúâêôãõç]+", norm(title)) if token not in stop and len(token) > 2}


def remove_quoted_text(text: str) -> str:
    text = re.sub(r"“[^”]*”", " ", text)
    text = re.sub(r'"[^"]*"', " ", text)
    return text


def char_flags(rich: dict) -> list[dict]:
    flags = []
    for run in rich["runs"]:
        for char in run["text"]:
            flags.append(
                {
                    "char": " " if char == "\xa0" else char,
                    "bold": run["bold"],
                    "italic": run["italic"],
                    "underline": run.get("underline", False),
                    "small_caps": run.get("small_caps", False),
                    "font": run.get("font") or "",
                    "size": run.get("size"),
                }
            )
    return flags


def run_text_from_flags(flags: list[dict]) -> str:
    return "".join(item["char"] for item in flags)


def is_bold_at(flags: list[dict], idx: int) -> bool:
    return 0 <= idx < len(flags) and bool(flags[idx]["bold"])


def previous_text_index(flags: list[dict], idx: int) -> int | None:
    for pos in range(idx - 1, -1, -1):
        if flags[pos]["char"].isalnum():
            return pos
    return None


def is_italic_at(flags: list[dict], idx: int) -> bool:
    return 0 <= idx < len(flags) and bool(flags[idx]["italic"])


def span_has_text(flags: list[dict], start: int, end: int) -> bool:
    return any(item["char"].isalnum() for item in flags[start:end])


def span_italic_ok(flags: list[dict], start: int, end: int) -> bool:
    chars = [item for item in flags[start:end] if item["char"].isalnum()]
    return bool(chars) and all(item["italic"] for item in chars)


def span_emphasis_ok(flags: list[dict], start: int, end: int) -> bool:
    chars = [item for item in flags[start:end] if item["char"].isalnum()]
    return bool(chars) and all(item["italic"] or item["underline"] for item in chars)


def span_plain_italic_ok(flags: list[dict], start: int, end: int) -> bool:
    chars = [item for item in flags[start:end] if item["char"].strip()]
    return bool(chars) and not any(item["italic"] for item in chars)


def explicit_run_check(rich: dict, prop: str, expected) -> bool | None:
    runs = [run for run in rich["runs"] if any(ch.isalnum() for ch in run["text"])]
    values = [run.get(prop) for run in runs if run.get(prop) not in {"", None}]
    if not values:
        return None
    if callable(expected):
        return all(expected(value) for value in values)
    return all(value == expected for value in values)


def bool_run_check(rich: dict, prop: str) -> bool:
    runs = [run for run in rich["runs"] if any(ch.isalnum() for ch in run["text"])]
    return bool(runs) and all(bool(run.get(prop)) for run in runs)


def paragraph_has_double_border(rich: dict) -> bool | None:
    borders = rich.get("borders") or []
    if not borders:
        return None
    return any(str(value).lower() == "double" for value in borders)


def twips_close(value: int, expected: int, tolerance: int = 90) -> bool:
    return abs(value - expected) <= tolerance


def find_text_span(text: str, needle: str) -> tuple[int, int] | None:
    if not needle:
        return None
    match = re.search(re.escape(needle), text, re.I)
    if match:
        return match.start(), match.end()
    compact_text = norm(text)
    compact_needle = norm(needle)
    idx = compact_text.find(compact_needle)
    if idx < 0:
        return None
    # Fallback aproximado: suficiente para evidência, mas não para edição fina.
    return None


def last_period_index(flags: list[dict]) -> int | None:
    for idx in range(len(flags) - 1, -1, -1):
        if flags[idx]["char"].isspace():
            continue
        return idx if flags[idx]["char"] == "." else None
    return None


def prefix_bold_ok(flags: list[dict], end_idx: int) -> bool:
    relevant = [item for item in flags[: end_idx + 1] if not item["char"].isspace()]
    return bool(relevant) and all(item["bold"] for item in relevant)


def numbered_item(text: str) -> bool:
    return bool(re.match(r"^\d{1,3}\.\s+", text.replace("\xa0", " ")))


def numbered_body_start(raw: str) -> int | None:
    match = re.match(r"^\s*\d{1,3}\.\s+", raw.replace("\xa0", " "))
    return match.end() if match else None


SECTION_COMPOSITE_TERMS = {
    "Sinergismologia": ("sinergismo",),
    "Principiologia": ("princípio", "principio"),
    "Codigologia": ("código", "codigo"),
    "Teoriologia": ("teoria",),
    "Tecnologia": ("técnica", "tecnica", "técnicas", "tecnicas"),
    "Laboratoriologia": ("laboratório", "laboratorio"),
    "Colegiologia": ("Colégio Invisível", "Colegio Invisivel", "colégio invisível", "colegio invisivel"),
    "Efeitologia": ("efeito",),
    "Ciclologia": ("ciclo",),
    "Binomiologia": ("binômio", "binomio"),
    "Interaciologia": ("interação", "interacao"),
    "Crescendologia": ("crescendo",),
    "Trinomiologia": ("trinômio", "trinomio"),
    "Polinomiologia": ("polinômio", "polinomio"),
    "Antagonismologia": ("antagonismo",),
    "Paradoxologia": ("paradoxo",),
    "Legislogia": ("lei",),
    "Sindromologia": ("síndrome", "sindrome"),
    "Mitologia": ("mito", "mitos"),
}

ITALIC_SUFFIXES = ("cracia", "filia", "fobia", "teca")


def iter_item_spans(raw: str, body_start: int = 0) -> list[tuple[int, int]]:
    spans = []
    start = body_start
    for match in re.finditer(r"[;,]", raw[body_start:]):
        end = body_start + match.start()
        if end > start:
            spans.append((start, end))
        start = body_start + match.end()
    if start < len(raw):
        spans.append((start, len(raw)))
    return [(s, e) for s, e in spans if raw[s:e].strip()]


def item_core_span(raw: str, start: int, end: int) -> tuple[int, int]:
    while start < end and not raw[start].isalnum():
        start += 1
    while end > start and not raw[end - 1].isalnum():
        end -= 1
    return start, end


def skip_leading_article(raw: str, start: int, end: int) -> int:
    match = re.match(r"\s*(?:o|a|os|as)\s+", raw[start:end], re.I)
    if match:
        return start + match.end()
    return start


def leading_article_span(raw: str, start: int, end: int) -> tuple[int, int] | None:
    match = re.match(r"\s*(?:a|as|o|os)\b", raw[start:end], re.I)
    if not match:
        return None
    article_start = start + match.start()
    while article_start < end and raw[article_start].isspace():
        article_start += 1
    return article_start, start + match.end()


def composite_expression_spans(section: str, raw: str, body_start: int) -> list[tuple[int, int]]:
    terms = SECTION_COMPOSITE_TERMS.get(section)
    if not terms:
        return []
    term_pattern = "|".join(re.escape(term) for term in terms)
    pattern = re.compile(rf"(?:^|[;:]\s*)\s*(?:o|a|os|as)\s+({term_pattern})\b", re.I)
    spans = []
    for match in pattern.finditer(raw[body_start:]):
        start = body_start + match.start(1)
        search_from = body_start + match.end(1)
        end_candidates = [idx for idx in (raw.find(";", search_from), raw.find(".", search_from)) if idx != -1]
        end = min(end_candidates) if end_candidates else len(raw)
        core_start, core_end = item_core_span(raw, start, end)
        if core_end > core_start:
            spans.append((core_start, core_end))
    return spans


def composite_signal_issue(section: str, text: str) -> tuple[str, str] | None:
    if section not in SECTION_COMPOSITE_TERMS:
        return None
    body = section_body(text)
    phrases = []
    for item in item_texts(section, text):
        pseudo = f"{section}: {item}"
        for start, end in composite_expression_spans(section, pseudo, len(section) + 1):
            phrases.append(pseudo[start:end])
    if not phrases and body:
        pseudo = f"{section}: {body}"
        phrases = [pseudo[start:end] for start, end in composite_expression_spans(section, pseudo, len(section) + 1)]
    for phrase in phrases:
        if section == "Antagonismologia":
            if "/" in phrase and not re.search(r"\s/\s", phrase):
                return "Barra do antagonismo", "Em Antagonismologia, usar barra com espaço antes e depois: ' / '."
            continue
        if "/" in phrase:
            return "Barra fora de Antagonismologia", "Usar barra apenas em Antagonismologia."
        if re.search(r"\w\s+-\s+\w", phrase):
            return "Hifen com espacos", "Em sublinhamentos, usar hifen sem espacos ou travessao quando houver expressao composta."
        term_pattern = "|".join(re.escape(term) for term in SECTION_COMPOSITE_TERMS.get(section, ()))
        core = re.sub(rf"^(?:{term_pattern})\b\s*", "", phrase, flags=re.I).strip()
        for match in re.finditer(r"(?<![–—])-(?![–—])", core):
            left = core[: match.start()].strip()
            right = core[match.end() :].strip()
            left_piece = re.split(r"[;–—/]", left)[-1].strip()
            right_piece = re.split(r"[;–—/]", right)[0].strip()
            if " " in left_piece or " " in right_piece:
                return "Travessao em expressao composta", "Quando houver expressao composta entre os elementos, usar travessao/en dash."
    return None


def analyze_docx_formatting(path: Path) -> dict:
    rich_paragraphs = read_docx_rich(path)
    layout = read_docx_layout(path)
    plain_paragraphs = [p["text"] for p in rich_paragraphs]
    sections, _, _, frase_enfatica_detectada = build_sections(plain_paragraphs)
    meta = extract_meta(plain_paragraphs, sections)
    checks = []
    ajustes = []

    def record(secao: str, regra: str, status: str, evidencia: str, sugestao: str = "") -> None:
        item = {"secao": secao, "regra": regra, "status": status, "evidencia": evidencia[:500], "sugestao": sugestao}
        checks.append(item)
        if status != "conforme":
            ajustes.append(item)

    if rich_paragraphs:
        for section in layout.get("sections", []):
            width = section.get("w")
            height = section.get("h")
            if width and height:
                letter = {width, height} == {12240, 15840}
                record("Layout", "Papel carta", "conforme" if letter else "ajustar", f"{width}x{height} twips", "Configurar papel carta.")
            margins = section.get("margins") or {}
            if margins:
                expected = {"top": 1706, "bottom": 1740, "left": 2137, "right": 2126}
                ok = all(key in margins and twips_close(margins[key], val) for key, val in expected.items())
                record("Layout", "Margens", "conforme" if ok else "ajustar", str(margins), "Configurar margens conforme a Chapa Verbetografica.")
                if margins.get("gutter"):
                    record("Layout", "Medianiz", "conforme", str(margins.get("gutter")))

        headers = layout.get("headers", [])
        if headers:
            header_text = " ".join(header.get("text", "") for header in headers)
            has_title = "Enciclopédia da Conscienciologia" in header_text or "Enciclopedia da Conscienciologia" in header_text
            record("Cabeçalho", "Texto do cabecalho", "conforme" if has_title else "ajustar", header_text[:500], 'Inserir "Enciclopédia da Conscienciologia" no cabecalho.')
            explicit_runs = [run for header in headers for run in header.get("runs", []) if run.get("font") or run.get("size") or run.get("italic")]
            if explicit_runs:
                ok = all(
                    (not run.get("font") or str(run.get("font")).lower() == "times new roman")
                    and (run.get("size") is None or abs(float(run.get("size")) - 9) < 0.1)
                    and bool(run.get("italic"))
                    for run in explicit_runs
                    if run.get("text", "").strip()
                )
                record("Cabeçalho", "Formato do cabecalho", "conforme" if ok else "ajustar", header_text[:500], "Aplicar Times New Roman 9 italico ao cabecalho.")
            has_page_field = any("PAGE" in " ".join(header.get("fields", [])) for header in headers)
            if has_page_field:
                record("Cabeçalho", "Paginacao", "conforme", "Campo PAGE detectado")

        title_rich = next((p for p in rich_paragraphs[:6] if not detect_division(p["text"]) and not detect_section(p["text"]) and not re.fullmatch(r"\(.+?logia\)", p["text"], re.I)), None)
        if title_rich:
            title_ok = bool_run_check(title_rich, "bold") and bool_run_check(title_rich, "italic") and bool_run_check(title_rich, "small_caps")
            font_ok = explicit_run_check(title_rich, "font", lambda value: str(value).lower() == "arial")
            size_ok = explicit_run_check(title_rich, "size", lambda value: abs(float(value) - 11) < 0.1)
            border_ok = paragraph_has_double_border(title_rich)
            if title_ok and font_ok is not False and size_ok is not False and border_ok is not False:
                record("Entrada", "Formato do titulo", "conforme", title_rich["text"])
            else:
                record("Entrada", "Formato do titulo", "ajustar", title_rich["text"], "Aplicar Arial 11, negrito-italico, versalete e borda dupla ao titulo.")

        specialty_rich = next((p for p in rich_paragraphs[:8] if re.fullmatch(r"\(.+?logia\)", p["text"], re.I)), None)
        if specialty_rich:
            specialty_ok = bool_run_check(specialty_rich, "small_caps")
            font_ok = explicit_run_check(specialty_rich, "font", lambda value: str(value).lower() == "arial")
            size_ok = explicit_run_check(specialty_rich, "size", lambda value: abs(float(value) - 11) < 0.1)
            spacing_ok = explicit_run_check(specialty_rich, "char_spacing", lambda value: abs(float(value) - 1.5) < 0.1)
            border_ok = paragraph_has_double_border(specialty_rich)
            if specialty_ok and font_ok is not False and size_ok is not False and spacing_ok is not False and border_ok is not False:
                record("Especialidade", "Formato da Especialidade", "conforme", specialty_rich["text"])
            else:
                record("Especialidade", "Formato da Especialidade", "ajustar", specialty_rich["text"], "Aplicar Arial 11, versalete, espacamento expandido de 1,5 pt e borda dupla na Especialidade.")

        common = [
            p for p in rich_paragraphs
            if detect_section(p["text"]) and not re.fullmatch(r"\(.+?logia\)", p["text"], re.I)
        ]
        explicit_font_runs = [run for p in common for run in p["runs"] if run.get("font") or run.get("size")]
        if explicit_font_runs:
            ok = all(
                (not run.get("font") or str(run.get("font")).lower() == "times new roman")
                and (run.get("size") is None or abs(float(run.get("size")) - 10) < 0.1)
                for run in explicit_font_runs
            )
            record("Texto", "Fonte do texto geral", "conforme" if ok else "ajustar", "Atributos explicitos em paragrafos comuns", "Aplicar Times New Roman 10 ao texto geral.")
        explicit_lines = [p.get("line_spacing") for p in common if p.get("line_spacing")]
        if explicit_lines:
            ok = all(str(value) == "240" for value in explicit_lines)
            record("Texto", "Espacamento simples", "conforme" if ok else "ajustar", ", ".join(map(str, explicit_lines[:10])), "Aplicar espacamento simples entre linhas.")

    current_format_sec = None
    for i, rich in enumerate(rich_paragraphs):
        text = rich["text"]
        flags = char_flags(rich)
        raw = run_text_from_flags(flags)
        if detect_division(text):
            current_format_sec = None
            continue
        secao = detect_section(text)
        if not secao:
            if current_format_sec and numbered_item(text):
                item_start = numbered_body_start(raw)
                item_colon_idx = raw.find(":", item_start or 0)
                final_dot = last_period_index(flags)
                if item_start is not None and item_colon_idx > item_start:
                    prev_idx = previous_text_index(flags, item_colon_idx)
                    item_epigraph_bold = prev_idx is not None and is_bold_at(flags, prev_idx) and is_bold_at(flags, item_colon_idx)
                    if item_epigraph_bold and final_dot is not None and is_bold_at(flags, final_dot):
                        record(current_format_sec, "Ponto final espelha epigrafe numerada", "conforme", text)
                    elif item_epigraph_bold and final_dot is not None:
                        record(
                            current_format_sec,
                            "Ponto final espelha epigrafe numerada",
                            "ajustar",
                            text,
                            "Como a epigrafe do item numerado termina com ':' em negrito, o ponto final do paragrafo tambem deve estar em negrito.",
                        )
            continue
        current_format_sec = secao

        prefix_match = re.match(r"^([^:.]{2,80})([:.])", raw)
        if prefix_match:
            prefix_end = prefix_match.end() - 1
            if prefix_bold_ok(flags, prefix_end):
                record(secao, "Epigrafe em negrito", "conforme", text)
            else:
                record(secao, "Epigrafe em negrito", "ajustar", text, "Aplicar negrito na epigrafe e no sinal imediato.")

        prefix_colon = re.match(r"^[^:.]{2,80}:", raw)
        final_dot = last_period_index(flags)
        if prefix_colon:
            colon_idx = prefix_colon.end() - 1
            if not is_bold_at(flags, colon_idx):
                record(secao, "Dois-pontos introdutorio", "ajustar", text, "Aplicar negrito no ':' que abre a enumeracao.")
            elif final_dot is not None and secao in {"Sinonimologia", "Antonimologia"}:
                if is_bold_at(flags, final_dot):
                    record(secao, "Ponto final normal", "ajustar", text, "Em Sinonimologia e Antonimologia, o ponto final deve ficar em fonte normal, sem negrito.")
                else:
                    record(secao, "Ponto final normal", "conforme", text)
            elif final_dot is not None and is_bold_at(flags, final_dot):
                record(secao, "Ponto final espelha dois-pontos", "conforme", text)
            elif final_dot is not None:
                record(secao, "Ponto final espelha dois-pontos", "ajustar", text, "Como o ':' introdutorio esta em negrito, o ponto final do paragrafo tambem deve estar em negrito.")

        intro_colon_idx = raw.rfind(":")
        next_is_numbered = i + 1 < len(rich_paragraphs) and numbered_item(rich_paragraphs[i + 1]["text"])
        if intro_colon_idx > 0 and next_is_numbered:
            prev_idx = previous_text_index(flags, intro_colon_idx)
            colon_requires_bold = prev_idx is not None and is_bold_at(flags, prev_idx)
            if secao == "Remissiologia":
                if is_bold_at(flags, intro_colon_idx):
                    record(secao, "Dois-pontos introdutorio normal", "ajustar", text, "Na Remissiologia, o ':' que abre a lista vertical deve ficar em fonte normal, sem negrito.")
                else:
                    record(secao, "Dois-pontos introdutorio normal", "conforme", text)
            else:
                if not colon_requires_bold and not is_bold_at(flags, intro_colon_idx):
                    record(secao, "Dois-pontos introdutorio de lista vertical", "conforme", text)
                elif colon_requires_bold and is_bold_at(flags, intro_colon_idx):
                    record(secao, "Dois-pontos introdutorio de lista vertical", "conforme", text)
                elif colon_requires_bold:
                    record(secao, "Dois-pontos introdutorio de lista vertical", "ajustar", text, "Aplicar negrito no ':' porque ele vem logo apos palavra em negrito.")
                else:
                    record(secao, "Dois-pontos introdutorio de lista vertical", "ajustar", text, "Remover negrito do ':' porque ele encerra paragrafo introdutorio em fonte normal.")

                last_item = None
                j = i + 1
                while j < len(rich_paragraphs) and numbered_item(rich_paragraphs[j]["text"]):
                    last_item = rich_paragraphs[j]
                    j += 1
                if last_item and colon_requires_bold and is_bold_at(flags, intro_colon_idx):
                    last_flags = char_flags(last_item)
                    last_dot = last_period_index(last_flags)
                    if last_dot is not None and is_bold_at(last_flags, last_dot):
                        record(secao, "Ponto final do ultimo item", "conforme", last_item["text"])
                    elif last_dot is not None:
                        record(secao, "Ponto final do ultimo item", "ajustar", last_item["text"], "O ponto final do ultimo item deve ficar em negrito quando a lista e aberta por ':' em negrito.")

        if prefix_match and prefix_match.group(2) == "." and final_dot is not None and final_dot != prefix_match.end() - 1:
            if is_bold_at(flags, final_dot):
                record(secao, "Ponto final de prosa", "ajustar", text, "Em paragrafo de prosa sem ':' introdutorio, o ponto final deve ficar plano.")
            else:
                record(secao, "Ponto final de prosa", "conforme", text)

        body_start = prefix_match.end() if prefix_match else 0

        if secao in {"Definologia", "Remissiologia"} and meta.get("titulo"):
            span = find_text_span(raw[body_start:], meta["titulo"])
            if span:
                start, end = body_start + span[0], body_start + span[1]
                if span_emphasis_ok(flags, start, end):
                    record(secao, "Titulo citado em italico", "conforme", text)
                else:
                    record(secao, "Titulo citado em italico", "ajustar", text, "Aplicar italico/sublinhamento ao titulo do verbete citado nesta secao.")

        composite_spans = composite_expression_spans(secao, raw, body_start)
        bad_composite = None
        checked_composite = None
        for start, end in composite_spans:
            if not span_has_text(flags, start, end):
                continue
            checked_composite = raw[start:end]
            if not span_italic_ok(flags, start, end):
                bad_composite = raw[start:end]
                break
        if bad_composite:
            record(secao, "Expressao composta em italico", "ajustar", bad_composite, "Aplicar italico da palavra de classe ate o proximo ';' ou '.'.")
        elif checked_composite:
            record(secao, "Expressao composta em italico", "conforme", checked_composite)

        semicolon_indexes = [] if secao == "Etimologia" else [idx for idx, char in enumerate(raw) if char == ";"]
        if semicolon_indexes:
            article_checked = False
            bad_article = None
            for start, end in iter_item_spans(raw, body_start):
                core_start, core_end = item_core_span(raw, start, end)
                article_span = leading_article_span(raw, core_start, core_end)
                if not article_span:
                    continue
                article_checked = True
                article_start, article_end = article_span
                if not span_plain_italic_ok(flags, article_start, article_end):
                    bad_article = raw[article_start:article_end]
                    break
            if bad_article:
                record(secao, "Artigo sem italico", "ajustar", text, "Remover italico do artigo inicial do item em listagem horizontal separada por ';'.")
            elif article_checked:
                record(secao, "Artigo sem italico", "conforme", text)

            bad_separator = False
            for idx in semicolon_indexes:
                if is_italic_at(flags, idx):
                    bad_separator = True
                    break
            if bad_separator:
                record(secao, "Separador sem italico", "ajustar", text, "Remover italico do separador ';'.")
            else:
                record(secao, "Separador sem italico", "conforme", text)

        for suffix in ITALIC_SUFFIXES:
            suffix_hits = list(re.finditer(rf"\b[\wÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç-]*{suffix}\b", raw, re.I))
            if len(suffix_hits) < 7:
                continue
            bad_suffix = None
            for match in suffix_hits:
                suffix_start = match.end() - len(suffix)
                if not span_italic_ok(flags, suffix_start, match.end()):
                    bad_suffix = match.group(0)
                    break
            if bad_suffix:
                record(secao, "Sufixo repetido em italico", "ajustar", text, f"Aplicar italico ao sufixo '{suffix}' quando houver 7 ou mais itens com esse sufixo.")
            else:
                record(secao, "Sufixo repetido em italico", "conforme", text)

    if frase_enfatica_detectada and "Frase Enfatica" in sections:
        frase_text = clean(sections["Frase Enfatica"])
        rich = next((p for p in rich_paragraphs if clean(p["text"]) == frase_text), None)
        if rich:
            flags = char_flags(rich)
            if rich.get("align") == "center":
                record("Frase Enfatica", "Centralizacao", "conforme", rich["text"])
            else:
                record("Frase Enfatica", "Centralizacao", "ajustar", rich["text"], "Centralizar a Frase Enfatica.")

            internal_grifo = [item for item in flags if item["char"].isalnum() and item["italic"] and not item["bold"]]
            if internal_grifo:
                record("Frase Enfatica", "Grifo interno em italico sem negrito", "conforme", rich["text"])
            else:
                record("Frase Enfatica", "Grifo interno em italico sem negrito", "ajustar", rich["text"], "Aplicar ao menos um grifo interno em italico sem negrito.")

            raw_frase = rich.get("raw_text", rich["text"])
            word_gaps = re.findall(r"[A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç)][ ]+['\"“”A-Za-zÁÉÍÓÚÂÊÔÃÕÇáéíóúâêôãõç(]", raw_frase)
            if word_gaps:
                ok = all("  " in gap for gap in word_gaps)
                record("Frase Enfatica", "Dois espacos entre palavras", "conforme" if ok else "ajustar", raw_frase, "Usar dois espacos entre as palavras da Frase Enfatica.")
            if "\n" in raw_frase:
                lines = [line for line in raw_frase.splitlines() if line.strip()]
                record("Frase Enfatica", "Quatro linhas", "conforme" if len(lines) == 4 else "ajustar", raw_frase, "Distribuir a Frase Enfatica em 4 linhas.")

    return {
        "tipo": "docx",
        "verificados": len(checks),
        "conformes": sum(1 for item in checks if item["status"] == "conforme"),
        "ajustes": ajustes,
        "checks": checks,
    }


def analyze(paragraphs: list[str], warnings: list[str] | None = None) -> dict:
    warnings = warnings or []
    sections, section_order, divisions_seen, frase_enfatica_detectada = build_sections(paragraphs)
    meta = extract_meta(paragraphs, sections)
    achados: list[dict] = []

    for warning in warnings:
        add(achados, "Importante", "Documento", "Aviso de leitura", warning, "Use DOCX para a analise mais confiavel.")

    missing_divs = [d for d in DIVISOES if d not in divisions_seen]
    for div in missing_divs:
        add(achados, "Crítico", div, "Divisao obrigatoria ausente", div, "Inserir a divisao na ordem canonica.")

    seen_index = [DIVISOES.index(d) for d in divisions_seen if d in DIVISOES]
    if seen_index != sorted(seen_index):
        add(achados, "Crítico", "Estrutura", "Divisoes fora de ordem", " > ".join(divisions_seen), "Reordenar as divisoes I a VI.")

    by_div: dict[str, list[tuple[str, int]]] = {}
    for sec, declared_div in section_order:
        rule = SECOES.get(sec, {})
        if not declared_div or "ordem" not in rule:
            continue
        by_div.setdefault(declared_div, []).append((sec, int(rule["ordem"])))
    for div, items in by_div.items():
        ordem = [item[1] for item in items]
        if ordem != sorted(ordem):
            add(
                achados,
                "Importante",
                div,
                "Secoes fora de ordem na divisao",
                " > ".join(item[0] for item in items),
                "Reordenar as secoes conforme a sequencia canonica da divisao.",
            )

    for sec in SECOES_FIXAS:
        if sec == "Frase Enfatica":
            found = frase_enfatica_detectada or any("?" not in p and detect_section(p) is None and len(p) > 40 and p.upper() == p for p in paragraphs)
            if not found:
                add(achados, "Importante", sec, "Secao fixa nao detectada", sec, "Confirmar presenca da Frase Enfatica centralizada.")
            continue
        if sec not in sections:
            add(achados, "Crítico", sec, "Secao fixa ausente", sec, "Adicionar a secao fixa.")

    for sec, declared_div in section_order:
        expected = SECOES.get(sec, {}).get("div")
        if expected and declared_div and expected != declared_div:
            add(achados, "Importante", sec, "Secao em divisao inesperada", f"{sec} em {declared_div}", f"Mover para {expected}.")

    counts = {}
    maximos = []
    for sec, text in sections.items():
        qtd = count_items(sec, text)
        counts[sec] = qtd
        rule = SECOES.get(sec, {})
        if rule.get("min") and qtd < rule["min"]:
            add(achados, "Importante", sec, "Minimo de itens nao atingido", f"{qtd} item(ns)", f"Completar ao menos {rule['min']} item(ns).")
        if rule.get("totais") and qtd not in rule["totais"]:
            add(achados, "Crítico", sec, "Total invalido", f"{qtd} item(ns)", f"Usar um dos totais permitidos: {sorted(rule['totais'])}.")
        if rule.get("always_max") or (rule.get("max") and qtd >= rule["max"]):
            maximos.append({"secao": sec, "itens": qtd, "limiar": rule["max"]})
        zero_issue = check_zero_pattern(sec, text)
        if zero_issue:
            regra, sugestao = zero_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        alpha_issue = alphabetical_issue(sec, text)
        if alpha_issue:
            regra, sugestao = alpha_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        spacing_issue = numbered_spacing_issue(sec, text)
        if spacing_issue:
            regra, sugestao = spacing_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        separator_issue = sinon_anton_separator_issue(sec, text)
        if separator_issue:
            regra, sugestao = separator_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        aspas_issue = quote_issue(sec, text)
        if aspas_issue:
            regra, sugestao = aspas_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Megapensenologia":
            mega_issue = megapensene_issue(text)
            if mega_issue:
                regra, sugestao = mega_issue
                add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Enumerologia":
            enum_issue = enumerologia_issue(text)
            if enum_issue:
                regra, sugestao = enum_issue
                add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Exemplologia":
            exemplo_issue = exemplologia_article_issue(text)
            if exemplo_issue:
                regra, sugestao = exemplo_issue
                add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Remissiologia":
            remiss_spacing = remissiologia_spacing_issue(text)
            if remiss_spacing:
                regra, sugestao = remiss_spacing
                add(achados, "Refinamento", sec, regra, text, sugestao)
        signal_issue = composite_signal_issue(sec, text)
        if signal_issue:
            regra, sugestao = signal_issue
            add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Ortopensatologia":
            orto_issue = ortopensatologia_issue(text)
            if orto_issue:
                regra, sugestao = orto_issue
                add(achados, "Importante", sec, regra, text, sugestao)
        if sec == "Bibliografia Especifica":
            biblio_issue = bibliography_issue(text)
            if biblio_issue:
                regra, sugestao = biblio_issue
                add(achados, "Importante", sec, regra, text, sugestao)

    detalhismo_count = sum(1 for sec in sections if SECOES.get(sec, {}).get("div") == "Detalhismo")
    if detalhismo_count >= 20:
        maximos.append({"secao": "Divisão Detalhismo", "itens": detalhismo_count, "limiar": 20})

    if "Questionologia" in sections and "leitor ou leitora" not in norm(sections["Questionologia"]):
        add(achados, "Importante", "Questionologia", "Formula de interlocucao", sections["Questionologia"], 'Usar a expressao "leitor ou leitora".')

    gender_issue = gender_pair_issue(sections)
    if gender_issue:
        regra, sugestao = gender_issue
        add(
            achados,
            "Importante",
            "Perfilologia",
            regra,
            f"Masculinologia: {sections.get('Masculinologia', '')}\nFemininologia: {sections.get('Femininologia', '')}",
            sugestao,
        )

    title_norm = norm(meta.get("titulo", ""))
    if title_norm:
        for sec in ["Definologia", "Neologia", "Exemplologia", "Remissiologia"]:
            section_norm = norm(sections.get(sec, ""))
            if sec == "Exemplologia":
                tokens = title_tokens(meta.get("titulo", ""))
                if tokens and len(tokens & title_tokens(sections.get(sec, ""))) >= max(1, min(2, len(tokens))):
                    continue
            if sec in sections and title_norm not in section_norm:
                add(achados, "Importante", sec, "Correlacao com titulo", sections[sec], "Repetir o titulo do verbete de modo coerente nesta secao.")

    specialty_norm = norm(meta.get("especialidade", ""))
    if specialty_norm:
        for sec in ["Neologia", "Interdisciplinologia"]:
            if sec in sections and specialty_norm not in norm(sections[sec]):
                add(achados, "Importante", sec, "Correlacao com Especialidade", sections[sec], "Repetir a Especialidade declarada nesta secao.")

    parasita_hits = []
    full_text = " ".join(paragraphs)
    parasita_text = remove_quoted_text(full_text)
    for word in PARASITAS:
        if re.search(rf"\b{re.escape(word)}\b", parasita_text, re.I):
            parasita_hits.append(word)
    if parasita_hits:
        add(achados, "Refinamento", "Confor", "Parasitas da linguagem", ", ".join(sorted(set(parasita_hits))[:20]), "Revisar ocorrencias fora de citacoes, titulos e nomes proprios.")

    logias, logias_lista = compute_logias(sections)
    marca = {
        "paginas": {"exigido": MARCA_EXCELENCIA["paginas"], "encontrado": meta.get("paginas"), "ok": False if meta.get("paginas") else None},
        "maximos": {"exigido": MARCA_EXCELENCIA["maximos"], "encontrado": len(maximos), "ok": len(maximos) >= MARCA_EXCELENCIA["maximos"]},
        "logias": {"exigido": MARCA_EXCELENCIA["logias"], "encontrado": logias, "ok": logias >= MARCA_EXCELENCIA["logias"]},
    }
    return {
        "meta": meta,
        "marca_excelencia": marca,
        "secoes_presentes": list(sections.keys()),
        "secoes_texto": sections,
        "contagens": counts,
        "maximos": maximos,
        "logias": logias,
        "logias_lista": logias_lista,
        "achados": achados,
        "formatacao": {"tipo": "texto", "verificados": 0, "conformes": 0, "ajustes": [], "checks": []},
    }


def render_markdown(result: dict) -> str:
    meta = result["meta"]
    marca = result["marca_excelencia"]
    achados = result["achados"]
    crit = [a for a in achados if a["severidade"] == "Crítico"]
    imp = [a for a in achados if a["severidade"] == "Importante"]
    ref = [a for a in achados if a["severidade"] == "Refinamento"]

    def status(ok):
        return "✅" if ok is True else "❌" if ok is False else "⚠️"

    if crit or imp:
        prioridade = "Corrigir os impedimentos formais antes da revisão semântica fina."
    elif ref:
        prioridade = "Conferir os refinamentos editoriais apontados."
    else:
        prioridade = "Prosseguir para a leitura semântica editorial."

    def checklist_status(ok: bool) -> str:
        return "✅" if ok else "⚠️"

    maximos_lista = "; ".join(f"{item['secao']} ({item['itens']})" for item in result.get("maximos", []))
    paginas_encontradas = marca["paginas"]["encontrado"]
    paginas_evidencia = f"{paginas_encontradas} página(s)" if paginas_encontradas is not None else "Não disponível no arquivo analisado"
    logias_lista = result.get("logias_lista", [])
    logias_evidencia = "; ".join(logias_lista) if logias_lista else f"{result.get('logias', 0)} logias detectadas"

    lines = [
        "# Relatório de Conformidade Verbetográfica",
        f"**Verbete:** {meta.get('titulo') or 'não identificado'} · **Especialidade declarada:** {meta.get('especialidade') or 'não identificada'} · **Tema central:** {meta.get('tema_central') or 'não identificado'}",
        "",
        "## 1. Síntese executiva",
        "| Checklist | Situação | Resultado | Encaminhamento |",
        "|---|---|---|---|",
        f"| Achados críticos | {checklist_status(len(crit) == 0)} | {len(crit)} | {'Sem bloqueio crítico.' if not crit else 'Corrigir antes de avançar.'} |",
        f"| Achados importantes | {checklist_status(len(imp) == 0)} | {len(imp)} | {'Sem ajuste importante.' if not imp else 'Priorizar correções formais.'} |",
        f"| Refinamentos | {'✅' if not ref else '💡'} | {len(ref)} | {'Sem refinamento automatizado.' if not ref else 'Conferir em revisão editorial.'} |",
        f"| Marca de excelência | {status(marca['paginas']['ok'] and marca['maximos']['ok'] and marca['logias']['ok'])} | Páginas: {paginas_encontradas if paginas_encontradas is not None else 'verificar'}; Máximos: {marca['maximos']['encontrado']}; Logias: {marca['logias']['encontrado']} | {prioridade} |",
        "",
        "## 2. Painel da marca de excelência",
        "| Indicador | Exigido | Encontrado | Situação | Itens encontrados no texto |",
        "|---|---:|---:|---|---|",
    ]
    evidencias_marca = {
        "paginas": paginas_evidencia,
        "maximos": maximos_lista or "Nenhum máximo detectado",
        "logias": logias_evidencia,
    }
    for key, label in [("paginas", "Páginas"), ("maximos", "Máximos"), ("logias", "Logias")]:
        item = marca[key]
        found = item["encontrado"] if item["encontrado"] is not None else "verificar"
        lines.append(f"| {label} | {item['exigido']} | {found} | {status(item['ok'])} | {evidencias_marca[key]} |")

    lines += ["", "## 3. Confor (estilística e formatação)"]
    confor = [a for a in achados if a["secao"] in {"Confor", "Documento"} or "Minimo" in a["regra"] or "Total" in a["regra"]]
    formatacao = result.get("formatacao", {})
    if formatacao.get("verificados"):
        ajustes = formatacao.get("ajustes", [])
        lines.append(
            f"- **Auditoria DOCX de pontuação, negrito, itálico e estilo:** {formatacao.get('conformes', 0)} de {formatacao.get('verificados', 0)} verificações conformes."
        )
        if ajustes:
            grouped: dict[tuple[str, str], dict] = {}
            for item in ajustes:
                key = (item["regra"], item["sugestao"])
                grouped.setdefault(key, {"regra": item["regra"], "sugestao": item["sugestao"], "secoes": set(), "qtd": 0})
                grouped[key]["secoes"].add(item["secao"])
                grouped[key]["qtd"] += 1
            lines += [
                "",
                "| Regra | Ocorrências | Seções afetadas | Correção resumida |",
                "|---|---:|---|---|",
            ]
            for group in grouped.values():
                secoes = "; ".join(sorted(group["secoes"], key=norm))
                lines.append(f"| {group['regra']} | {group['qtd']} | {secoes} | {group['sugestao']} |")
        else:
            lines.append("- Pontos finais, epígrafes e sinais introdutórios verificados estão conformes nas regras automatizadas.")
    elif formatacao.get("tipo") and formatacao.get("tipo") != "docx":
        lines.append("- Auditoria fina de negrito/itálico não executada: envie DOCX para leitura de runs do Word.")
    if confor:
        lines += ["", "| Tipo | Seção | Regra | Correção resumida |", "|---|---|---|---|"]
        for a in confor:
            lines.append(f"| {a['severidade']} | {a['secao']} | {a['regra']} | {a['sugestao']} |")
    elif not formatacao.get("verificados"):
        lines.append("- Nenhum achado formal específico nesta execução.")

    lines += ["", "## 4. Correlações internas"]
    corr = [a for a in achados if "Correlacao" in a["regra"] or "Formula" in a["regra"]]
    lines += [f"- **{a['severidade']} · {a['secao']}:** {a['sugestao']}" for a in corr] or ["- Sem divergências automáticas nas correlações string-matcháveis."]

    lines += ["", "## 5. Seções fixas ausentes ou fora de ordem"]
    fixed = [a for a in achados if "ausente" in a["regra"].lower() or "fora de ordem" in a["regra"].lower()]
    lines += [f"- **{a['severidade']} · {a['secao']}:** {a['sugestao']}" for a in fixed] or ["- Não foram detectadas ausências críticas de seções fixas."]

    lines += [
        "",
        "## 6. Análise semântica editorial",
        "- Etapa não executada pelo motor Python. Use `prompts/codex-analise-semantica.md` com o verbete e este JSON formal para complementar o relatório.",
        "",
        "## 7. Lista priorizada de correções",
    ]
    for i, a in enumerate(crit + imp + ref, 1):
        icon = "❌ Crítico" if a["severidade"] == "Crítico" else "⚠️ Importante" if a["severidade"] == "Importante" else "💡 Refinamento"
        lines.append(f"{i}. [{icon}] {a['secao']}: {a['sugestao']}")
    if not achados:
        lines.append("1. [💡 Refinamento] Prosseguir com revisão semântica editorial.")
    return "\n".join(lines)


def analyze_file(path: str | Path) -> dict:
    path = Path(path)
    paragraphs, warnings = read_document(path)
    if not paragraphs:
        return analyze([], warnings)
    result = analyze(paragraphs, warnings)
    pages = None
    if path.suffix.lower() == ".docx":
        pages = read_docx_page_count(path)
    elif path.suffix.lower() == ".pdf":
        try:
            from pypdf import PdfReader

            pages = len(PdfReader(str(path)).pages)
        except Exception:
            pages = None
    if pages:
        result["meta"]["paginas"] = pages
        result["marca_excelencia"]["paginas"]["encontrado"] = pages
        result["marca_excelencia"]["paginas"]["ok"] = pages >= MARCA_EXCELENCIA["paginas"]
    if path.suffix.lower() == ".docx":
        result["formatacao"] = analyze_docx_formatting(path)
        for item in result["formatacao"]["ajustes"]:
            add(
                result["achados"],
                "Importante",
                item["secao"],
                item["regra"],
                item["evidencia"],
                item["sugestao"],
            )
    return result


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser()
    parser.add_argument("arquivo")
    parser.add_argument("--resumo", action="store_true")
    parser.add_argument("--markdown", action="store_true")
    args = parser.parse_args()
    result = analyze_file(args.arquivo)
    if args.markdown:
        print(render_markdown(result))
    elif args.resumo:
        meta = result["meta"]
        print(f"Verbete: {meta.get('titulo') or 'nao identificado'}")
        print(f"Achados: {len(result['achados'])}")
        print(f"Maximos: {len(result['maximos'])}")
        print(f"Logias: {result['logias']}")
    else:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
