from __future__ import annotations

from html import escape
import re
import tempfile
from pathlib import Path

import streamlit as st

from scripts.analisar_verbete import analyze_file, clean, convert_doc_to_docx, detect_division, detect_independent_epigraph, detect_section, read_docx_rich, render_markdown, is_inside_parentheses, last_period_index, logia_terms, section_body, get_effective_section, clean_sec, starts_with_argumentologia_epigraph
from scripts.catalogo import SECOES
from scripts.editar_verbete import gerar_pdf_marcado, gerar_verbete_corrigido, gerar_verbete_marcado
from scripts.relatorio_word import markdown_to_report_docx

st.set_page_config(
    page_title="Revisor Verbetográfico",
    page_icon="",
    layout="wide",
)



st.markdown(
    """
    <style>
    :root {
      --bg: #f6f3ee;
      --surface: #fffdf9;
      --ink: #23211f;
      --muted: #756f66;
      --line: #e7dfd3;
      --accent: #406f63;
      --accent-soft: #e6f0ec;
      --violet: #6f4c7d;
      --violet-soft: #e6e0e6;
      --warn: #a9672b;
      --danger: #a54845;
    }
    .stApp { background: var(--bg); color: var(--ink); }
    [data-testid="stSidebar"] { display: none; }
    .block-container {
      padding-top: 2.2rem;
      padding-left: 1.5rem !important;
      padding-right: 1.5rem !important;
    }
    @media (min-width: 2000px) {
      .block-container {
        padding-left: 1rem !important;
        padding-right: 1rem !important;
      }
    }
    h1, h2, h3 { letter-spacing: 0; color: var(--ink); }
    h1 { font-size: 2.1rem !important; line-height: 1.1 !important; margin-bottom: .35rem !important; }
    h2 { font-size: 1.25rem !important; margin-top: 1.2rem !important; }
    p, li, label, div { font-family: Inter, Segoe UI, sans-serif; }
    .main-title-shell {
      position: relative;
      background: var(--surface);
      border: 1px solid #d8e3ee;
      border-left: 5px solid #1f6fbd;
      border-radius: 8px;
      padding: 22px 26px 20px;
      margin: 18px 0 18px;
      box-shadow: 0 18px 42px rgba(31, 73, 125, .08);
    }
    .main-kicker {
      color: #5b718a;
      font-size: .76rem;
      font-weight: 750;
      text-transform: uppercase;
      letter-spacing: .08em;
      margin-bottom: 6px;
    }
    .main-title {
      color: #155da8 !important;
      font-size: 2.15rem !important;
      line-height: 1.12 !important;
      margin: 0 !important;
      font-weight: 800;
    }
    .main-subtitle {
      color: #5e6874;
      margin: .45rem 0 0;
      max-width: 780px;
      font-size: .98rem;
    }
    .metric-row { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; margin: 14px 0 20px; }
    .metric-card {
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 15px 16px;
    }
    .metric-label { color: var(--muted); font-size: .78rem; text-transform: uppercase; letter-spacing: .04em; }
    .metric-value { color: var(--ink); font-weight: 750; font-size: 1.55rem; margin-top: 4px; }
    .status-ok { color: #2f725b; }
    .status-warn { color: var(--warn); }
    .status-danger { color: var(--danger); }
    .soft-panel {
      background: rgba(255,253,249,.72);
      border: 1px solid var(--line);
      border-radius: 8px;
      padding: 18px;
      margin-bottom: 16px;
    }
    div[data-testid="stFileUploader"] section {
      background: var(--surface);
      border: 1px dashed #b9aa97;
      border-radius: 8px;
    }
    .stButton > button[kind="primary"], .stDownloadButton > button {
      border-radius: 7px;
      border: 1px solid #315d53;
      background: var(--accent);
      color: white;
      font-weight: 650;
    }
    .stButton > button[kind="primary"]:hover, .stDownloadButton > button:hover {
      border-color: #244d45;
      background: #315d53;
      color: white;
    }
    .stButton > button[kind="secondary"] {
      border-radius: 7px;
      border: 1px solid var(--line);
      background: var(--surface);
      color: var(--muted);
      font-weight: 500;
    }
    .stButton > button[kind="secondary"]:hover {
      border-color: var(--muted);
      color: var(--ink);
    }
    .st-key-report_toggle_buttons .stButton > button[kind="primary"] {
      border: 1px solid var(--violet);
      background: var(--violet);
      color: white;
    }
    .st-key-report_toggle_buttons .stButton > button[kind="primary"]:hover {
      border-color: #5a3d66;
      background: #5a3d66;
      color: white;
    }
    div[data-testid="stDataFrame"] { border: 1px solid var(--line); border-radius: 8px; overflow: hidden; }
    code, pre { background: #f0ebe3 !important; border-radius: 6px; }
    div[data-testid="stMarkdownContainer"] table {
      background: #ffffff;
    }
    div[data-testid="stMarkdownContainer"] table th,
    div[data-testid="stMarkdownContainer"] table td {
      background: #ffffff;
    }
    .wrap-table {
      width: 100%;
      border-collapse: collapse;
      table-layout: fixed;
      background: var(--surface);
      border: 1px solid var(--line);
      border-radius: 8px;
      overflow: hidden;
      font-size: .9rem;
    }
    .wrap-table th {
      background: #eee7dc;
      color: var(--ink);
      text-align: left;
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
    }
    .wrap-table td {
      padding: 10px 12px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      white-space: normal;
      overflow-wrap: anywhere;
      word-break: normal;
      line-height: 1.38;
    }
    .wrap-table tr:last-child td { border-bottom: 0; }
    .wrap-table td:nth-child(7), .wrap-table th:nth-child(7) {
      white-space: nowrap !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


def status_class(value: bool | None) -> str:
    if value is True:
        return "status-ok"
    if value is False:
        return "status-danger"
    return "status-warn"


def status_text(value: bool | None) -> str:
    if value is True:
        return "Conforme"
    if value is False:
        return "Ajustar"
    return "Verificar"


def render_wrap_table(
    rows: list[dict],
    columns: list[str] | None = None,
    widths: list[int] | None = None,
    html_columns: set[str] | None = None,
) -> None:
    if not rows:
        return
    columns = columns or list(rows[0].keys())
    html_columns = html_columns or set()
    colgroup = ""
    if widths and len(widths) == len(columns):
        colgroup = "<colgroup>" + "".join(f'<col style="width: {width}%;">' for width in widths) + "</colgroup>"
    header = "".join(f"<th>{escape(str(label))}</th>" for label in columns)
    def cell_value(row: dict, label: str) -> str:
        value = str(row.get(label, ""))
        if label in html_columns:
            return value.replace("\n", "<br>")
        return escape(value).replace(chr(10), "<br>").replace("; ", ";<br>")

    body = "".join(
        "<tr>"
        + "".join(f"<td>{cell_value(row, label)}</td>" for label in columns)
        + "</tr>"
        for row in rows
    )
    st.markdown(f'<table class="wrap-table">{colgroup}<thead><tr>{header}</tr></thead><tbody>{body}</tbody></table>', unsafe_allow_html=True)


def rich_runs_to_html(runs: list[dict]) -> str:
    parts = []
    for run in runs:
        text = escape(str(run.get("text", "")))
        if not text:
            continue
        if run.get("highlight"):
            text = f'<span style="background-color: #ffeba0; color: #333333; padding: 1px 3px; border-radius: 3px;">{text}</span>'
        if run.get("bold"):
            text = f"<strong>{text}</strong>"
        if run.get("italic"):
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts)


def find_rich_paragraph(rich_paragraphs: list[dict], evidence: str) -> dict | None:
    target = clean(evidence)
    if not target:
        return None
    target_prefix = target[:180]
    for p in rich_paragraphs:
        text = clean(p["text"])
        if text == target or text.startswith(target_prefix) or target.startswith(text[:180]):
            return p
    for p in rich_paragraphs:
        text = clean(p["text"])
        if target in text:
            return p
    return None


def highlight_runs_span(rich: dict, start_idx: int, end_idx: int) -> None:
    if start_idx >= end_idx or start_idx < 0:
        return
    import copy
    new_runs = []
    cursor = 0
    for run in rich["runs"]:
        run_text = run["text"]
        run_len = len(run_text)
        run_end = cursor + run_len
        
        overlap_start = max(cursor, start_idx)
        overlap_end = min(run_end, end_idx)
        
        if overlap_start < overlap_end:
            local_start = overlap_start - cursor
            local_end = overlap_end - cursor
            
            if local_start > 0:
                r1 = copy.deepcopy(run)
                r1["text"] = run_text[:local_start]
                new_runs.append(r1)
                
            r2 = copy.deepcopy(run)
            r2["text"] = run_text[local_start:local_end]
            r2["highlight"] = True
            new_runs.append(r2)
            
            if local_end < run_len:
                r3 = copy.deepcopy(run)
                r3["text"] = run_text[local_end:]
                new_runs.append(r3)
        else:
            new_runs.append(run)
            
        cursor = run_end
    rich["runs"] = new_runs


def target_spans(rule: str, raw_text: str, suggestion: str, evidence_text: str) -> list[tuple[int, int]]:
    import re
    rule_lower = rule.lower()
    spans = []
    
    if "ponto final" in rule_lower:
        for idx in range(len(raw_text) - 1, -1, -1):
            if raw_text[idx] == ".":
                spans.append((idx, idx + 1))
                break
    elif "dois-pontos" in rule_lower or "dois pontos" in rule_lower:
        idx = raw_text.rfind(":")
        if idx >= 0:
            spans.append((idx, idx + 1))
    elif "epígrafe" in rule_lower or "epigrafe" in rule_lower:
        match_dot = raw_text.find(".")
        match_colon = raw_text.find(":")
        candidates = [i for i in [match_dot, match_colon] if i >= 0]
        if candidates:
            idx = min(candidates)
            spans.append((0, idx + 1))
    elif "expressão composta" in rule_lower or "expressao composta" in rule_lower:
        idx = raw_text.find(evidence_text)
        if idx >= 0:
            spans.append((idx, idx + len(evidence_text)))
    elif "artigo sem itálico" in rule_lower:
        for m in re.finditer(r"\b(o|a|os|as)\b", raw_text):
            spans.append((m.start(), m.end()))
    elif "separador sem itálico" in rule_lower:
        for idx, char in enumerate(raw_text):
            if char == ";":
                spans.append((idx, idx + 1))
    elif "título em negrito" in rule_lower or "capitalização do título" in rule_lower:
        m = re.match(r"^\s*(\d{1,3})\.?\s+(.+)$", raw_text)
        if m:
            num_end = m.end()
            colon_pos = raw_text[num_end:].find(":")
            if colon_pos != -1:
                spans.append((num_end, num_end + colon_pos))
    elif "capitalização dos itens secundários" in rule_lower or "estilo dos itens secundários" in rule_lower:
        m = re.search(r"O item '([^']+)'", suggestion)
        if not m:
            m = re.search(r"O item secundário '([^']+)'", suggestion)
        if m:
            target_item = m.group(1)
            idx = raw_text.find(target_item)
            if idx >= 0:
                spans.append((idx, idx + len(target_item)))
    elif "espaçamento duplo" in rule_lower:
        for m in re.finditer(r"(?<=\S) (?=\S)", raw_text):
            spans.append((m.start(), m.end()))
    elif "dois espaços após número" in rule_lower:
        m = re.search(r"(\d{1,2}\.) (\S)", raw_text)
        if m:
            spans.append((m.start(2) - 1, m.start(2)))
    elif "padrão zero" in rule_lower:
        m = re.match(r"^\s*\d{1,3}\.?\s*", raw_text)
        if m:
            spans.append((m.start(), m.end()))
    elif "sufixo repetido" in rule_lower:
        m = re.search(r"sufixo '([^']+)'", suggestion)
        if m:
            target_suf = m.group(1)
            for sm in re.finditer(r"\b\w*" + re.escape(target_suf) + r"\b", raw_text):
                spans.append((sm.start(), sm.end()))
    elif "ortopensatologia" in rule_lower or "fórmula da ortopensatologia" in rule_lower:
        for idx, char in enumerate(raw_text):
            if char in "“”\"":
                spans.append((idx, idx + 1))
    elif "interlocução" in rule_lower:
        for m in re.finditer(r"leitor ou leitora", raw_text, re.I):
            spans.append((m.start(), m.end()))
    elif "parasitas" in rule_lower:
        for word in evidence_text.split(","):
            word = word.strip()
            if word:
                for m in re.finditer(r"\b" + re.escape(word) + r"\b", raw_text, re.I):
                    spans.append((m.start(), m.end()))
                    
    return spans


def build_rich_docx_views(path: Path, result: dict | None = None) -> tuple[dict[str, str], dict[str, str]]:
    try:
        rich_paragraphs = read_docx_rich(path)
    except Exception:
        return {}, {}

    if result:
        ajustes = result.get("formatacao", {}).get("ajustes", []) + result.get("achados", [])
        for item in list(ajustes):
            p_rich = find_rich_paragraph(rich_paragraphs, item.get("evidencia", ""))
            if p_rich:
                raw_text = p_rich.get("raw_text", p_rich["text"])
                spans = target_spans(item.get("regra", ""), raw_text, item.get("sugestao", ""), item.get("evidencia", ""))
                for start, end in spans:
                    highlight_runs_span(p_rich, start, end)

    current_div = None
    current_sec = None
    sections: dict[str, list[str]] = {}
    evidence: dict[str, str] = {}
    frase_enfatica_detectada = False

    for i, rich in enumerate(rich_paragraphs):
        text = rich["text"]
        html = rich_runs_to_html(rich["runs"])
        evidence.setdefault(clean(text), html)

        div = detect_division(text)
        if div:
            current_div = div
            current_sec = None
            continue

        if current_div == "Argumentologia":
            if starts_with_argumentologia_epigraph(text):
                idx = len([k for k in sections if k.startswith("Argumentologia_")])
                current_sec = f"Argumentologia_{idx}"
                sections.setdefault(current_sec, []).append(html)
            else:
                if not current_sec:
                    idx = len([k for k in sections if k.startswith("Argumentologia_")])
                    current_sec = f"Argumentologia_{idx}"
                    sections.setdefault(current_sec, []).append(html)
                else:
                    sections[current_sec].append(html)
            continue

        sec = detect_independent_epigraph(text)
        if sec:
            current_sec = sec
            sections.setdefault(sec, []).append(html)
            continue

        next_sec = detect_independent_epigraph(rich_paragraphs[i + 1]["text"]) if i + 1 < len(rich_paragraphs) else None
        if current_sec == "Remissiologia" and next_sec == "Questionologia" and len(text) > 40 and not frase_enfatica_detectada:
            current_sec = "Frase Enfatica"
            frase_enfatica_detectada = True
            sections.setdefault(current_sec, []).append(html)
            continue

        if current_sec:
            sections.setdefault(current_sec, []).append(html)

    return {sec: "<br><br>".join(parts) for sec, parts in sections.items()}, evidence


def evidence_html(text: str, rich_evidence: dict[str, str]) -> str:
    plain = clean(text)
    if plain in rich_evidence:
        return rich_evidence[plain]
    for key, html in rich_evidence.items():
        if plain and (plain in key or key in plain):
            return html
    return escape(text)


def highlight_evidence(text: str, items: list[dict], rich_evidence: dict[str, str]) -> str:
    html_content = evidence_html(text, rich_evidence)
    if html_content == escape(text):
        import re
        def wrap_text(t):
            return f'<span style="background-color: #ffeba0; color: #333333; padding: 1px 3px; border-radius: 3px;">{t}</span>'
            
        for item in items:
            rule_lower = item["regra"].lower()
            suggestion = item["sugestao"]
            ev_text = item["evidencia"]
            
            if "ponto final" in rule_lower:
                html_content = re.sub(r"\.(?=\s*$)", wrap_text('.'), html_content)
            elif "dois-pontos" in rule_lower or "dois pontos" in rule_lower:
                html_content = re.sub(r":(?=[^:]*$)", wrap_text(':'), html_content)
            elif "epígrafe" in rule_lower or "epigrafe" in rule_lower:
                html_content = re.sub(r"^([^:.]+[:.])", lambda m: wrap_text(m.group(1)), html_content)
            elif "expressão composta" in rule_lower or "expressao composta" in rule_lower:
                html_content = re.sub(re.escape(escape(ev_text)), lambda m: wrap_text(m.group(0)), html_content, flags=re.IGNORECASE)
            elif "artigo sem itálico" in rule_lower:
                html_content = re.sub(r"\b(o|a|os|as)\b", lambda m: wrap_text(m.group(0)), html_content)
            elif "separador sem itálico" in rule_lower:
                html_content = re.sub(r";", wrap_text(';'), html_content)
            elif "título em negrito" in rule_lower or "capitalização do título" in rule_lower:
                html_content = re.sub(r"(\d{1,3}\.\s+)([^:]+)(:)", lambda m: m.group(1) + wrap_text(m.group(2)) + m.group(3), html_content)
            elif "capitalização dos itens secundários" in rule_lower or "estilo dos itens secundários" in rule_lower:
                match = re.search(r"O item '([^']+)'", suggestion)
                if not match:
                    match = re.search(r"O item secundário '([^']+)'", suggestion)
                if match:
                    target_item = match.group(1)
                    html_content = re.sub(re.escape(escape(target_item)), lambda m: wrap_text(m.group(0)), html_content)
            elif "espaçamento duplo" in rule_lower:
                html_content = re.sub(r"(?<=\S) (?=\S)", wrap_text(' '), html_content)
            elif "dois espaços após número" in rule_lower:
                html_content = re.sub(r"(\d{1,2}\.) (\S)", lambda m: m.group(1) + wrap_text(' ') + m.group(2), html_content)
            elif "padrão zero" in rule_lower:
                html_content = re.sub(r"(^\s*\d{1,3}\.?\s*)", lambda m: wrap_text(m.group(1)), html_content)
            elif "sufixo repetido" in rule_lower:
                match = re.search(r"sufixo '([^']+)'", suggestion)
                if match:
                    target_suf = match.group(1)
                    html_content = re.sub(r"\b\w*" + re.escape(escape(target_suf)) + r"\b", lambda m: wrap_text(m.group(0)), html_content)
            elif "ortopensatologia" in rule_lower or "fórmula da ortopensatologia" in rule_lower:
                html_content = re.sub(r"[“”\"]", lambda m: wrap_text(m.group(0)), html_content)
            elif "interlocução" in rule_lower:
                html_content = re.sub(r"leitor ou leitora", lambda m: wrap_text(m.group(0)), html_content, flags=re.IGNORECASE)
            elif "parasitas" in rule_lower:
                for word in ev_text.split(","):
                    word = word.strip()
                    if word:
                        html_content = re.sub(r"\b" + re.escape(escape(word)) + r"\b", lambda m: wrap_text(m.group(0)), html_content, flags=re.IGNORECASE)
    return html_content

col_title, col_slider = st.columns([12, 1])
with col_title:
    st.markdown(
        """
        <div class="main-title-shell">
          <h1 class="main-title">Revisor Verbetográfico</h1>
          <p class="main-subtitle">Auditoria formal do verbete e relatório de ajustes necessários</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with col_slider:
    st.write("<div style='height: 38px;'></div>", unsafe_allow_html=True)
    width_pct = st.slider("", min_value=40, max_value=80, value=60, step=5, format="%d%%")

st.markdown(
    f"""
    <style>
    .block-container {{
      max-width: {width_pct}% !important;
    }}
    @media (min-width: 2000px) {{
      .block-container {{
        max-width: {width_pct}% !important;
      }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

uploaded = st.file_uploader(
    "Selecione um verbete",
    type=["docx", "pdf", "doc"],
    help="DOCX é o formato mais fiel. PDF funciona em modo melhor esforço; DOC deve ser convertido para DOCX. Limite local de upload: 1 GB.",
)

if not uploaded:
    #st.markdown('<div class="soft-panel">Envie um arquivo para iniciar a análise. O app gera JSON formal, painel de achados e relatório Word para revisão.</div>', unsafe_allow_html=True)
    st.stop()

if uploaded:
    suffix = Path(uploaded.name).suffix
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        tmp.write(uploaded.getbuffer())
        tmp_path = Path(tmp.name)

    docx_path = None
    conversion_warning = None
    if suffix.lower() == ".doc":
        docx_path = convert_doc_to_docx(tmp_path)
        if not docx_path:
            conversion_warning = "Não foi possível converter o DOC para DOCX neste ambiente (instale o LibreOffice ou o Microsoft Word). A análise seguirá apenas com texto, sem formatação."
    elif suffix.lower() == ".docx":
        docx_path = tmp_path

    try:
        result = analyze_file(tmp_path)
        rich_sections = {}
        rich_evidence = {}
        marked_docx = None
        corrected_docx = None
        marked_pdf = None
        if docx_path is not None:
            rich_sections, rich_evidence = build_rich_docx_views(docx_path, result)
            marked_docx = gerar_verbete_marcado(docx_path, result)
            corrected_docx = gerar_verbete_corrigido(docx_path, result)
        elif suffix.lower() == ".pdf":
            marked_pdf = gerar_pdf_marcado(tmp_path, result)
    finally:
        tmp_path.unlink(missing_ok=True)
        if docx_path is not None and docx_path != tmp_path:
            docx_path.unlink(missing_ok=True)

    if conversion_warning:
        st.warning(conversion_warning)

    report_md = render_markdown(result)
    report_docx = markdown_to_report_docx(report_md)
    achados = result["achados"]
    crit = sum(1 for a in achados if a["severidade"] == "Crítico")
    imp = sum(1 for a in achados if a["severidade"] == "Importante")
    marca = result["marca_excelencia"]
    base_name = Path(uploaded.name).stem

    if marked_docx and corrected_docx:
        download_cols = st.columns(3)
        with download_cols[0]:
            st.download_button(
                "Baixar relatório Word",
                report_docx,
                file_name="relatorio-conformidade.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "Baixar verbete marcado",
                marked_docx,
                file_name=f"{base_name}-marcado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with download_cols[2]:
            st.download_button(
                "Baixar verbete corrigido",
                corrected_docx,
                file_name=f"{base_name}-corrigido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
    elif marked_pdf:
        download_cols = st.columns(2)
        with download_cols[0]:
            st.download_button(
                "Baixar relatório Word",
                report_docx,
                file_name="relatorio-conformidade.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with download_cols[1]:
            st.download_button(
                "Baixar verbete marcado",
                marked_pdf,
                file_name=f"{base_name}-marcado.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
        st.caption("Correção automática final disponível apenas para DOCX.")
    else:
        st.download_button(
            "Baixar relatório Word",
            report_docx,
            file_name="relatorio-conformidade.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.caption("Verbete marcado/corrigido disponível para DOCX. PDF marcado depende de busca textual.")

    st.markdown(
        f"""
        <div class="metric-row">
          <div class="metric-card"><div class="metric-label">Achados críticos</div><div class="metric-value status-danger">{crit}</div></div>
          <div class="metric-card"><div class="metric-label">Importantes</div><div class="metric-value status-warn">{imp}</div></div>
          <div class="metric-card"><div class="metric-label">Máximos</div><div class="metric-value {status_class(marca['maximos']['ok'])}">{marca['maximos']['encontrado']} / {marca['maximos']['exigido']}</div></div>
          <div class="metric-card"><div class="metric-label">Logias</div><div class="metric-value {status_class(marca['logias']['ok'])}">{marca['logias']['encontrado']} / {marca['logias']['exigido']}</div></div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    tab_report, tab_summary, tab_template = st.tabs(
        ["Relatório", "Resumo", "Template"]
    )

    with tab_report:
        st.markdown(report_md)

    with tab_summary:
        if "show_only_inconsistencies" not in st.session_state:
            st.session_state["show_only_inconsistencies"] = False

        col_title, col_btn = st.columns([2, 1])
        with col_title:
            st.subheader("Resumo Integrado de Seções")
        with col_btn:
            st.write("<div style='height: 14px;'></div>", unsafe_allow_html=True)
            with st.container(key="report_toggle_buttons"):
                b1, b2 = st.columns(2)
                with b1:
                    is_active = not st.session_state["show_only_inconsistencies"]
                    if st.button("Completo", type="primary" if is_active else "secondary", key="btn_show_completo", use_container_width=True):
                        st.session_state["show_only_inconsistencies"] = False
                        st.rerun()
                with b2:
                    is_active = st.session_state["show_only_inconsistencies"]
                    if st.button("Inconsistências", type="primary" if is_active else "secondary", key="btn_show_inconsistencies", use_container_width=True):
                        st.session_state["show_only_inconsistencies"] = True
                        st.rerun()

        secoes_texto = result.get("secoes_texto", {})
        contagens = result.get("contagens", {})
        achados = result.get("achados", [])
        
        summary_rows = []
        for sec in result.get("secoes_presentes", []):
            related = [a for a in achados if a["secao"] == sec]
            situacao = "❌" if any(a["severidade"] == "Crítico" for a in related) else "⚠️" if related else "✅"
            
            rules_list = []
            suggs_list = []
            for a in related:
                r_esc = escape(a["regra"])
                s_esc = escape(a["sugestao"])
                if r_esc not in rules_list:
                    rules_list.append(r_esc)
                if s_esc not in suggs_list:
                    suggs_list.append(s_esc)
            
            regra_str = "<br>".join(f"• {r}" for r in rules_list) if len(rules_list) > 1 else (rules_list[0] if rules_list else "")
            sugestao_str = "<br>".join(f"• {s}" for s in suggs_list) if len(suggs_list) > 1 else (suggs_list[0] if suggs_list else "")
            
            sec_html = rich_sections.get(sec)
            if not sec_html:
                sec_items = [a for a in achados if a["secao"] == sec]
                sec_html = highlight_evidence(secoes_texto.get(sec, ""), sec_items, {})
                
            eff_sec = get_effective_section(sec, secoes_texto.get(sec, ""))
            max_rule = SECOES.get(eff_sec, {}).get("max")
            if max_rule:
                reached = any(m["secao"] == sec for m in result.get("maximos", []))
                max_status = f">\xa0{max_rule} ✅" if reached else f"<\xa0{max_rule}"
            else:
                max_status = ""
                
            sec_plain = secoes_texto.get(sec, "")
            terms = []
            skip_body = {"Interdisciplinologia", "Remissiologia", "Frase Enfatica"}
            skip_title = {"Interdisciplinologia", "Frase Enfatica"}
            if eff_sec not in skip_title and re.search(r"logia$", eff_sec, re.I):
                terms.append(eff_sec)
            if eff_sec not in skip_body:
                terms.extend(logia_terms(section_body(sec_plain)))
            logias_count = len(terms)
            
            summary_rows.append({
                "Seção": clean_sec(sec),
                "Situação": situacao,
                "Texto": sec_html,
                "Regra": regra_str,
                "Sugestão": sugestao_str,
                "Itens": contagens.get(sec, ""),
                "Máximos": max_status,
                "Logias": logias_count
            })
            
        if st.session_state["show_only_inconsistencies"]:
            summary_rows = [row for row in summary_rows if row["Situação"] != "✅"]

        render_wrap_table(
            summary_rows,
            ["Seção", "Situação", "Texto", "Regra", "Sugestão", "Itens", "Máximos", "Logias"],
            [12, 8, 30, 16, 18, 6, 6, 4],
            {"Texto", "Regra", "Sugestão"}
        )

        st.subheader("Marca de excelência")
        render_wrap_table(
            [
                {
                    "Indicador": label,
                    "Exigido": item["exigido"],
                    "Encontrado": item["encontrado"] if item["encontrado"] is not None else "verificar",
                    "Situação": status_text(item["ok"]),
                }
                for label, item in [
                    ("Páginas", marca["paginas"]),
                    ("Máximos", marca["maximos"]),
                    ("Logias", marca["logias"]),
                ]
            ],
            ["Indicador", "Exigido", "Encontrado", "Situação"],
            [25, 20, 35, 20],
        )

    with tab_template:
        st.subheader("Diretrizes de Formatação (Template)")
        st.markdown(
            r"""
            Esta tabela apresenta as diretrizes oficiais de fontes, tamanhos, espaçamentos e layout para orientação e verificação manual:

            | Elemento / Layout | Fonte | Tamanho | Estilo / Realce | Espaçamento / Alinhamento / Detalhes |
            | :--- | :--- | :--- | :--- | :--- |
            | **Entrada (Título)** | Arial | 11 | Negrito-Itálico, Versalete | Borda dupla |
            | **Especialidade** | Arial | 11 | Versalete | Espaçamento de caractere expandido em 1,5 pt |
            | **Texto Geral (Corpo)** | Times New Roman | 10 | Regular | Espaçamento simples |
            | **Cabeçalho** | Times New Roman | 9 | Itálico | Centralizado ("Enciclopédia da Conscienciologia") |
            | **Paginação** | Times New Roman | 9 | Regular | Canto superior direito |
            | **Ficha Técnica (Filmografia)** | Times New Roman | 8 | Regular | Alinhamento à esquerda |
            | **Papel (Tamanho)** | Carta (Letter) | - | - | 21,59 cm x 27,94 cm (8.5" x 11") |
            | **Margens** | - | - | - | Superior: 3,0 cm \| Inferior: 3,07 cm \| Esquerda: 3,77 cm \| Direita: 3,75 cm |
            """
        )
