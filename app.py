from __future__ import annotations

from html import escape
import json
import tempfile
from pathlib import Path

import streamlit as st

from scripts.analisar_verbete import analyze_file, clean, detect_division, detect_independent_epigraph, detect_section, read_docx_rich, render_markdown
from scripts.editar_verbete import gerar_pdf_marcado, gerar_verbete_corrigido, gerar_verbete_marcado
from scripts.relatorio_word import markdown_to_report_docx

ROOT = Path(__file__).parent
REFERENCES = ROOT / "references"

st.set_page_config(
    page_title="Revisor Verbetográfico",
    page_icon="",
    layout="wide",
    initial_sidebar_state="expanded",
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
      --warn: #a9672b;
      --danger: #a54845;
    }
    .stApp { background: var(--bg); color: var(--ink); }
    [data-testid="stSidebar"] { background: #eee7dc; border-right: 1px solid var(--line); }
    .block-container { padding-top: 2.2rem; max-width: 1280px; }
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
    .stButton > button, .stDownloadButton > button {
      border-radius: 7px;
      border: 1px solid #315d53;
      background: var(--accent);
      color: white;
      font-weight: 650;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
      border-color: #244d45;
      background: #315d53;
      color: white;
    }
    .json-download-panel {
      border-top: 1px solid #d4c8b8;
      margin-top: 18px;
      padding-top: 14px;
    }
    .sidebar-title {
      display: block;
      margin: 1.35rem 0 .25rem 0;
      padding: 0;
      color: #23211f;
      font-size: 1.22rem;
      line-height: 1.45;
      font-weight: 750;
      letter-spacing: 0;
      white-space: normal;
      overflow: visible;
    }
    [data-testid="stSidebar"] .stDownloadButton:last-of-type > button,
    [data-testid="stSidebar"] .stDownloadButton:last-of-type button {
      background: #6f5b40;
      border-color: #5b4933;
    }
    [data-testid="stSidebar"] .stDownloadButton:last-of-type > button:hover,
    [data-testid="stSidebar"] .stDownloadButton:last-of-type button:hover {
      background: #5b4933;
      border-color: #493a29;
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
    </style>
    """,
    unsafe_allow_html=True,
)


def read_reference(name: str) -> str:
    return (REFERENCES / name).read_text(encoding="utf-8")


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
        if run.get("bold"):
            text = f"<strong>{text}</strong>"
        if run.get("italic"):
            text = f"<em>{text}</em>"
        parts.append(text)
    return "".join(parts)


def build_rich_docx_views(path: Path) -> tuple[dict[str, str], dict[str, str]]:
    try:
        rich_paragraphs = read_docx_rich(path)
    except Exception:
        return {}, {}

    current_sec = None
    sections: dict[str, list[str]] = {}
    evidence: dict[str, str] = {}
    frase_enfatica_detectada = False

    for i, rich in enumerate(rich_paragraphs):
        text = rich["text"]
        html = rich_runs_to_html(rich["runs"])
        evidence.setdefault(clean(text), html)

        if detect_division(text):
            current_sec = None
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

with st.sidebar:
    st.markdown('<span class="sidebar-title">Revisor Verbetográfico</span>', unsafe_allow_html=True)
    st.divider()
    st.markdown("#### Referências")
    ref_files = sorted(p.name for p in REFERENCES.glob("*.md"))
    selected_ref = st.selectbox("Abrir referência", ref_files, index=ref_files.index("estrutura-canonica.md") if "estrutura-canonica.md" in ref_files else 0)
    with st.expander("Ler referência selecionada"):
        st.markdown(read_reference(selected_ref))

st.markdown(
    """
    <div class="main-title-shell">
      
      <h1 class="main-title">Revisor Verbetográfico</h1>
      <p class="main-subtitle">Auditoria formal do verbete e relatório de ajustes necessários</p>
    </div>
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

    try:
        result = analyze_file(tmp_path)
        rich_sections = {}
        rich_evidence = {}
        marked_docx = None
        corrected_docx = None
        marked_pdf = None
        if suffix.lower() == ".docx":
            rich_sections, rich_evidence = build_rich_docx_views(tmp_path)
            marked_docx = gerar_verbete_marcado(tmp_path, result)
            corrected_docx = gerar_verbete_corrigido(tmp_path, result)
        elif suffix.lower() == ".pdf":
            marked_pdf = gerar_pdf_marcado(tmp_path, result)
    finally:
        tmp_path.unlink(missing_ok=True)

    report_md = render_markdown(result)
    report_docx = markdown_to_report_docx(report_md)
    achados = result["achados"]
    crit = sum(1 for a in achados if a["severidade"] == "Crítico")
    imp = sum(1 for a in achados if a["severidade"] == "Importante")
    marca = result["marca_excelencia"]
    base_name = Path(uploaded.name).stem

    with st.sidebar:
        st.divider()
        st.markdown("#### Downloads")
        st.download_button(
            "Baixar relatório Word",
            report_docx,
            file_name="relatorio-conformidade.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        if marked_docx and corrected_docx:
            st.download_button(
                "Baixar verbete marcado",
                marked_docx,
                file_name=f"{base_name}-marcado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
            st.download_button(
                "Baixar verbete corrigido",
                corrected_docx,
                file_name=f"{base_name}-corrigido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        elif marked_pdf:
            st.download_button(
                "Baixar verbete marcado",
                marked_pdf,
                file_name=f"{base_name}-marcado.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.caption("Correção automática final disponível apenas para DOCX.")
        else:
            st.caption("Verbete marcado/corrigido disponível para DOCX. PDF marcado depende de busca textual.")
        st.markdown('<div class="json-download-panel"></div>', unsafe_allow_html=True)
        st.markdown("#### Dados técnicos")
        st.download_button(
            "Baixar JSON",
            json.dumps(result, ensure_ascii=False, indent=2),
            file_name="analise-verbetografica.json",
            mime="application/json",
            use_container_width=True,
            type="secondary",
        )

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

    tab_report, tab_text, tab_findings = st.tabs(
        ["Relatório", "Texto", "Achados"]
    )

    with tab_report:
        st.markdown(report_md)

    with tab_text:
        st.subheader("Texto por seção")
        secoes_texto = result.get("secoes_texto", {})
        contagens = result.get("contagens", {})
        rows = []
        for sec in result.get("secoes_presentes", []):
            related = [a for a in achados if a["secao"] == sec]
            situacao = "❌" if any(a["severidade"] == "Crítico" for a in related) else "⚠️" if related else "✅"
            correcao = "\n".join(a["sugestao"] for a in related) if related else ""
            rows.append(
                {
                    "Seção": sec,
                    "Texto": rich_sections.get(sec, escape(secoes_texto.get(sec, ""))),
                    "Situação": situacao,
                    "Correção": correcao,
                    "Itens": contagens.get(sec, ""),
                }
            )
        render_wrap_table(rows, ["Seção", "Texto", "Situação", "Correção", "Itens"], [14, 46, 10, 22, 8], {"Texto"})

    with tab_findings:
        st.subheader("Achados priorizados")
        if achados:
            render_wrap_table(
                [
                    {
                        "Severidade": a["severidade"],
                        "Seção": a["secao"],
                        "Regra": a["regra"],
                        "Evidência": evidence_html(a["evidencia"], rich_evidence),
                        "Sugestão": a["sugestao"],
                    }
                    for a in achados
                ],
                ["Severidade", "Seção", "Regra", "Evidência", "Sugestão"],
                [12, 14, 18, 28, 28],
                {"Evidência"},
            )
        else:
            st.success("Nenhum achado formal foi emitido nesta execução.")

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
