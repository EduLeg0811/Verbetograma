import { ChangeEvent, ComponentPropsWithoutRef, DragEvent, useEffect, useMemo, useState } from "react";
import { AlertCircle, Check, ChevronRight, CircleX, Download, FileText, Lightbulb, LoaderCircle, Moon, RotateCcw, Search, ShieldCheck, Sparkles, Sun, TriangleAlert, UploadCloud, X } from "lucide-react";
import Markdown from "react-markdown";
import remarkGfm from "remark-gfm";
import favicon from "../../favicon.svg";

type Finding = { severidade: string; secao: string; regra: string; evidencia: string; sugestao: string };
type Mark = { exigido: number; encontrado: number | null; ok: boolean | null };
type Result = {
  meta: Record<string, string | number | null>;
  marca_excelencia: Record<string, Mark>;
  secoes_presentes: string[];
  secoes_texto: Record<string, string>;
  contagens: Record<string, number>;
  achados: Finding[];
};
type SummaryRow = { section: string; key: string; html: string; count: number | null; maximum: string; logias: number; findings: Finding[] };
type Payload = { filename: string; warning?: string; note?: string; result: Result; summary: SummaryRow[]; report: string; downloads: { name: string; url: string }[] };
type Tab = "report" | "summary" | "template";

const severity = (value: string) => value === "Crítico" ? "critical" : value === "Importante" ? "warning" : "refine";
const label = (key: string) => ({ paginas: "Páginas", maximos: "Máximos", logias: "Logias" }[key] || key);
const escapeHtml = (value: string) => value.replace(/[&<>"']/g, (char) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[char]!));

function StatusIcon({ status }: { status: "ok" | "warning" | "critical" | "refine" }) {
  const label = { ok: "Conforme", warning: "Atenção", critical: "Crítico", refine: "Refinamento" }[status];
  return <span className={`status-icon ${status}`} aria-label={label} title={label}>
    {status === "ok" ? <Check /> : status === "warning" ? <TriangleAlert /> : status === "critical" ? <CircleX /> : <Lightbulb />}
  </span>;
}

function ReportCell({ children, ...props }: ComponentPropsWithoutRef<"td">) {
  const text = (Array.isArray(children) ? children.join("") : String(children ?? "")).trim();
  const status = text.includes("✅") ? "ok" : text.includes("❌") ? "critical" : text.includes("⚠") ? "warning" : text.includes("💡") ? "refine" : null;
  return <td {...props}>{status ? <StatusIcon status={status} /> : children}</td>;
}

function ThemeButton() {
  const [theme, setTheme] = useState(() => localStorage.getItem("app-theme") || "light");
  useEffect(() => {
    document.documentElement.classList.toggle("dark", theme === "dark");
    localStorage.setItem("app-theme", theme);
  }, [theme]);
  return <button className="icon-button group" onClick={() => setTheme(theme === "light" ? "dark" : "light")} aria-label="Alternar tema">
    {theme === "light" ? <Moon className="group-hover:-rotate-12" /> : <Sun className="group-hover:rotate-45" />}
  </button>;
}

function Header() {
  return <header className="topbar"><div className="header-inner">
    <a className="brand" href="/"><img className="brand-mark" src={favicon} alt="" /><span className="brand-name">Verbeto<em>grama</em></span><span className="brand-category">Análise técnica</span></a>
    <ThemeButton />
  </div></header>;
}

function Upload({ busy, onFile }: { busy: boolean; onFile: (file: File) => void }) {
  const [drag, setDrag] = useState(false);
  const pick = (event: ChangeEvent<HTMLInputElement>) => event.target.files?.[0] && onFile(event.target.files[0]);
  const drop = (event: DragEvent) => { event.preventDefault(); setDrag(false); event.dataTransfer.files[0] && onFile(event.dataTransfer.files[0]); };
  return <label className={`dropzone ${drag ? "dragging" : ""} ${busy ? "busy" : ""}`} onDragOver={(e) => { e.preventDefault(); setDrag(true); }} onDragLeave={() => setDrag(false)} onDrop={drop}>
    <input type="file" accept=".doc,.docx,.pdf" onChange={pick} disabled={busy} />
    <span className="upload-icon">{busy ? <LoaderCircle className="spin" /> : <UploadCloud />}</span>
    <span className="upload-title">{busy ? "Auditando o verbete…" : "Arraste o verbete para iniciar"}</span>
    <span className="upload-copy">{busy ? "Estrutura, confor e marca de excelência" : <>ou <strong>selecione no computador</strong> · DOCX, DOC ou PDF</>}</span>
    <span className="format-note">DOCX preserva a auditoria tipográfica completa</span>
  </label>;
}

function Metrics({ data }: { data: Result }) {
  const critical = data.achados.filter((x) => x.severidade === "Crítico").length;
  const important = data.achados.filter((x) => x.severidade === "Importante").length;
  const marks = data.marca_excelencia;
  return <div className="metrics">
    <article><span>Achados críticos</span><strong className="critical-text">{critical}</strong></article>
    <article><span>Importantes</span><strong className="warning-text">{important}</strong></article>
    <article><span>Máximos</span><strong>{marks.maximos.encontrado} <i>/ {marks.maximos.exigido}</i></strong></article>
    <article><span>Logias</span><strong>{marks.logias.encontrado} <i>/ {marks.logias.exigido}</i></strong></article>
  </div>;
}

function Summary({ data, summary = [] }: { data: Result; summary?: SummaryRow[] }) {
  const [onlyIssues, setOnlyIssues] = useState(false);
  const [query, setQuery] = useState("");
  const rows = useMemo(() => (summary.length ? summary : data.secoes_presentes.map((key) => ({
    section: key.replace(/_/g, " · "), key, html: escapeHtml(data.secoes_texto[key] || ""), count: data.contagens[key] ?? null,
    maximum: "", logias: 0, findings: data.achados.filter((item) => item.secao === key),
  }))).filter((row) =>
    (!onlyIssues || row.findings.length) && `${row.section} ${data.secoes_texto[row.key] || ""}`.toLowerCase().includes(query.toLowerCase())
  ), [data, summary, onlyIssues, query]);

  return <div className="summary">
    <div className="section-tools"><div><h2>Resumo integrado</h2><p>{rows.length} seções em exibição</p></div><div className="tool-actions">
      <label className="search"><Search /><input value={query} onChange={(e) => setQuery(e.target.value)} placeholder="Buscar seção" /></label>
      <div className="segmented"><button className={!onlyIssues ? "active" : ""} onClick={() => setOnlyIssues(false)}>Completo</button><button className={onlyIssues ? "active" : ""} onClick={() => setOnlyIssues(true)}>Inconsistências</button></div>
    </div></div>
    <div className="section-list">{rows.map((row) => {
      const status = row.findings.some((item) => item.severidade === "Crítico") ? "critical" : row.findings.length ? "warning" : "ok"; return <details key={row.key} className={status} open={!onlyIssues && row.findings.length > 0}>
        <summary><StatusIcon status={status} /><span className="summary-title"><strong>{row.section}</strong>{row.maximum.startsWith(">=") && <span className="maximum-pill">Máximo</span>}</span><span className="summary-count">Itens: {row.count ?? "—"} · Máx.: {row.maximum || "—"} · Logias: {row.logias}</span><span>{row.findings.length ? `${row.findings.length} ajuste${row.findings.length > 1 ? "s" : ""}` : "Conforme"}</span><ChevronRight /></summary>
        <div className="detail-body"><div className="section-text" dangerouslySetInnerHTML={{ __html: row.html }} />{row.findings.map((item, index) => <article className="finding" key={`${item.regra}-${index}`}><span className={`badge ${severity(item.severidade)}`}>{item.severidade}</span><h3>{item.regra}</h3><blockquote>{item.evidencia}</blockquote><p>{item.sugestao}</p></article>)}</div>
      </details>;
    })}</div>
    <h2 className="mark-title">Marca de excelência</h2><div className="mark-grid">{Object.entries(data.marca_excelencia).map(([key, item]) => <article key={key}><span>{label(key)}</span><strong>{item.encontrado ?? "Verificar"}</strong><small>exigido: {item.exigido}</small><i className={item.ok ? "pass" : item.ok === false ? "fail" : "pending"}>{item.ok ? "Conforme" : item.ok === false ? "Ajustar" : "Verificar"}</i></article>)}</div>
  </div>;
}

const templateRows = [
  ["Entrada (título)", "Arial", "11", "Negrito-itálico, versalete", "Borda dupla"],
  ["Especialidade", "Arial", "11", "Versalete", "Caracteres expandidos em 1,5 pt"],
  ["Texto geral", "Times New Roman", "10", "Regular", "Espaçamento simples"],
  ["Cabeçalho", "Times New Roman", "9", "Itálico", "Centralizado"],
  ["Paginação", "Times New Roman", "9", "Regular", "Canto superior direito"],
  ["Ficha técnica", "Times New Roman", "8", "Regular", "Alinhamento à esquerda"],
  ["Papel", "Carta", "—", "—", "21,59 × 27,94 cm"],
  ["Margens", "—", "—", "—", "Sup. 3,0 · Inf. 3,07 · Esq. 3,77 · Dir. 3,75 cm"],
];

function Template() { return <div className="template"><h2>Diretrizes de formatação</h2><p>Referência rápida para conferência manual do documento.</p><div className="table-wrap"><table><thead><tr>{["Elemento", "Fonte", "Tamanho", "Estilo", "Detalhes"].map((x) => <th key={x}>{x}</th>)}</tr></thead><tbody>{templateRows.map((row) => <tr key={row[0]}>{row.map((cell, i) => <td key={i}>{cell}</td>)}</tr>)}</tbody></table></div></div>; }

function Results({ payload, reset }: { payload: Payload; reset: () => void }) {
  const [tab, setTab] = useState<Tab>("report");
  const title = String(payload.result.meta.titulo || payload.filename);
  return <main className="results-page">
    <section className="result-heading"><div><span className="eyebrow"><ShieldCheck /> Auditoria concluída</span><h1>{title}</h1><p>{payload.filename}</p></div><button className="secondary-button" onClick={reset}><RotateCcw /> Analisar outro</button></section>
    {payload.warning && <div className="notice"><AlertCircle />{payload.warning}</div>}
    <Metrics data={payload.result} />
    <section className="download-bar"><div><span>Documentos gerados</span><small>{payload.note || `${payload.downloads.length} arquivos prontos`}</small></div><div>{payload.downloads.map((file) => <a href={file.url} key={file.name}><Download />{file.name.replace("relatorio-conformidade", "Relatório").replace("resultado-formal", "Dados formais")}</a>)}</div></section>
    <nav className="tabs" aria-label="Resultados">{(["report", "summary", "template"] as Tab[]).map((key) => <button className={tab === key ? "active" : ""} onClick={() => setTab(key)} key={key}>{key === "report" ? "Relatório" : key === "summary" ? "Resumo" : "Template"}</button>)}</nav>
    <section className="panel">{tab === "report" ? <article className="markdown"><Markdown remarkPlugins={[remarkGfm]} components={{ td: ReportCell }}>{payload.report}</Markdown></article> : tab === "summary" ? <Summary data={payload.result} summary={payload.summary ?? []} /> : <Template />}</section>
  </main>;
}

export default function App() {
  const [busy, setBusy] = useState(false);
  const [payload, setPayload] = useState<Payload | null>(null);
  const [error, setError] = useState("");
  const analyze = async (file: File) => {
    setError("");
    if (!/\.(docx?|pdf)$/i.test(file.name)) return setError("Selecione um arquivo DOC, DOCX ou PDF.");
    setBusy(true);
    try {
      const form = new FormData(); form.append("upload", file);
      const response = await fetch("/api/analyze", { method: "POST", body: form });
      const body = await response.json();
      if (!response.ok) throw new Error(body.detail || "Falha na análise.");
      setPayload(body);
    } catch (err) { setError(err instanceof Error ? err.message : "Falha na análise."); }
    finally { setBusy(false); }
  };
  return <><Header />{payload ? <Results payload={payload} reset={() => setPayload(null)} /> : <main className="landing">
    <section className="hero"><span className="hero-kicker"><Sparkles /></span><h1>Auditoria<br /><em>Verbetográfica.</em></h1><p>Estrutura • Confor • Máximos • Logias • Relatório técnico</p></section>
    <section className="workbench"><Upload busy={busy} onFile={analyze} />{error && <div className="error"><X />{error}</div>}</section>
    <footer><span><FileText /> @2026 ● Cons-IA</span></footer>
  </main>}</>;
}
