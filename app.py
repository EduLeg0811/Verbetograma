from __future__ import annotations

import json
import shutil
import tempfile
import uuid
from pathlib import Path

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from scripts.analisar_verbete import analyze_file, convert_doc_to_docx, render_markdown
from scripts.editar_verbete import gerar_pdf_marcado, gerar_verbete_corrigido, gerar_verbete_marcado
from scripts.relatorio_word import markdown_to_report_docx
from scripts.rich_views import build_rich_docx_views, build_section_summary

ROOT = Path(__file__).parent
WORK = Path(tempfile.gettempdir()) / "verbetograma"
ALLOWED = {".doc", ".docx", ".pdf"}
WORK.mkdir(exist_ok=True)

app = FastAPI(title="Verbetograma", version="1.0.0")


def save_output(folder: Path, name: str, data: bytes, mime: str) -> dict:
    path = folder / name
    path.write_bytes(data)
    return {"name": name, "url": f"/api/download/{folder.name}/{name}", "mime": mime}


@app.get("/api/health")
def health() -> dict:
    return {"status": "ok"}


@app.get("/favicon.svg", include_in_schema=False)
def favicon() -> FileResponse:
    return FileResponse(ROOT / "favicon.svg", media_type="image/svg+xml")


@app.post("/api/analyze")
async def analyze(upload: UploadFile = File(...)) -> dict:
    filename = Path(upload.filename or "verbete").name
    suffix = Path(filename).suffix.lower()
    if suffix not in ALLOWED:
        raise HTTPException(415, "Use um arquivo DOC, DOCX ou PDF.")

    job = WORK / uuid.uuid4().hex
    job.mkdir()
    source = job / filename
    source.write_bytes(await upload.read())
    docx_path = source if suffix == ".docx" else convert_doc_to_docx(source) if suffix == ".doc" else None
    warning = None
    if suffix == ".doc" and not docx_path:
        warning = "Não foi possível converter o DOC. A análise seguirá sem auditoria completa de formatação."

    try:
        result = analyze_file(source)
        rich_sections = build_rich_docx_views(docx_path, result) if docx_path else {}
        summary = build_section_summary(result, rich_sections)
        report = render_markdown(result)
        downloads = [save_output(job, "relatorio-conformidade.docx", markdown_to_report_docx(report), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
        downloads.append(save_output(job, "resultado-formal.json", json.dumps(result, ensure_ascii=False, indent=2).encode(), "application/json"))
        base = Path(filename).stem
        if docx_path:
            downloads.extend([
                save_output(job, f"{base}-marcado.docx", gerar_verbete_marcado(docx_path, result), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
                save_output(job, f"{base}-corrigido.docx", gerar_verbete_corrigido(docx_path, result), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
            ])
        elif suffix == ".pdf":
            marked_pdf = gerar_pdf_marcado(source, result)
            if marked_pdf:
                downloads.append(save_output(job, f"{base}-marcado.pdf", marked_pdf, "application/pdf"))
    except Exception as exc:
        shutil.rmtree(job, ignore_errors=True)
        raise HTTPException(422, f"Não foi possível analisar o arquivo: {exc}") from exc
    finally:
        if docx_path and docx_path != source and docx_path.exists():
            docx_path.unlink()

    source.unlink(missing_ok=True)
    note = "Correção automática final disponível apenas para DOCX." if suffix == ".pdf" else None
    return {"filename": filename, "warning": warning, "note": note, "result": result, "summary": summary, "report": report, "downloads": downloads}


@app.get("/api/download/{job}/{filename}")
def download(job: str, filename: str) -> FileResponse:
    path = WORK / Path(job).name / Path(filename).name
    if not path.is_file():
        raise HTTPException(404, "Arquivo não encontrado.")
    return FileResponse(path, filename=path.name)


DIST = ROOT / "frontend" / "dist"
if DIST.exists():
    app.mount("/", StaticFiles(directory=DIST, html=True), name="frontend")
