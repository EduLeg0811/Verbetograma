# Verbetograma

Aplicação local para revisão formal de verbetes da Enciclopédia da Conscienciologia. A interface usa React, TypeScript e Tailwind; a API FastAPI preserva o motor determinístico Python.

## Execução

No Windows, execute:

```powershell
.\run-app.cmd
```

O launcher instala dependências globalmente, encerra instâncias antigas, inicia FastAPI e Vite e abre `http://127.0.0.1:5173` no navegador. Não utiliza ambiente virtual.

Para desenvolvimento:

```powershell
python -m uvicorn app:app --reload --port 8765
npm --prefix frontend run dev
```

O Vite atende em `http://127.0.0.1:5173` e encaminha `/api` ao backend na porta 8765.

## Arquitetura

- `app.py`: API, upload, análise e downloads.
- `frontend/`: SPA React + TypeScript + Tailwind.
- `scripts/analisar_verbete.py`: análise formal e CLI.
- `scripts/editar_verbete.py`: documentos marcados e corrigidos.
- `scripts/relatorio_word.py`: relatório Word.
- `references/`: normas verificadas pelo revisor humano e pelo Codex.
- `prompts/`: análise semântica e integração do relatório.

## Fluxo

1. Enviar `.docx`, `.doc` ou `.pdf` pela interface.
2. Conferir métricas, relatório e resumo integrado.
3. Baixar JSON, relatório e documento marcado/corrigido quando suportado.
4. Para revisão semântica, usar o JSON com `prompts/codex-analise-semantica.md`.

O `.docx` é o formato mais fiel. PDF opera em melhor esforço; DOC depende de LibreOffice ou Microsoft Word para conversão local.
