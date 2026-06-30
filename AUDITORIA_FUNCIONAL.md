# Auditoria funcional da migração

Comparação entre `originais/app.py`, as rotinas em `originais/` e o projeto FastAPI + React. Revisão executada em 30 de junho de 2026.

## 1. Integridade do motor Python

Os arquivos atuais abaixo possuem o mesmo SHA-256 dos respectivos arquivos do backup:

- [x] `analisar_verbete.py`: análise estrutural, textual e DOCX.
- [x] `catalogo.py`: catálogo de seções, mínimos, máximos e regras.
- [x] `editar_verbete.py`: DOCX marcado, DOCX corrigido e PDF marcado.
- [x] `relatorio_word.py`: relatório Word.
- [x] `inspecionar_confor.py`: inspeção auxiliar.

Conclusão: nenhuma regra determinística foi removida ou modificada durante a troca de interface.

## 2. Workflow original versus atual

| Etapa ou função | Situação | Implementação atual |
|---|---|---|
| Upload DOCX, DOC e PDF | Equivalente | `POST /api/analyze` |
| Validação da extensão | Equivalente | Resposta HTTP 415 para formato inválido |
| Arquivo temporário local | Equivalente | Diretório de job em `%TEMP%/verbetograma` |
| Conversão DOC para DOCX | Equivalente | `convert_doc_to_docx` |
| Aviso quando DOC não converte | Equivalente | Campo `warning` exibido na interface |
| Análise formal | Idêntica | `analyze_file` original |
| Leitura dos runs DOCX | Restaurada | `scripts/rich_views.py` |
| Negrito, itálico e negrito-itálico no resumo | Restaurada | HTML seguro por run |
| Destaque dos trechos com ajuste | Restaurada | Mesmas regras de localização do app original |
| Métricas críticas, importantes, máximos e logias | Equivalente | Cards de resultado |
| Relatório Markdown | Equivalente | Aba Relatório |
| Resumo completo ou só inconsistências | Equivalente | Estado React; inclui busca adicional |
| Itens, máximo e logias por seção | Restaurada | Cabeçalho de cada seção |
| Marca de excelência | Equivalente | Páginas, máximos e logias |
| Template de formatação | Equivalente | Aba Template |
| Relatório Word | Equivalente | Download DOCX |
| Verbete DOCX marcado | Equivalente | Download DOCX |
| Verbete DOCX corrigido | Equivalente | Download DOCX |
| PDF marcado | Equivalente | Download quando gerado |
| Aviso de correção exclusiva para DOCX | Restaurada | Nota na área de downloads |
| JSON formal | Adição | Download que não existia na tela original |
| Slider de largura do Streamlit | Substituído | Layout responsivo; não interfere na análise ou nos dados |
| Estado de sessão do Streamlit | Substituído | Estado local React, sem sessão de upload |

## 3. Checklist normativo de 30 itens

O checklist continua dividido conforme a responsabilidade original:

- [x] Itens 1–2 — Remissiologia/Acabativa: verificações formais Python e revisão editorial quando necessária.
- [x] Itens 3–17 — forma e confor: regras determinísticas e auditoria dos runs/XML DOCX.
- [x] Itens 18–20 — correlações internas: verificações objetivas no Python; equivalência conceitual permanece editorial.
- [x] Itens 21–30 — leitura corrida: heurísticas objetivas disponíveis e análise semântica reservada ao Codex conforme documentação.

Limites já existentes no projeto original permanecem documentados em `REGRAS.md`: viúvas, quebras de linha, herança completa de estilos do Word e critérios semânticos não são apresentados como verificações determinísticas completas.

## 4. Critérios de regressão

- [x] Quantidade e conteúdo do JSON formal não mudam.
- [x] Todas as seções analisadas possuem uma linha correspondente no resumo.
- [x] DOCX preserva os estilos dos runs na visualização.
- [x] Texto enviado pelo documento é escapado antes da renderização HTML.
- [x] Downloads são produzidos antes de o arquivo-fonte temporário ser removido.
- [x] PDF e DOC mantêm os avisos e limitações do fluxo original.
