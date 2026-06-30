---
name: revisor-verbete-conscienciologia
description: "Use este skill sempre que o usuário enviar um documento Word (.doc ou .docx) que seja um verbete da Enciclopédia da Conscienciologia (neoverbete) e quiser revisá-lo, avaliá-lo, conferi-lo ou verificar se está conforme as normas verbetográficas — de estrutura, conteúdo, estilística e formatação (confor). Dispare também quando o pedido mencionar 'verbete', 'neoverbete', 'verbetografia', 'Enciclopédia da Conscienciologia', 'revisão verbetográfica', 'checar máximos', 'conferir logias', 'confor', 'Chapa Verbetográfica', 'está dentro das normas?', ou quando o usuário pedir um parecer/relatório de conformidade sobre um texto no estilo conscienciológico. Use mesmo que o usuário não diga explicitamente a palavra 'verbete', desde que o documento tenha a forma de um verbete (Definologia, Etimologia, Especialidade entre parênteses, divisões I–VI, seções terminadas em -logia, Frase Enfática centralizada, etc.). NÃO use para criar um verbete do zero sem texto-base, nem para planilhas."
license: Proprietary
---

# Revisor de Verbetes da Enciclopédia da Conscienciologia

## O que este skill faz

Recebe um verbete e produz um Relatório de Conformidade Verbetográfica combinando três camadas:

1. **Camada formal Python.** `scripts/analisar_verbete.py` faz a verificação objetiva e repetível: estrutura, divisões, seções, mínimos, máximos, logias, páginas, padrão zero, ordem alfabética objetiva, totais da Remissiologia, Enumerologia, Exemplologia formal, espaçamento/separação formal de enumerações, correlações string-matcháveis, parasitas, regras textuais automatizadas e auditoria DOCX de realce/pontuação/layout/estilo quando possível.
2. **Camada semântica Codex/LLM.** Depois do JSON formal, Codex avalia o que exige juízo editorial: critérios qualificadores, pesquisa exaustiva, Definologia, tema central, sinônimos, antônimos, sublinhamentos conceituais, Estrangeirismologia, Exemplologia, Enumerologia semântica, Remissiologia semântica, Ortopensatologia, coerência, Especialidade, Frase Enfática, Questionologia e língua fina.
3. **Camada de integração.** O relatório final funde JSON formal e achados semânticos, sem recalcular o que o Python já calculou.

## Princípio orientador

Forma objetiva é Python; sentido é Codex/LLM. O modelo não deve recontar manualmente máximos, logias, mínimos, padrão zero, ordem alfabética objetiva, totais, regras textuais automatizadas ou achados de formatação. Deve usar o JSON formal como fonte de verdade e acrescentar interpretação editorial.

## Fluxo recomendado

### Passo 1 — Rodar o app Streamlit ou CLI Python

No modo local, use a superfície principal do projeto:

```powershell
pip install -r requirements.txt
streamlit run app.py
```

Envie o verbete pelo upload do app. O `.docx` é o formato mais confiável porque preserva runs de Word usados na auditoria de negrito/itálico/pontuação. `.pdf` funciona em modo melhor esforço. `.doc` depende de conversão local por LibreOffice ou Microsoft Word/COM.

Pela linha de comando:

```powershell
python .\scripts\analisar_verbete.py ".\assets\Taxologia dos Cons.docx" --markdown
python .\scripts\analisar_verbete.py "C:\caminho\para\verbete.docx"
```

O JSON formal traz `meta`, `marca_excelencia`, `secoes_presentes`, `secoes_texto`, `contagens`, `maximos`, `logias`, `logias_lista`, `formatacao` e `achados[]`.

### Passo 2 — Submeter à análise semântica Codex

Use `prompts/codex-analise-semantica.md` com:

- texto integral do verbete;
- JSON formal gerado pelo Python;
- relatório preliminar, se houver.

Não recalcule dados formais. A análise semântica deve produzir apenas achados editoriais com evidência textual.

### Passo 3 — Integrar o relatório final

Use `prompts/codex-integracao-relatorio.md` para fundir:

- JSON formal;
- relatório Markdown preliminar;
- achados semânticos.

O relatório final deve preservar os números e achados formais do Python e acrescentar a seção semântica editorial.

### Passo 4 — Baixar saídas pelo app

Volte ao app Streamlit para baixar:

- relatório Word;
- JSON formal;
- verbete marcado com highlights, quando suportado;
- verbete corrigido, quando o arquivo de origem for DOCX.

## Referências normativas

- `references/estrutura-canonica.md`: ordem oficial, cabeçalho, divisões, seções fixas e Chapa Verbetográfica.
- `references/regras-por-secao.md`: finalidade, conteúdo, mínimos, máximos e correlações seção a seção.
- `references/confor-e-estilistica.md`: critérios globais, marca de excelência, logias, máximos, mínimos e tipos de enumeração.
- `references/formatacao-especifica.md`: fonte única para negrito, itálico, pontuação, versalete, separadores, padrão zero e Frase Enfática.
- `references/checklist-revisao.md`: Checklist Revisão oficial em 30 itens.
- `references/motor-formal.md`: documentação da camada Python e da separação Python × Codex.

## Estrutura do relatório

O relatório final deve seguir esta estrutura:

```markdown
# Relatório de Conformidade Verbetográfica
**Verbete:** <título> · **Especialidade declarada:** <especialidade> · **Tema central:** <tema>

## 1. Síntese executiva
<tabela de checklist>

## 2. Painel da marca de excelência
| Indicador | Exigido | Encontrado | Situação | Itens encontrados no texto |
|---|---:|---:|---|---|

## 3. Confor (estilística e formatação)
<achados formais>

## 4. Correlações internas
<correlações automáticas e observações editoriais pertinentes>

## 5. Seções fixas ausentes ou fora de ordem
<achados estruturais>

## 6. Análise semântica editorial
<achados da LLM/Codex>

## 7. Lista priorizada de correções
<lista única por severidade>
```

Severidades: **Crítico** para impedimento de conformidade; **Importante** para correção relevante; **Refinamento** para oportunidade de qualificação.

## Postura

Seja técnico, preciso e construtivo. Não invente regras fora das referências. Quando algo depender de juízo editorial sem resposta única, declare a ambiguidade e ofereça alternativas. Responda sempre em português.
