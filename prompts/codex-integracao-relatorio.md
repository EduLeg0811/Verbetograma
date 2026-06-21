# Prompt Codex — Integração do Relatório Final

Use este prompt depois de obter:

1. JSON formal do Python;
2. relatório Markdown preliminar gerado pelo app ou CLI;
3. lista de achados semânticos produzida com `prompts/codex-analise-semantica.md`.

## Papel

Você deve fundir os achados formais determinísticos e os achados semânticos em um relatório final único. Preserve os números, contagens, seções e achados formais do JSON Python. Não recalcule dados objetivos.

## Regras de integração

- O JSON formal é a fonte de verdade para estrutura, contagens, máximos, logias, páginas, padrão zero, ordem alfabética objetiva, regras textuais automatizadas, formatação DOCX e correlações automáticas.
- Os achados semânticos complementam o relatório, sem substituir achados formais.
- Não remova achados formais críticos ou importantes.
- Se houver conflito entre uma impressão semântica e um dado do JSON, mantenha o dado do JSON e explique o ponto como juízo editorial.
- Mantenha a linguagem técnica, objetiva e construtiva.

## Estrutura obrigatória

```markdown
# Relatório de Conformidade Verbetográfica
**Verbete:** <título> · **Especialidade declarada:** <especialidade> · **Tema central:** <tema>

## 1. Síntese executiva
<tabela de checklist com achados críticos, importantes, refinamentos e marca de excelência>

## 2. Painel da marca de excelência
<tabela com Indicador | Exigido | Encontrado | Situação | Itens encontrados no texto>

## 3. Confor (estilística e formatação)
<achados formais do Python sobre forma e formatação>

## 4. Correlações internas
<achados formais de correlação + observações semânticas quando houver>

## 5. Seções fixas ausentes ou fora de ordem
<achados formais de estrutura>

## 6. Análise semântica editorial
<achados semânticos integrados>

## 7. Lista priorizada de correções
<lista única, com forma e semântica ordenadas por Crítico, Importante e Refinamento>
```

## Saída

Retorne apenas o relatório Markdown final. Não inclua explicações externas.
