# Projeto Verb-v0

`Verb-v0` é um projeto local para revisão de verbetes da Enciclopédia da Conscienciologia. Ele combina uma aplicação Streamlit, um motor formal em Python, referências normativas e prompts Codex para análise semântica.

## Estrutura geral

```text
Verb-v0/
├── app.py
├── requirements.txt
├── assets/
│   └── README.md
├── prompts/
│   ├── codex-analise-semantica.md
│   └── codex-integracao-relatorio.md
├── references/
│   ├── checklist-revisao.md
│   ├── confor-e-estilistica.md
│   ├── estrutura-canonica.md
│   ├── formatacao-especifica.md
│   ├── motor-formal.md
│   └── regras-por-secao.md
└── scripts/
    ├── analisar_verbete.py
    ├── catalogo.py
    ├── editar_verbete.py
    ├── inspecionar_confor.py
    └── SKILL.md
```

## Aplicação Streamlit

`app.py` é a interface executável do projeto. Ela permite enviar verbetes em `.docx`, `.pdf`, `.txt`, `.md` e, quando o ambiente permitir conversão via LibreOffice ou Microsoft Word/COM, `.doc`.

Principais funções:

- upload do verbete;
- execução do motor formal local;
- painel de achados críticos, importantes, máximos e logias;
- consulta das referências normativas pela barra lateral;
- abas de Relatório, Texto, Achados e JSON;
- download do relatório Word e do JSON formal;
- download do verbete marcado e, para DOCX, do verbete corrigido.

Para executar:

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Camada Python

A pasta `scripts/` contém a camada determinística.

- `scripts/SKILL.md`: define o comportamento operacional do skill local.
- `scripts/catalogo.py`: catálogo declarativo de divisões, seções, mínimos, máximos, logias, marca de excelência e parasitas.
- `scripts/analisar_verbete.py`: motor formal, CLI e gerador do relatório Markdown preliminar.
- `scripts/editar_verbete.py`: gera cópias marcadas/corrigidas quando o formato permite.
- `scripts/inspecionar_confor.py`: utilitário auxiliar para inspeção pontual.

O Python cobre: estrutura, seções fixas, mínimos, máximos, logias, páginas, padrão zero, total da Remissiologia, correlações string-matcháveis, Questionologia, parasitas e auditoria DOCX de epígrafe/pontuação quando possível.

## Referências normativas

`references/` contém a base normativa usada por Codex, pela interface e pelo revisor humano. Esses arquivos não são carregados diretamente pelo algoritmo Python.

- `estrutura-canonica.md`: ordem oficial do verbete, cabeçalho, divisões I a VI, seções fixas e Chapa Verbetográfica.
- `regras-por-secao.md`: finalidade, conteúdo, mínimos, máximos e correlações por seção.
- `confor-e-estilistica.md`: critérios qualificadores, marca de excelência, logias, máximos, mínimos e tipos de enumeração.
- `formatacao-especifica.md`: fonte única para negrito, itálico, pontuação, versalete, separadores, padrão zero, Frase Enfática e viúvas.
- `checklist-revisao.md`: Checklist Revisão oficial em 30 itens.
- `motor-formal.md`: documentação da camada Python e da separação Python × Codex.

## Prompts Codex

`prompts/` contém instruções executáveis para a etapa LLM.

- `codex-analise-semantica.md`: submeter o verbete + JSON formal ao Codex para análise editorial sem recalcular dados do Python.
- `codex-integracao-relatorio.md`: fundir JSON formal, relatório preliminar e achados semânticos no relatório final.

## Fluxo indicado do verbete

1. Rodar primeiro o app Streamlit ou CLI Python sobre o `.docx`, gerando JSON formal e relatório preliminar.
2. Submeter à LLM Codex o verbete + JSON formal usando `prompts/codex-analise-semantica.md`.
3. Integrar achados formais e semânticos com `prompts/codex-integracao-relatorio.md`.
4. Voltar ao app Streamlit para baixar o relatório, o verbete marcado e, quando possível, o verbete corrigido.

Essa ordem evita redundância: o Python produz os dados objetivos; Codex interpreta o conteúdo; a integração final apenas funde as camadas.

## Saída esperada

O relatório final segue a estrutura definida em `scripts/SKILL.md`:

- Relatório de Conformidade Verbetográfica;
- Síntese executiva em tabela de checklist;
- Painel da marca de excelência com itens encontrados;
- Confor, estilística e formatação;
- Correlações internas;
- Seções fixas ausentes ou fora de ordem;
- Análise semântica editorial;
- Lista priorizada de correções.

A visão por seção, com texto completo e correções, fica na aba `Texto` do app para evitar redundância dentro do relatório.

As severidades usadas são:

- **Crítico:** impede conformidade.
- **Importante:** exige correção relevante.
- **Refinamento:** oportunidade de qualificação.

## Checklist geral do projeto

1. Rodar motor Python e gerar JSON formal.
2. Conferir painel de marca de excelência.
3. Conferir a aba `Texto` para tabela de conformidade por seção.
4. Baixar verbete marcado/corrigido quando aplicável.
5. Submeter verbete + JSON ao prompt semântico Codex.
6. Integrar relatório final com o prompt de integração.
7. Validar se as referências normativas não duplicam regras formais.
8. Manter `references/` para normas, `prompts/` para LLM e `scripts/` para algoritmo.

## Próximos passos técnicos

1. Incluir `assets/chapa-verbetografica-oficial.doc`, se fizer parte do pacote oficial.
2. Criar um pequeno conjunto de documentos de teste para confirmar JSON formal, relatório final e downloads marcados/corrigidos.
