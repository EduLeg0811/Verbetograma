# Motor Formal e Catálogo de Regras

A camada formal vive em `scripts/` e separa dados declarativos, lógica de análise e interface. Os arquivos `.md` de `references/` não são carregados pelo algoritmo Python; eles servem como base normativa para Codex/LLM, UI e revisor humano.

## Arquivos usados pelo Python

- `scripts/catalogo.py`: catálogo declarativo de divisões, ordem, seções fixas e opcionais, mínimos, limiares de máximos, tipos de enumeração, totais permitidos, marca de excelência e parasitas.
- `scripts/analisar_verbete.py`: motor formal. Faz parsing de `.docx`, `.doc`, `.pdf`, `.txt` e `.md`; classifica divisões/seções; aplica regras objetivas; gera JSON e relatório Markdown.
- `scripts/editar_verbete.py`: gera DOCX marcado, DOCX corrigido e PDF marcado quando há suporte local.
- `scripts/inspecionar_confor.py`: utilitário auxiliar para inspeção pontual de parágrafos e pontuação.

## Saída JSON

```json
{
  "meta": {"titulo": "...", "especialidade": "...", "tema_central": "...", "paginas": 4},
  "marca_excelencia": {"paginas": {}, "maximos": {}, "logias": {}},
  "secoes_presentes": ["Definologia"],
  "secoes_texto": {"Definologia": "..."},
  "contagens": {"Definologia": 1},
  "maximos": [{"secao": "Cognatologia", "itens": 48, "limiar": 10}],
  "logias": 40,
  "logias_lista": ["Cogniciologia"],
  "formatacao": {"tipo": "docx", "verificados": 0, "conformes": 0, "ajustes": [], "checks": []},
  "achados": [{"severidade": "Importante", "secao": "Sinonimologia", "regra": "...", "evidencia": "...", "sugestao": "..."}]
}
```

Severidades emitidas: `Crítico`, `Importante`, `Refinamento`.

## Implementado

- Estrutura: divisões presentes/ordenadas; seções fixas presentes; seção em divisão esperada; ordem interna das seções dentro da divisão.
- Contagens: itens por seção; enumerações verticais por número no início de linha; Etimologia tratada como parágrafo comum mesmo quando contém `;`; Enumerologia com exatamente 7 itens e máximo automático quando existe; mínimos; máximos por seção; máximo da divisão Detalhismo; páginas quando disponíveis.
- Marca de excelência: páginas, máximos e logias.
- Logias com exclusões automáticas: Entrada, Especialidade, divisões, Frase Enfática, corpo da Interdisciplinologia e corpo da Remissiologia.
- Padrão zero em enumerações numeradas.
- Remissiologia: total permitido em 7, 10, 12 ou 15 e alerta de espaçamento duplo entre itens quando detectável.
- Ordem alfabética em Cognatologia, Sinonimologia, Antonimologia, Taxologia, Caracterologia, Tipologia, Tabelologia e Remissiologia.
- Sinonimologia/Antonimologia: dois espaços após número e separação de acepções principais por ponto.
- Correlações objetivas: título em Definologia/Neologia/Exemplologia/Remissiologia; Especialidade em Neologia/Interdisciplinologia.
- Questionologia com `leitor ou leitora`.
- Parasitas da linguagem por busca textual fora de aspas.
- Regras textuais específicas: aspas abre-fecha em Etimologia/Citaciologia/Ortopensatologia; Megapensene trivocabular com verbo explícito ou elipse por `:`; limite, ordem e fórmulas objetivas de Ortopensatologia; artigo apenas após `=` na Exemplologia; estrutura mínima de Bibliografia; espelhamento quantitativo Masculinologia x Femininologia.
- Auditoria DOCX de negrito/pontuação: epígrafe, `:` introdutório, ponto final em negrito, ponto final em item numerado com epígrafe em negrito, ponto final normal em Sinonimologia/Antonimologia, `:` normal na Remissiologia e ponto final plano em prosa.
- Auditoria DOCX de itálico e sinais: título citado, expressões compostas notáveis, sinais objetivos de sublinhamento, artigo inicial sem itálico, separador `;` sem itálico, sufixos repetidos e grifo interno em itálico sem negrito da Frase Enfática.
- Auditoria DOCX de estilo explícito: papel carta, margens, medianiz, cabeçalho, paginação, texto geral, título, Especialidade e Frase Enfática quando os atributos aparecem diretamente no XML; Frase Enfática sem verificação de tamanho/formato da fonte.

## Parcialmente Implementado

- Logias: aplica exclusões automáticas, mas ainda pode exigir revisão humana quando a ocorrência aparece em texto corrido ambíguo.
- Frase Enfática: detecta presença, centralização, grifo interno, 2 espaços entre palavras e 4 linhas quando as quebras estiverem detectáveis; não valida tamanho nem formato da fonte.
- Neologia x Exemplologia: há correlação textual aproximada, mas não equivalência conceitual completa.
- Estrangeirismos: identificação e recomendação de itálico ficam reservadas à LLM/Codex, pois dependem de juízo semântico/editorial.
- Contagens: usam heurísticas textuais em seções atípicas.
- Layout e estilo DOCX: leitura limitada a atributos explícitos no XML, sem resolver toda a herança de estilos do Word.
- Ordem alfabética e Bibliografia: verificações objetivas úteis, mas ainda sujeitas a revisão humana em casos editoriais complexos.

## Não Implementado no Python

- Espaçamento duplo completo da Remissiologia entre todas as palavras e pontuações quando não houver quebras/texto detectáveis.
- Viúvas, quebras de linha e divisão de palavras.
- Agrupamento interno de subtítulos em Ortopensatologia além das fórmulas objetivas.
- Revisão bibliográfica completa além da estrutura mínima objetiva.
- Iniciais do verbetógrafo ao final.

## Reservado à LLM/Codex

Tudo que exige juízo editorial:

- adequação do tema central;
- qualidade da Definologia;
- pertinência de sinônimos e antônimos;
- construção conceitual de binômios, trinômios, polinômios, antagonismos, ciclos, interações e paradoxos;
- critérios qualificadores gerais, pesquisa exaustiva, Enumerologia semântica, Megapensenologia não justaposta, Estrangeirismologia editorial, Arcaisticologia, Unidade, Seciologia e fidelidade de transcrição em Ortopensatologia;
- pertinência e exaustividade da Remissiologia;
- coerência intra e interverbetes;
- escolha da Especialidade;
- força tarística da Frase Enfática e derivação da Questionologia;
- concordância, regência, clareza e cacófatos não detectáveis automaticamente;
- recomendações de qualificação editorial.

Use `prompts/codex-analise-semantica.md` para essa etapa e `prompts/codex-integracao-relatorio.md` para fundir achados formais e semânticos.

## Como Estender

- Mudar mínimo, máximo, divisão ou tipo de enumeração: editar `scripts/catalogo.py`.
- Nova regra objetiva: implementar função em `scripts/analisar_verbete.py` e incluir o achado no JSON.
- Nova regra normativa para revisor/LLM: atualizar `references/`.
- Novo fluxo operacional para Codex: atualizar `prompts/` e `scripts/SKILL.md`.
