# Regras do Projeto

Este documento separa as regras do projeto em 4 grupos: **implementado no Python**, **parcialmente implementado**, **não implementado no Python** e **reservado à LLM/Codex**.

## 1. Implementado no Python

Estas regras são verificadas de modo determinístico pelo app/CLI.

### Entrada e leitura

| Regra | Como o código verifica | Saída |
|---|---|---|
| Formatos aceitos | Lê `.docx`, `.pdf`, `.txt`, `.md` e tenta converter `.doc` para `.docx` com LibreOffice ou Microsoft Word/COM | Avisos em `achados` quando há limitação |
| DOCX como fonte fiel | Usa XML do Word para parágrafos, páginas e runs de negrito/itálico | `meta`, `formatacao`, `achados` |
| PDF em melhor esforço | Extrai linhas por `pypdf`, sem runs de Word | Aviso de confiabilidade |
| Texto plano | Lê linhas não vazias de `.txt` e `.md` | Sem auditoria fina de formatação |

### Metadados

| Regra | Como o código verifica | Saída |
|---|---|---|
| Título provável | Procura nas primeiras linhas texto que não seja divisão nem seção | `meta.titulo` |
| Especialidade declarada | Procura linha entre parênteses terminada em `logia` | `meta.especialidade` |
| Tema central textual | Busca `homeostático`, `neutro` ou `nosográfico` na Tematologia | `meta.tema_central` |
| Páginas | Lê metadado de páginas no DOCX ou número de páginas no PDF | `meta.paginas`, `marca_excelencia.paginas` |

### Estrutura canônica

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Divisões obrigatórias | Confere presença de Conformática, Fatuística, Detalhismo, Perfilologia, Argumentologia e Acabativa | Crítico |
| Ordem das divisões | Compara a sequência detectada com a ordem canônica | Crítico |
| Ordem interna das seções | Compara a sequência de seções dentro de cada divisão com o campo `ordem` do catálogo | Importante |
| Seções fixas e opcionais | Confere Definologia, Sinonimologia, Antonimologia, Remissiologia, Frase Enfática e Questionologia; reconhece Arcaisticologia, Unidade e Seciologia quando presentes | Crítico ou Importante |
| Seção na divisão esperada | Compara a divisão corrente com a divisão declarada em `scripts/catalogo.py` | Importante |

### Seções e textos

| Regra | Como o código verifica | Saída |
|---|---|---|
| Detecção de seção | Reconhece epígrafe antes de `.` ou `:` conforme nomes do catálogo e também novas epígrafes independentes com formato de seção, especialmente terminadas em `logia` | `secoes_presentes` |
| Texto completo por seção | Agrupa parágrafos sob a seção corrente até a próxima seção/divisão | `secoes_texto` |
| Frase Enfática | Tenta detectar bloco entre Remissiologia e Questionologia ou texto em maiúsculas sem `?` | `achados`, `secoes_texto` |

### Contagens, mínimos, máximos e padrão zero

| Regra | Como o código verifica | Saída |
|---|---|---|
| Contagem horizontal numerada | Conta padrões `N.` no corpo da seção | `contagens` |
| Contagem vertical numerada | Conta padrões `N.` no corpo da seção | `contagens` |
| Contagem por linhas | Para Fatologia/Parafatologia, estima por segmentos separados por `.` ou `;` | `contagens` |
| Contagem horizontal | Conta itens por `;` ou, em algumas seções, por vírgula; Etimologia é exceção e é tratada como parágrafo comum | `contagens` |
| Contagem vertical | Conta apenas números no início de linha em seções verticais, evitando números internos de referências bibliográficas | `contagens` |
| Mínimos obrigatórios | Compara a contagem com o campo `min` do catálogo | `achados` |
| Máximos | Se a contagem atinge o campo `max`, a seção entra em `maximos` | `maximos` |
| Enumerologia | Quando presente, deve conter exatamente 7 itens e entra como máximo automaticamente | `achados`, `maximos` |
| Máximo da divisão Detalhismo | Conta as seções detectadas na divisão Detalhismo e marca máximo com 20+ seções | `maximos` |
| Padrão zero | Em enumerações numeradas, exige `01` a `09` quando há 10+ itens e forma simples `1` a `9` quando há até 9 itens | `achados` |
| Ordem alfabética | Confere a ordenação dos itens em Cognatologia, Sinonimologia, Antonimologia, Taxologia, Caracterologia, Tipologia, Tabelologia e Remissiologia | `achados` |
| Dois espaços após número | Confere o espaço duplo depois do número em Sinonimologia e Antonimologia | `achados` |
| Separação de acepções | Em Sinonimologia e Antonimologia, alerta quando acepções principais parecem separadas por `;` em vez de ponto | `achados` |
| Artigo na Exemplologia | Confere se artigos aparecem apenas após o sinal `=` | `achados` |

### Mínimos verificados

| Seção | Mínimo |
|---|---:|
| Cognatologia | 3 itens |
| Sinonimologia | 2 acepções |
| Antonimologia | 2 acepções |
| Neologia | 2 itens |
| Exemplologia | 2 itens |

### Limiares de máximos verificados

| Seção | Conta como máximo a partir de |
|---|---:|
| Definologia | 7 itens |
| Sinonimologia | 10 acepções |
| Cognatologia | 10 cognatos |
| Neologia | 4 neologismos |
| Antonimologia | 10 acepções |
| Estrangeirismologia | 7 itens |
| Fatologia | 20 linhas/segmentos |
| Parafatologia | 20 linhas/segmentos |
| Seções do Detalhismo em catálogo, exceto Interdisciplinologia | 7 itens |
| Interdisciplinologia | 10 itens |
| Hominologia | 7 itens |
| Taxologia | 100 itens |
| Remissiologia | 10 itens |
| Enumerologia | 7 itens, sempre máximo quando existe |
| Bibliografia Específica | 10 itens |
| Divisão Detalhismo | 20 seções |

### Marca de excelência

| Indicador | Regra automatizada | Observação |
|---|---:|---|
| Páginas | Mínimo 3 | Só fica conclusivo quando o arquivo fornece páginas |
| Máximos | Mínimo 5 | Conta seções que atingem os limiares de máximo |
| Logias | Mínimo 25 | Conta nomes de seções terminados em `logia` e logias detectadas nos corpos das seções elegíveis |

### Remissiologia

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Total permitido | A quantidade de itens deve ser exatamente 7, 10, 12 ou 15; qualquer outro total (inclusive abaixo de 7) é inválido | Crítico |
| Máximo | A seção conta como máximo a partir de 10 itens | Painel de máximos |
| Espaçamento duplo entre itens | Alerta quando itens verticais aparecem sem linha dupla detectável | Refinamento |

### Correlações objetivas

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Título na Definologia | O título normalizado deve aparecer na seção | Importante |
| Título na Neologia | O título normalizado deve aparecer na seção | Importante |
| Título na Exemplologia | Usa correspondência por tokens relevantes do título | Importante |
| Título na Remissiologia | O título normalizado deve aparecer na seção | Importante |
| Especialidade na Neologia | A Especialidade declarada deve aparecer na seção | Importante |
| Especialidade na Interdisciplinologia | A Especialidade declarada deve aparecer na seção | Importante |

### Questionologia

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Fórmula de interlocução | A seção deve conter literalmente `leitor ou leitora` | Importante |

### Parasitas da linguagem

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Parasitas catalogados | Busca ocorrências fora de trechos entre aspas | Refinamento |

Termos buscados: `um`, `uma`, `uns`, `umas`, `num`, `numa`, `nuns`, `numas`, `meu`, `minha`, `meus`, `minhas`, `nosso`, `nossa`, `nossos`, `nossas`, `seu`, `sua`, `seus`, `suas`, `teu`, `tua`, `teus`, `tuas`.

### Regras textuais específicas

| Regra | Como o código verifica | Severidade |
|---|---|---|
| Aspas abre-fecha | Em Etimologia, Citaciologia e Ortopensatologia, confere pares de aspas retas, curvas simples e curvas duplas | Importante |
| Megapensene trivocabular | Em Megapensenologia, cada sentença deve ter exatamente 3 palavras e conter verbo explícito ou elipse por `:` | Importante |
| Ortopensatologia | Confere limite de até 3 ortopensatas, ordem alfabética e fórmulas objetivas com aspas, travessão ou lista numerada | Importante |
| Bibliografia Específica | Confere autor/Idem inicial, repetição de autor sem "Idem", ano, paginação, abreviação de edição ("Ed.") e capitalização após ":" em cada item | Importante |
| Filmografia Específica | Confere ficha técnica com 27 campos e ao menos 5 atores listados | Importante |

### Formatação DOCX: negrito e pontuação

Estas regras só são verificadas quando a entrada é `.docx`.

| Regra | Como o código verifica | Severidade quando falha |
|---|---|---|
| Epígrafe em negrito | Confere se o prefixo da seção até o primeiro `.` ou `:` está em negrito | Importante |
| Dois-pontos introdutório | Confere se o `:` que abre enumeração está em negrito | Importante |
| Ponto final espelha dois-pontos | Se a seção abre por `:` em negrito, o ponto final do parágrafo deve estar em negrito, exceto Sinonimologia e Antonimologia | Importante |
| Ponto final em item numerado com epígrafe | Em itens numerados, a numeração inicial é ignorada; se a epígrafe do item termina com `:` em negrito, o ponto final do parágrafo deve ficar em negrito | Importante |
| Ponto final normal em Sinonimologia/Antonimologia | Nessas seções, o ponto final deve ficar sempre sem negrito | Importante |
| Ponto final plano em prosa | Se a epígrafe usa `.` e não há `:` introdutório, o ponto final do parágrafo deve ficar sem negrito | Importante |
| Lista vertical aberta por `:` | Se o `:` vem imediatamente após palavra em negrito, também deve estar em negrito; se encerra parágrafo introdutório em fonte normal, deve ficar normal | Importante |
| Remissiologia com `:` normal | Excepcionalmente na Remissiologia, o `:` que abre a lista vertical deve ficar sem negrito | Importante |
| Último item da lista vertical | O ponto final do último item deve ficar em negrito apenas quando a lista foi aberta por `:` em negrito; não verificado em Bibliografia Específica | Importante |

### Formatação DOCX: itálico

| Regra | Como o código verifica | Severidade quando falha |
|---|---|---|
| Título citado em itálico/sublinhado | Em Definologia e Remissiologia, procura o título do verbete no corpo da seção e confere se os caracteres alfanuméricos estão em itálico ou sublinhados | Importante |
| Expressões compostas notáveis | Apenas nas seções formais correspondentes, confere itálico da palavra de classe até o próximo `;` ou `.`: Sinergismologia/sinergismo, Principiologia/princípio, Codigologia/código, Teoriologia/teoria, Tecnologia/técnica, Laboratoriologia/laboratório, Colegiologia/Colégio Invisível, Efeitologia/efeito, Ciclologia/ciclo, Binomiologia/binômio, Interaciologia/interação, Crescendologia/crescendo, Trinomiologia/trinômio, Polinomiologia/polinômio, Antagonismologia/antagonismo, Paradoxologia/paradoxo, Legislogia/lei, Sindromologia/síndrome e Mitologia/mito | Importante |
| Sinais de sublinhamento | Nas seções formais, confere barra com espaços em Antagonismologia e restringe barra à Antagonismologia | Importante |
| Travessão/hífen em expressão composta (manual) | "Quando houver expressão composta entre os elementos, usar travessão/en dash ou hífen sem espaços, nunca hífen com espaços" exige interpretação semântica do que conta como expressão composta; não é verificado pelo código, apenas pela LLM na revisão | — |
| Artigo inicial sem itálico | Em listagens horizontais separadas por `;`, confere se o artigo inicial do item (`a`, `as`, `o`, `os`) está sem itálico, mesmo quando a palavra seguinte está em itálico | Importante |
| Separador sem itálico | Confere se o caractere separador `;` não está em itálico; o espaço posterior não é considerado erro | Importante |
| Sufixos repetidos em itálico | Quando há 7 ou mais ocorrências com sufixos `cracia`, `filia`, `fobia` ou `teca`, confere se o sufixo está em itálico | Importante |
| Grifo interno da Frase Enfática | Confere se há trecho em itálico sem negrito na Frase Enfática | Importante |

### Formatação DOCX: layout e estilo por XML

Estas regras só são verificadas quando a entrada é `.docx` e os atributos aparecem explicitamente no XML do Word.

| Regra | Como o código verifica | Severidade quando falha |
|---|---|---|
| Título | Confere Arial 11, negrito-itálico e versalete quando os atributos estão explícitos | Importante |
| Especialidade | Confere parênteses, Arial 11 e versalete quando os atributos estão explícitos | Importante |
| Frase Enfática | Confere centralização, grifo interno em itálico sem negrito, 2 espaços entre palavras e 4 linhas quando detectáveis; não verifica tamanho nem formato da fonte | Importante |
| Filmografia Específica | Confere fonte tamanho 8 na ficha técnica quando o atributo está explícito | Importante |
| Layout DOCX | Confere papel carta, margens, medianiz, cabeçalho, paginação, texto geral, borda dupla, título e Especialidade quando os atributos aparecem explicitamente no XML | Importante |

### Downloads gerados pelo app

| Saída | Regra operacional |
|---|---|
| Verbete marcado DOCX | Aplica highlight amarelo nos trechos relacionados aos achados de correção |
| Verbete corrigido DOCX | Aplica correções automáticas possíveis de negrito/pontuação |
| PDF marcado | Marca trechos por busca textual quando a entrada é PDF |

## 2. Parcialmente Implementado

Estas regras têm alguma verificação automática, mas ainda dependem de heurística ou não cobrem todos os detalhes normativos.

| Regra | O que já existe | Limite atual |
|---|---|---|
| Logias | Exclui Entrada, Especialidade, divisões, Frase Enfática, corpo da Interdisciplinologia e corpo da Remissiologia | Ainda pode contar ocorrência em texto corrido que não funciona como subtítulo/item |
| Frase Enfática | Detecta presença, centralização, grifo interno, 2 espaços entre palavras e 4 linhas quando detectáveis | Ainda não valida tamanho nem formato da fonte |
| Título citado na Remissiologia | Confere presença e itálico/sublinhado quando encontra o título | Não valida maiúscula apenas na primeira palavra do verbete remetido |
| Neologia x Exemplologia | Confere correlação textual aproximada com o título | Não compara equivalência conceitual completa dos itens |
| Contagens horizontais | Conta por `;`, vírgulas em algumas seções e padrões numéricos | Pode falhar em redações atípicas ou itens compostos |
| Layout e estilo DOCX | Lê atributos explícitos de runs e propriedades do DOCX | Não resolve herança completa de estilos do Word |
| Ordem alfabética | Confere por chave textual normalizada | Pode exigir revisão humana em itens com artigos, siglas, nomes próprios ou subtítulos complexos |
| Bibliografia Específica | Confere autor/Idem, ano, paginação, abreviação de edição e capitalização pós dois-pontos por item | Heurística textual; não substitui revisão bibliográfica completa de conteúdo |
| Filmografia Específica | Conta campos rotulados (`Campo:`) e nomes em "Atores"/"Elenco" por ficha | Depende de a ficha seguir o padrão `Rótulo: valor`; não valida o conteúdo de cada campo |

## 3. Não Implementado no Python

Estas regras existem nas referências, mas não são verificadas automaticamente hoje.

- Espaçamento duplo completo da Remissiologia entre todas as palavras e pontuações quando não houver estrutura textual detectável.
- Viúvas, quebras de linha e divisão de palavras.
- Agrupamento interno de subtítulos em Ortopensatologia além das fórmulas objetivas.
- Iniciais do verbetógrafo ao final.

## 4. Reservado à LLM/Codex

Estes itens exigem juízo editorial e devem ser verificados com os prompts em `prompts/`.

- Adequação semântica do tema central.
- Qualidade conceitual da Definologia.
- Pertinência dos sinônimos e antônimos.
- Qualidade de binômios, trinômios, polinômios, antagonismos, ciclos, interações e paradoxos.
- Pertinência e exaustividade semântica da Remissiologia.
- Coerência intra e interverbetes.
- Adequação editorial da Especialidade.
- Força tarística da Frase Enfática.
- Identificação de estrangeirismos e expressões estrangeiras que devem receber itálico.
- Critérios qualificadores gerais e pesquisa exaustiva.
- Progressão semântica da Enumerologia, Megapensenologia não justaposta, conteúdo de Arcaisticologia, Unidade e Seciologia, e fidelidade de transcrição em Ortopensatologia.
- Derivação conceitual da Questionologia.
- Revisão fina de concordância, regência, clareza e cacófatos.
- Travessão/en dash ou hífen sem espaços vs. hífen com espaços em expressões compostas dos sublinhamentos formais (exige julgar o que conta como expressão composta).
- Vínculo entre título neologismo, Neologia e Exemplologia.
