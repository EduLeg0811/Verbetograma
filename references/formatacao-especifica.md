# Formatação Específica

Fonte normativa única para as regras formais finas de realce, pontuação, itálico, versalete, separadores e padrão zero. Use este arquivo para interpretar achados do motor Python e para revisão humana/LLM. A verificação automática, quando possível, fica em `scripts/analisar_verbete.py`; estes textos `.md` não são lidos diretamente pelo algoritmo.

## 1. Epígrafe e pontuação em negrito

### Regra absoluta do dois-pontos

Quando o `:` em negrito abre uma enumeração, o ponto final que encerra a enumeração também deve ficar em negrito, salvo as exceções explícitas abaixo.

| Caso | Ponto final esperado |
|---|---|
| Epígrafe ou subtítulo com `:` em negrito abrindo enumeração | Negrito |
| Epígrafe encerrada por `.` em negrito, sem `:` introdutório | Plano, sem negrito |
| Parágrafo em prosa após uma lista já encerrada | Plano, sem negrito |

Exceção: em Sinonimologia e Antonimologia, mesmo quando a seção abre com `:` em negrito, o ponto final deve ficar em fonte normal, sem negrito.

Em itens numerados, a numeração inicial (`01.`, `1.`, etc.) não conta como epígrafe. A regra deve desconsiderar a numeração e avaliar a primeira expressão do item. Exemplo: em `01. Adcons: ...`, se `Adcons:` está em negrito, o ponto final do parágrafo deve ficar em negrito.

### Listas verticais

Em listas verticais, o `:` só deve ficar em negrito quando vier imediatamente após palavra ou trecho em negrito. Nesse caso, o ponto final do último item da lista também deve ficar em negrito.

Quando o `:` encerra parágrafo introdutório em fonte normal, como em enunciados longos de Taxologia, o `:` deve permanecer normal e não governa o ponto final do último item.

Exceção: na Remissiologia, o `:` que abre a lista vertical deve ficar em fonte normal, sem negrito.

### Epígrafe

A epígrafe e o sinal imediato (`.` ou `:`) devem ficar em negrito. O texto explicativo subsequente fica plano, salvo grifos próprios, itálicos ou sublinhamentos exigidos por regra.

### Frase adicional

Quando uma lista é seguida por frase adicional em prosa, apenas o ponto que fecha a lista acompanha o negrito do `:` introdutório. O ponto da frase adicional fica plano.

## 2. Ponto final plano

Em parágrafo de prosa sem `:` introdutório em negrito, o ponto final fica plano. Exemplo típico: `Definologia. O(a) <título> é ... .`

## 3. Padrão zero e numeração

Em enumeração numerada com 10 ou mais itens, os itens 1 a 9 recebem zero à esquerda: `01.` a `09.`. Listagens de até 9 itens usam forma simples: `1.` a `9.`.

Em enumeração vertical com mais de 99 itens, alinhar visualmente a numeração conforme o padrão da chapa.

## 4. Separadores e sinais gráficos

- Em Sinonimologia e Antonimologia: dois espaços entre número e acepção; acepções separadas por ponto; subitens por ponto e vírgula.
- Em listagens horizontais por ponto e vírgula, o separador `; ` permanece sem itálico.
- Hífen (`-`): usado quando todos os elementos interconectados são palavras simples.
- Travessão (`–`): usado quando há expressão composta entre elementos.
- Barra (` / `): usada somente em antagonismos, com espaço antes e depois.
- Aspas devem abrir e fechar, especialmente em Etimologia, Citaciologia e Ortopensatologia.

## 5. Itálico e sublinhamento

- Título do verbete citado na Definologia e na Remissiologia: em itálico/sublinhamento, mantendo maiúscula apenas na primeira palavra quando citado como verbete remetido.
- Estrangeirismos no corpo do texto e em Estrangeirismologia: em itálico. A identificação do que é estrangeirismo ou expressão estrangeira fica para LLM/revisor humano; o motor Python não verifica essa regra.
- Expressões compostas notáveis nas seções formais correspondentes ficam em itálico da palavra de classe até o próximo `;` ou `.`. Casos automatizados: Sinergismologia/sinergismo, Principiologia/princípio, Codigologia/código, Teoriologia/teoria, Tecnologia/técnica, Laboratoriologia/laboratório, Colegiologia/Colégio Invisível, Efeitologia/efeito, Ciclologia/ciclo, Binomiologia/binômio, Interaciologia/interação, Crescendologia/crescendo, Trinomiologia/trinômio, Polinomiologia/polinômio, Antagonismologia/antagonismo, Paradoxologia/paradoxo, Legislogia/lei, Sindromologia/síndrome e Mitologia/mito.
- Sufixos repetidos, como `cracia`, `filia`, `fobia` e `teca`, ficam em itálico apenas quando a seção tem 7 ou mais itens terminados naquele sufixo. O separador `; ` não fica em itálico.
- Artigos em Estrangeirismologia, Detalhismo, Perfilologia e Culturologia não devem ficar em itálico quando a regra exigir apenas o termo/construto.

## 6. Versalete, fonte e realces especiais

- Entrada/título: Arial, tamanho 11, negrito-itálico, versalete, com borda dupla.
- Especialidade: Arial, tamanho 11, versalete, espaçamento de caractere expandido em 1,5 pt, sem subtítulo.
- Texto geral: Times New Roman, tamanho 10, espaçamento simples.
- Cabeçalho: "Enciclopédia da Conscienciologia", Times New Roman, itálico, tamanho 9; paginação no canto superior direito.
- Frase Enfática: verificar centralização, 4 linhas quando detectáveis, 2 espaços entre palavras e grifo interno; o motor Python não verifica tamanho nem formato da fonte nesta seção.
- Na Frase Enfática, ao menos uma expressão-chave deve estar em itálico sem negrito. Megapensene trivocabular, se incluído, fica em itálico.

## 7. Viúvas e quebras

- Evitar número ou vogal isolada no fim da linha.
- Na Frase Enfática, vocábulo de duas letras no fim de linha também é viúva.
- Conferir divisão de palavras e repetição de sinais no fim/início de linha.

## 8. Checklist rápido

1. Epígrafe e sinal imediato estão em negrito?
2. Se há `:` em negrito abrindo enumeração, o ponto final correspondente está em negrito?
3. Se é prosa sem `:` introdutório, o ponto final está plano?
4. Listas verticais encerram com ponto final em negrito quando abertas por `:` em negrito?
5. Numeração com 10 ou mais itens usa padrão zero?
6. Separadores `; ` permanecem sem itálico?
7. Estrangeirismos e expressões compostas notáveis estão em itálico?
8. Sufixos repetidos estão em itálico apenas quando a seção atinge 7 itens?
9. Título e Especialidade respeitam fonte, versalete e realces? Na Frase Enfática, conferir apenas centralização e grifo interno no motor Python.
10. Não há viúvas ou quebras indevidas?
