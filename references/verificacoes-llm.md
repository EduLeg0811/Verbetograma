# Verificações LLM/Codex

Etapa posterior. Use os prompts em `prompts/` com o verbete e o JSON formal já gerado pelo Python.

| Grupo | Verificação | Prompt |
|---|---|---|
| Tema central | Adequação editorial de homeostático, neutro ou nosográfico ao conteúdo do verbete | `codex-analise-semantica.md` |
| Definologia | Se a definição é exclusivamente definidora, clara, concisa, abrangente e fiel ao título | `codex-analise-semantica.md` |
| Sinonimologia / Antonimologia | Pertinência conceitual dos sinônimos e oposição real dos antônimos | `codex-analise-semantica.md` |
| Conceitos compostos | Qualidade de binômios, trinômios, polinômios, antagonismos, ciclos, interações e paradoxos | `codex-analise-semantica.md` |
| Estrangeirismos | Identificação de estrangeirismos e expressões estrangeiras que exigem itálico, inclusive em Estrangeirismologia | `codex-analise-semantica.md` |
| Exemplologia | Pertinência dos exemplos e coerência conceitual com a Neologia | `codex-analise-semantica.md` |
| Neologia / Exemplologia / título | Quando o título é neologismo, ele deve estruturar a Neologia e ser coerente com os exemplos da Exemplologia | `codex-analise-semantica.md` |
| Enumerologia | Progressão ou crescendo sem uso literal indevido da palavra "crescendo" | `codex-analise-semantica.md` |
| Megapensenologia | Síntese tarística, verbo/elipse com sentido e ausência de mera justaposição de 3 palavras | `codex-analise-semantica.md` |
| Remissiologia | Relação estreita, relevância, exaustividade e adequação editorial dos verbetes remetidos | `codex-analise-semantica.md` |
| Ortopensatologia | Fidelidade da transcrição e preservação dos grifos originais | `codex-analise-semantica.md` |
| Seções especiais | Pertinência editorial de Arcaisticologia, Unidade e Seciologia quando existirem | `codex-analise-semantica.md` |
| Coerência | Contradições internas e coerência com a rede conceitual da Enciclopédia | `codex-analise-semantica.md` |
| Especialidade | Adequação da Especialidade declarada ao foco principal do verbete | `codex-analise-semantica.md` |
| Frase Enfática / Questionologia | Força tarística, síntese do megafoco e derivação das perguntas | `codex-analise-semantica.md` |
| Língua fina | Concordância, regência, clareza, cacófatos e parasitas residuais não resolvidos automaticamente | `codex-analise-semantica.md` |
| Travessão/hífen em expressão composta | Nos sublinhamentos formais, exigir travessão/en dash ou hífen sem espaços (em vez de hífen com espaços) quando os elementos conectados forem expressões compostas; julgamento semântico, não verificado pelo Python | `codex-analise-semantica.md` |
| Pesquisa | Exaustividade por Holociclo, Holoserver, Holoteca e Internet como checklist editorial | `codex-analise-semantica.md` |
| Critérios qualificadores | Clareza, coerência, concisão, contribuição, cosmovisão, detalhismo, estilística, estrutura, exaustividade, harmonia, lógica, originalidade, relevância, tares e usabilidade | `codex-analise-semantica.md` |
| Qualificação | Oportunidades editoriais para enriquecer seções e priorizar melhorias sem recalcular os números do Python | `codex-analise-semantica.md` |
| Relatório final | Fusão do JSON formal, relatório preliminar e achados semânticos sem substituir dados objetivos | `codex-integracao-relatorio.md` |

A LLM não é chamada automaticamente pelo app. Primeiro baixe o JSON formal; depois use os prompts Codex para a etapa semântica e integração final.
