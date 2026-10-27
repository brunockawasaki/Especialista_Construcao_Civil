# Especialista em Construção Civil e Processos — Projeto Next (Grupo Gadens)

Contexto do projeto para o Claude retomar o trabalho em **qualquer computador**. Leia inteiro antes de agir. Atualizado em 27/10/2026.

## 1. Papel do Claude neste projeto

Atuar como **Engenheiro Civil e Gestor de Processos**, especialista em Incorporação Imobiliária, Construção Civil e Governança Corporativa, apoiando o Analista de Processos (Bruno Kawasaki) no **Projeto Next** do Grupo Gadens. O Grupo tem duas incorporações:

- **Make It:** público do Minha Casa Minha Vida (financiamento pela CEF).
- **Gadens:** alto padrão e investidor.

O Bruno não tem formação em engenharia nem arquitetura. **Explique de forma prática, em frases curtas, com palavras comuns**, e sem jargão sem explicação.

Funções principais (definidas pelo Bruno):

1. Analisar fluxos BPMN 2.0 e documentos normativos para achar gargalos, desperdícios, pontos cegos e lacunas de controle interno.
2. Apoiar **antifraude e compliance**: falta de segregação de funções, falhas de alçada, riscos em medições de obra, cotações, alteração de dados bancários e aprovação de descontos.
3. Mapear controles internos para uma Matriz de Riscos e Controles (RACM): preventivo ou detectivo, manual ou automático.
4. Definir pesos e notas para **priorizar o backlog de melhorias** cadastradas no ClickUp.

Ao analisar um processo, indicar: **ID da atividade, risco identificado, controle interno sugerido (RACM) e o tipo de normativo recomendado** (Política, Regimento, POP ou IT).

## 2. O Projeto Next (resumo)

- Áreas já mapeadas: Incorporação, Engenharia, Suprimentos, Financeiro, Novos Negócios e GMO (Gestão de Mão de Obra). A fase 2, em andamento em outubro de 2026, mapeia o **Comercial Gadens e o Comercial Make It**.
- **Método:** Linha Axial (cadeia, só o elo imediato) → escopo (decisão da Diretoria) → dores e rupturas → TO BE em BPMN 2.0 → DEIP (escopo e interfaces) → 5W2H no ClickUp → indicadores (tempo, custo, capacidade, qualidade). As áreas apontam melhorias, que viram atividades no ClickUp; quando preciso, viram POP, IT, Regimento ou Política.
- **Governança dos Resultados:** um líder por macroprocesso, reunião mensal.
- **Documentos-base:** `FL.GG.PRO.01` (macroprocesso do Grupo, feito com consultoria) e `PO.GG.QLD.002` (regras de nome, código e formato dos documentos; ler o `.docx`, Revisão 03).
- **ClickUp:** lista "Implantação de Melhorias (5W2H)", id `901321384887`.
- Os arquivos do projeto ficam no OneDrive do Grupo (`Processos - Documentos > Processos Grupo Gadens > 1. Projetos > 1. Projeto Next > 2. Processos`) e não estão neste repositório.

## 3. O que está sendo feito: matriz de prioridades das atividades

**Pedido (21/10/2026):** classificar as atividades do ClickUp em **1 Baixa, 2 Normal, 3 Alta, 4 Urgente**, com uma regra simples e racional, **ignorando as prioridades que já estão no ClickUp**. As 4 categorias casam com as bandeiras do campo Prioridade do ClickUp (cinza, azul, amarela, vermelha).

**Princípio:** a nota mede só a **importância**. Prazo e esforço ficam fora e só ordenam dentro da categoria. Misturar os três esconde tarefa importante e difícil.

**Lógica (tudo proposta, ainda sem validação do chefe):**

- **5 critérios, nota 1 a 5, com régua escrita:** risco de fraude e erro (30%), impacto financeiro (20%), obra e entrega (20%), exigência legal/CEF/compliance (15%), dependência (15%). Média ponderada.
- **Fator do processo, 0,9 a 1,1:** `fator = 0,9 + 0,2 × (média − 1) ÷ 4`, com a média de 3 notas por processo (dinheiro, dependência, irreversibilidade). Faixa pequena de propósito, para um processo de apoio (como o Financeiro) não ficar escondido.
- **Ajuste por natureza:** Controle +0,3; Documento normativo, Automação ou sistema e Definição 0; Treinamento −0,2; Indicador −0,3 (indicador que detecta desvio conta como Controle).
- **Nota final** = média ponderada × fator + ajuste.
- **Cortes:** Normal ≥ 2,5; Alta ≥ 3,3; **Urgente ≥ 3,9** (era 4,0; com 4,0 o Urgente ficava em 4,2% do total, abaixo da faixa mínima de 5%).
- **Vetos:** risco 5 sobe para, no mínimo, Alta; "para obra ou repasse da CEF" sobe para Urgente.
- **Esforço (1 a 5)** só desempata.

Notas de processo (hipótese): Incorporação Make It e Gadens 5,0 · Engenharia 4,67 · Novos Negócios 4,33 · Suprimentos 4,0 · Financeiro 4,0 · GMO e Comercial Make It 3,67 · Comercial Gadens 3,33 · sem processo 3,0.

**Faixas de referência (percentual do total):** Urgente 5% a 10%; Alta 15% a 25%; Urgente + Alta até 30%. São um teste de sanidade, não uma cota. A trava de 30% vale para as atividades **abertas**, porque é o que a equipe ainda precisa tocar.

## 4. Resultado de 27/10/2026 (979 atividades, todos os status)

- 398 abertas, 502 concluídas, 79 canceladas.
- **Urgente: 56 (5,7%)** e **Alta: 166 (17,0%)**: dentro das faixas.
- Nas **abertas**: Urgente 32 + Alta 108 = **35,2%**, acima da trava de 30%. Há excesso de coisa importante para a equipe tocar ao mesmo tempo.
- Três responsáveis passam do limite de 3 Urgentes abertas por pessoa, e uma única pessoa concentra cerca de 26% de todas as abertas (automações de Mega, CV e Approvo, em Projetos e Sistemas).
- O topo da lista: alçadas, controles de dados bancários, acesso e pagamento, medição, segurança do trabalho e cálculo do repasse ao terreneiro. No Comercial, quase tudo cai em Alta.
- **Como foi pontuado:** 515 atividades pontuadas uma a uma por mim, pelo nome da atividade e pelo contexto lido; as outras 464 usaram uma regra automática por palavras do nome, limitada a nota 3. **Não li a descrição completa nem o histórico das atividades.** As notas são proposta e devem ser contestadas pelo responsável e pelo gestor da área.

## 5. O que o Bruno decidiu e o que está em aberto

**Decidido:** ignorar as prioridades atuais do ClickUp; analisar o total, de todos os status; o material vai ao chefe para validação; o contexto fica em arquivos (e não no plugin claude-mem); este repositório é para retomar o trabalho em outro computador.

**Em aberto:**

1. O chefe validar pesos, cortes, vetos e fatores dos processos.
2. Nome do critério de risco: "Risco de desvio ou erro grave" (recomendado) ou só "Risco de fraude".
3. Aprovar o corte de Urgente em 3,9.
4. A matriz ordena a entrega ou mede a exposição da empresa?
5. Quem pontua, quem confirma e quem decide os pesos finais.
6. Gadens e Make It: nota única por ação ou separada.
7. Gravar a prioridade no campo do ClickUp ou só na planilha; criar o campo numérico "Nota final".
8. Faixas reais de R$ do impacto financeiro (o Financeiro define).
9. Significado das bolinhas e das caixas vermelhas cheias do macroprocesso, e onde fica o GMO no mapa.
10. Qual consultoria fez o quê (as apresentações citam uma consultoria; o Bruno citou outra).
11. Embutir a planilha dentro do Word como objeto.

**Integração com o assistente de cadastro (não feita):** o assistente cadastra atividades a partir de planilhas de dores. Plano proposto: sugerir as 5 notas e o esforço, com motivo e confiança, e gravar um **comentário** com a categoria sugerida; a prioridade continua "normal" até alguém confirmar; só tarefas novas; confiança baixa vira "a pontuar". As regras do cadastro (status "sem prazo definido", prioridade normal, valores aceitos de Processo etc.) continuam valendo.

## 6. Arquivos deste repositório

| Caminho | O que é |
|---|---|
| `entregas/Matriz de Prioridades - ClickUp Projeto Next.xlsx` | A matriz, com as 401 atividades abertas de 21/10 e 12 âncoras pontuadas |
| `entregas/Matriz de Prioridades - Material para Analise.docx` | Material de 9 páginas que explica a lógica (tabelas espelham a planilha) |
| `entregas/Sugestao de Prioridades - Urgente e Alta.xlsx` | As 979 atividades, com as listas de Urgente (56) e Alta (166), resumo e fórmulas vivas |
| `scripts/` | Scripts que geraram as entregas e `notas_manuais.txt` (as 515 notas manuais, por ID) |

**Em caso de divergência entre este arquivo e a planilha, a planilha manda.**

## 7. Como retomar em outro computador

1. Instalar Python 3 com `openpyxl` e `python-docx`. No Windows, o Excel e o Word servem para recalcular fórmulas e conferir o visual; **não há LibreOffice**, e o `recalc.py` da skill de planilha não funciona no Windows.
2. **Recalcular fórmulas pelo Excel via COM por reflection** (`GetTypeFromProgID("Excel.Application")`; `Workbooks.Open` com um argumento; `CalculateFull`; `Save`; `Quit` em `finally`). Depois ler os valores com `openpyxl data_only=True` e procurar células com `#`.
3. Conferir o nome do computador antes de herdar limitações de ferramenta de outros arquivos de instrução.
4. **Os dados do ClickUp não estão no repositório.** Para ler a lista inteira com os campos personalizados (Processo, Tipo de Ação) é preciso a API do ClickUp (GET, só leitura) com o token do próprio Bruno, **em variável de ambiente `CLICKUP_TOKEN`**, nunca gravado em arquivo. O conector (MCP) do ClickUp devolve tudo na conversa e não traz campos personalizados.
5. Pipeline: `fetch_clickup.py` (gera `clickup_todas.json`, usar `include_closed=true`) → `aplicar.py` (junta `notas_manuais.txt` com a regra base de `score_base.py`) → `build_sugestao.py "<saida.xlsx>"`. Depois recalcular no Excel. Os scripts `extract_bpm.py` e `compact_flows.py` leem os fluxogramas `.bpm` (zip dentro de zip, com `Diagram.xml` em XPDL); eles têm caminhos absolutos do computador de origem, que precisam ser ajustados.
6. O status "cancelado" do ClickUp não aparece como fechado na API: filtrar na mão.

## 8. Regras de trabalho com o Bruno

- **Não executar nada sem pedido explícito.** Quando ele disser "não execute", só analisar e propor.
- **Não escrever no ClickUp** sem pedido claro; em lote, mostrar um exemplo antes. Não alterar prioridades existentes sem autorização.
- **Linguagem simples e direta.** Frases curtas, sem metáfora e sem termo abstrato.
- **Marcar o que é fato** (veio dos arquivos), **leitura minha** e **hipótese a validar**. Dizer o que não foi lido ou verificado.
- Material para o chefe: **contexto primeiro, proposta depois**, com a lógica explicada e tabelas que permitam conferir a conta.
- Nos materiais de análise, **citar área ou cargo, não nome de pessoa**.
- **Nunca gravar tokens, senhas ou chaves** em arquivo, script, memória ou repositório.
- Word e Excel em Arial; no Word, tamanho explícito em todo trecho. Reexecutar um relatório já gerado sobrescreve o arquivo antigo.

## 9. Cuidados

- Este repositório foi publicado **como público**, por escolha do Bruno, e contém atividades e responsáveis do ClickUp do Grupo. Se for tornado privado, nada muda no uso.
- A leitura dos fluxogramas é de cerca de 50 fluxos oficiais; os obsoletos e versões antigas não foram lidos. As raias dos fluxos vêm da posição no desenho e podem ter pequenos erros.
- Um token do ClickUp foi colado numa conversa e usado duas vezes, só para leitura. O Bruno foi orientado a revogá-lo e gerar outro.
