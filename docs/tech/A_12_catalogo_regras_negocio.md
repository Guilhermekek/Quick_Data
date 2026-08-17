# 12. Catálogo de Regras de Negócio

> 112 regras de negócio identificadas no código, numeradas RN-001 a RN-115 (faixas reservadas por cluster durante a análise paralela). Índice compacto na seção 24.5.

## 12.1 Cluster Extração / ETL (RN-001 a RN-030)
### RN-001 — Identificador de Fonte por origem de extração
- **Nome:** Campo `Fonte` como chave de rastreabilidade de origem.
- **Descrição:** Cada linha inserida na aba Base carrega um valor fixo no campo `Fonte`
  ("Base_Hubble", "Base_1009", "Base_RGM", "Base_MOCKUP_RGM", "Base_OtherIncome", "Base_Consolidada",
  "Base_Quick_Data", "Base_Ajustes", "QD AUTOMATIC - SELEÇÃO MANUAL", "QD AUTOMATIC - INPUT AJUSTE"),
  usado tanto para gravação quanto para limpeza seletiva de histórico antes de nova extração.
- **Condição:** toda extração de qualquer um dos 12 módulos.
- **Ação executada:** `Form_Preenchimento_Generico(Sh_Destino, ..., "FONTE", <valor fixo>)` grava o literal
  correspondente em cada linha nova.
- **Exceções:** nenhuma — todas as 7 fontes externas + Ajustes + Pré-Closing seguem o padrão.
- **Local de implementação:** todos os `Processo_Extrair_Base_*` (item 15, 23, 26, 31, 35, 40, 43, 47, 51) e
  `Gerar_Base_Versao_Pre_Closing` (item 53).
- **Evidência:** ex. `Extracao_SQL_Hubble.bas:491` (`'Base_Hubble' AS Fonte`), `Extracao_Base_1009.bas:301-303`.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-002 — Limpeza prévia inconsistente entre módulos
- **Nome:** Remoção de histórico da mesma Fonte antes de reextrair.
- **Descrição:** Para evitar duplicidade, o padrão esperado é remover (`Limpar_Base_Historica`) os registros da
  mesma Fonte antes de gravar os novos. Isso está ativo em Hubble e 1009 (parcialmente) e Ajustes (sempre), mas
  **desativado ou ausente** em MOCKUP_RGM, RGM, Fixed_Revenues, Quick_Data e Other_Income/Consolidada (chamada
  comentada).
- **Condição:** execução do botão "Extrair_Base_X" sem passar antes pelo botão de limpeza correspondente.
- **Ação executada:** nas fontes afetadas, a linha `Call Processo_Limpar_Base_X` está comentada ou ausente —
  a extração é sempre um `INSERT`/`APPEND`, nunca um `REPLACE`.
- **Exceções:** `Extrair_Base_Hubble` chama `Processo_Limpar_Base_Hubble` ativamente (linha 11);
  `Extrair_Base_Ajustes` limpa incondicionalmente (linhas 10-12).
- **Local de implementação:** `Extracao_Base_MOCKUP_RGM.bas:3-19` (sem chamada, nem comentada),
  `Extracao_Base_RGM.bas:3-19` (idem), `Extracao_Fixed_Revenues.bas:3-19` (idem),
  `Extracao_Base_1009.bas:10` (comentada), `Extracao_Base_Consolidad.bas:11` (comentada),
  `Extracao_Base_Other_Inco.bas:11` (comentada), `Extracao_Base_Quick_Data.bas:11` (comentada).
- **Evidência:** ausência/comentário de `Call Processo_Limpar_Base_*` nas Subs de topo listadas.
- **Nível de confiança:** confirmado (verificado por leitura direta + grep).
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a limpeza é feita manualmente por um botão separado
  antes de cada extração (processo operacional) ou se há risco real de duplicação acumulada em produção.

### RN-003 — Escala de valores (multiplicador de milhão)
- **Nome:** Ajuste de escala para fontes RGM/MOCKUP_RGM/Fixed_Revenues.
- **Descrição:** RGM, MOCKUP_RGM e Fixed_Revenues aplicam `Form_Acertar_Escala(..., Multiplicador:=1000000)`
  pois a fonte externa registra valores em milhões; as demais fontes (1009, Hubble, Other Income, Consolidada,
  Quick Data, Ajustes) não aplicam esse multiplicador.
- **Condição:** sempre, para as 3 fontes citadas.
- **Ação executada:** multiplica todos os valores mensais por 1.000.000 antes de gravar na Base.
- **Exceções:** nenhuma dentro dessas 3 fontes.
- **Local de implementação:** `Extracao_Base_RGM.bas:233`, `Extracao_Base_MOCKUP_RGM.bas:235`,
  `Extracao_Fixed_Revenues.bas:236`.
- **Evidência:** as 3 linhas acima (`Call Form_Acertar_Escala(..., Multiplicador)` com `Multiplicador = 1000000`
  definido na linha imediatamente anterior em cada módulo).
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a unidade das demais 8 fontes está de fato correta sem
  necessidade de escala, ou se há uma inconsistência de unidade não tratada.

### RN-004 — Centro de Custo inferido por nome de arquivo (Other Income)
- **Nome:** Mapeamento de CDC por padrão textual do nome do arquivo.
- **Descrição:** Em Other Income, o Centro de Custo é determinado por 8 padrões `InStr` fixos aplicados ao
  nome do arquivo (TBP→XPHP999, CTCEL→XPHR999, TPAR→XPHS100, FIBER_RS→XPFR000, FIBER_RR→XPRU999,
  FIBER_OG→XPFS020, METIS_CZ→XPMRJ000, INTELIG_OG→XPHI999), com *fallback* XPHR999.
- **Condição:** para cada arquivo processado na pasta de Other Income.
- **Ação executada:** grava o CDC correspondente ao primeiro padrão que casar com `UCase(Arq)`.
- **Exceções:** se nenhum padrão casar, usa XPHR999 silenciosamente (sem aviso).
- **Local de implementação:** `Extracao_Base_Other_Inco.bas:201-211`.
- **Evidência:** linhas 202-211.
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** a lista completa de operadoras esperadas — novas
  operadoras exigirão alteração de código.

### RN-005 — Granularidade de Centro de Custo configurável no Hubble
- **Nome:** Extração "com CdC fiel" vs. "sem CdC" (consolidado).
- **Descrição:** Cada linha de configuração da matriz de extração Hubble pode pedir agregação por Centro de
  Custo granular ("c/ CdC Fiel") ou consolidada em `'-'` ("s/ CdC"), conforme a coluna "Extrair*CDC*" = "SIM".
- **Condição:** valor da coluna "Extrair*CDC*" na linha de configuração.
- **Ação executada:** monta `Filtro_CDC = [CENTRO CUSTO]` (granular) ou `Filtro_CDC = '-'` (consolidado), usado
  no `GROUP BY` do SQL.
- **Exceções:** nenhuma — é binário.
- **Local de implementação:** `Extracao_SQL_Hubble.bas:276-281`.
- **Evidência:** linhas 276-281.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-006 — Wildcard implícito em filtros de Visão/IFRS
- **Nome:** Valor vazio ou default equivale a "aceitar qualquer valor" (SQL `LIKE '%'`).
- **Descrição:** No Hubble, se o filtro de Visão/IFRS_9/IFRS_15 estiver vazio ou igual ao rótulo padrão ("IFRS
  Itália"/"w/ IFRS 9"/"w/ IFRS 15"), o filtro é substituído por `%` (wildcard SQL).
- **Condição:** `Filtro = "" Or UCase(Filtro) = UCase(<rótulo padrão>)`.
- **Ação executada:** `Filtro = "%"`, incorporado em `... LIKE '%' AND`.
- **Exceções:** aplica-se somente às 3 colunas citadas (VISAO, IFRS_9, IFRS_15).
- **Local de implementação:** `Extracao_SQL_Hubble.bas:260-266`.
- **Evidência:** linhas 260-266.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-007 — Tratamento IFRS16 seletivo por fonte
- **Nome:** IFRS16 aplicado à Base 1009, não ao Hubble.
- **Descrição:** `UPDATE_Tratar_IFRS16` é chamado ativamente pelo wrapper de extração da Base 1009, mas a
  chamada equivalente para Hubble está comentada.
- **Condição:** fonte extraída = "Base_1009" (ativa) vs. "Base_Hubble" (inativa).
- **Ação executada:** para 1009, aplica transformação IFRS16 (módulo `fx_IFRS16.bas`, fora do cluster) sobre os
  dados recém-extraídos; para Hubble, nenhuma transformação adicional.
- **Exceções:** nenhuma outra fonte do cluster chama `UPDATE_Tratar_IFRS16` diretamente no texto lido.
- **Local de implementação:** `Extracao_Base_1009.bas:16` (ativa) vs. `Extracao_SQL_Hubble.bas:16` (comentada).
- **Evidência:** as duas linhas citadas.
- **Nível de confiança:** confirmado (fato do código); a intenção de negócio por trás da diferença é inferida.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a ausência de tratamento IFRS16 no Hubble é intencional
  (dado já vem tratado do SQL) ou uma lacuna de manutenção.

### RN-008 — Formato combinado "KPI > Versão"
- **Nome:** Campo único que carrega KPI e Versão separados por " > ".
- **Descrição:** Em pelo menos 3 pontos do cluster, um campo de configuração/entrada permite informar
  "KPI > Versão" em uma única célula/coluna, sendo depois separado por `Split(" > ")`.
- **Condição:** o texto contém a substring " > ".
- **Ação executada:** `KPI = Split(Texto," > ")(0)`; `Versão = Split(Texto," > ")(1)`.
- **Exceções:** se não houver " > ", o valor é tratado como Versão pura (não some KPI).
- **Local de implementação:** `Extracao_SQL_Hubble.bas:217-218` (filtro KPI_VERSAO), `Extracao_SQL_Hubble.bas:
  299-336` (redefinição pós-extração via fórmula `FIND(">",...)`), `Extracao_Sheet_Ajustes.bas:158-161`.
- **Evidência:** as linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-009 — Exclusão interativa de receita para evitar dupla contagem
- **Nome:** Confirmação manual para remover "Other Income"/"Product Revenues"/"Net Service Revenues".
- **Descrição:** Em 1009, Other Income e Consolidada, o usuário é perguntado via `MsgBox` se deseja excluir
  linhas de receita que podem já estar cobertas por outra fonte, evitando dupla contagem no resultado
  consolidado.
- **Condição:** decisão manual do usuário a cada execução, sem valor padrão persistido.
- **Ação executada:** se "Sim", `Processo_Exclusao_Linhas_Base` remove as linhas do intervalo recém-inserido
  cujo campo Abertura_3 corresponde ao critério.
- **Exceções:** em Other Income, a exclusão de custo (Market/Process/Labour/Volume Driven Costs) é automática
  (sem pergunta); apenas "Net Service Revenues" é interativa; "Product Revenues" está desativada (comentada).
- **Local de implementação:** `Extracao_Base_1009.bas:406-461`, `Extracao_Base_Other_Inco.bas:393-411`
  (`Excluir_Linhas_Base_Other_Income`), `Extracao_Base_Consolidad.bas:320-334`
  (`Excluir_Linhas_Base_Consolidada`).
- **Evidência:** as linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a falta de padronização entre módulos (automática vs.
  interativa) é intencional.

### RN-010 — Versão default "Divulgado" no Pré-Closing
- **Nome:** Valor padrão de Versão quando não especificado.
- **Descrição:** Em `Filtrar_Item_Base` (Pré-Closing), se o campo Versão não tiver valor após os demais
  tratamentos, assume-se "Divulgado".
- **Condição:** `Campo_Filtr = "VERSÃO"` e `Valor_Filtr = ""` após os tratamentos anteriores.
- **Ação executada:** `Valor_Filtr = "Divulgado"`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Gerar_Base_Pre_Closing.bas:368-369`.
- **Evidência:** linha 369.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-011 — Filtros binários S/N reaproveitando o valor como nome de campo
- **Nome:** Empresa/Proforma/IFRS_Contábil tratados como flags "S" por coluna dinâmica.
- **Descrição:** No Pré-Closing, ao filtrar por Empresa/Proforma/IFRS_Contábil, o **valor selecionado** vira o
  **nome da coluna a filtrar**, e o critério aplicado é sempre "S" — reflete um layout de planilha onde cada
  Empresa/efeito tem sua própria coluna binária S/N.
- **Condição:** `Campo_Filtr` ∈ {EMPRESA, PROFORMA, IFRS_CONTABIL} e `Valor_Filtr <> ""`.
- **Ação executada:** `Campo_Filtr = Trim(Valor_Filtr)`; `Valor_Filtr = "S"`.
- **Exceções:** se `Valor_Filtr` vazio, cai no default de RN-010 (Proforma/IFRS) — "All effects Proforma"/"All
  effects IFRS" = "S".
- **Local de implementação:** `Gerar_Base_Pre_Closing.bas:341-345, 361-367`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-012 — Zeragem de meses fora do período do cenário (Pré-Closing)
- **Nome:** Isolamento temporal de cada cenário simulado.
- **Descrição:** Após copiar as linhas de um cenário para a pasta temporária, os meses fora do intervalo
  Início/Fim configurado para aquele cenário são zerados, evitando que meses fora do período de simulação
  contaminem o resultado.
- **Condição:** `Mes < Mes_Inicial Or Mes > Mes_Final` do cenário.
- **Ação executada:** zera a faixa de células do mês correspondente na cópia temporária (não na Base original).
- **Exceções:** nenhuma.
- **Local de implementação:** `Gerar_Base_Pre_Closing.bas:163-170, 228-235`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-013 — Visão fixa "IFRS Brasil" nas fontes externas
- **Nome:** Padrão único de Visão para quase todas as extrações.
- **Descrição:** 1009, RGM, MOCKUP_RGM, Other Income, Consolidada, Quick Data e Ajustes gravam
  `VISAO = "IFRS Brasil"` como valor fixo; a variação de Visão (ex. "IFRS Itália") só é tratada depois, pela
  rotina `Gerar_Visao_Italia` (fora do cluster).
- **Condição:** sempre, para essas 7 fontes.
- **Ação executada:** `Form_Preenchimento_Generico(..., "VISAO", "IFRS Brasil")`.
- **Exceções:** Hubble não fixa Visão da mesma forma (vem filtrada dinamicamente da query SQL).
- **Local de implementação:** ex. `Extracao_Base_1009.bas:322-324`, `Extracao_Base_RGM.bas:173-175`.
- **Evidência:** linhas citadas (padrão repetido em 7 módulos).
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-014 — Proforma default "w/o Proforma" (com inconsistência de capitalização)
- **Nome:** Valor padrão de Proforma.
- **Descrição:** A maioria das fontes grava `PROFORMA = "w/o Proforma"`; Ajustes grava
  `"w/o proforma"` (minúsculo); Quick Data não fixa (traz Proforma direto da origem).
- **Condição:** sempre, exceto Quick Data.
- **Ação executada:** `Form_Preenchimento_Generico(..., "PROFORMA", "w/o Proforma")` (ou variante minúscula em
  Ajustes).
- **Exceções:** Quick Data (linha 259-263, comentada — usa valor copiado da origem).
- **Local de implementação:** `Extracao_Base_1009.bas:344-346`, `Extracao_Sheet_Ajustes.bas:214-216` (minúsculo),
  `Extracao_Base_Quick_Data.bas:259-263` (comentado).
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a diferença de capitalização entre módulos causa
  problemas de agrupamento/filtro em relatórios que comparam texto exato.

### RN-015 — Validação estrutural obrigatória (RGM/MOCKUP_RGM/Fixed_Revenues)
- **Nome:** Verificação de integridade da planilha fonte antes de extrair.
- **Descrição:** Antes de copiar valores, cada linha de chave é comparada com o rótulo esperado na planilha de
  origem; divergência dispara `MsgBox` crítico e, em RGM/MOCKUP_RGM, encerra todo o processo Excel com `End`.
- **Condição:** `UCase(Sh_Origem.Cells(Linha,Col_Menu_Origem)) <> UCase(Valor_Verif)`.
- **Ação executada:** `MsgBox` crítico detalhando linha/valor esperado vs. real; em RGM/MOCKUP_RGM, `End`
  abrupto; em Fixed_Revenues, apenas aviso (sem `End`).
- **Exceções:** Fixed_Revenues não executa `End` (inconsistência entre os 3 módulos irmãos).
- **Local de implementação:** `Extracao_Base_RGM.bas:308-334`, `Extracao_Base_MOCKUP_RGM.bas:317-337`,
  `Extracao_Fixed_Revenues.bas:305-329`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma (mas ver risco ALTO correspondente na Seção D).

### RN-016 — Processamento seletivo por tipo de linha de chave ("INF"/"Fórmula")
- **Nome:** Filtro de tipo de linha nas tabelas de chaves De-Para.
- **Descrição:** RGM, MOCKUP_RGM e Fixed_Revenues só processam linhas da aba de chaves marcadas "INF" (valor
  direto) ou "Fórmula" (recalculada dinamicamente); demais linhas (cabeçalhos, comentários) são ignoradas.
- **Condição:** `UCase(Sh_Chaves.Cells(Lin,1)) = "INF" Or "FÓRMULA"`.
- **Ação executada:** processa a linha; senão, pula para a próxima.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Base_RGM.bas:363`, `Extracao_Base_MOCKUP_RGM.bas:364`,
  `Extracao_Fixed_Revenues.bas:344`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-017 — Chave de coluna por concatenação Ano+Mês (Fixed Revenues)
- **Nome:** Localização dinâmica de coluna de cenário via fórmula injetada.
- **Descrição:** Fixed Revenues monta, para os 14 "cenários" configurados, uma chave `AAAAMM` e injeta a
  fórmula `=R[1]C&TEXT(R[2]C,"AAAAMM")` na primeira linha da planilha fonte para localizar dinamicamente a
  coluna correspondente a cada cenário.
- **Condição:** para cada uma das 14 posições de `Lista_Meses_Cenario`.
- **Ação executada:** grava a fórmula; usa `fn.Match` sobre o resultado para achar a coluna.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Fixed_Revenues.bas:82-88, 299`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-018 — Composição lógica OR/AND configurável via tabela de chaves (Hubble)
- **Nome:** Filtros complementares de Empresa/IFRS/Organic combináveis sem alterar código.
- **Descrição:** Para os campos Empresa e IFRS_Contábil, o Hubble monta filtros complementares a partir da
  tabela de chaves `Sheet9`, combinando valores com operador "=" via OR e valores com operador "<>" via AND —
  configurável linha a linha na planilha, sem exigir alteração de VBA.
- **Condição:** presença de colunas adicionais em `Sheet9` além da chave principal.
- **Ação executada:** monta `(<cond1> OR <cond2> ...) AND <cond3> AND ...` dinamicamente.
- **Exceções:** se nenhuma condição complementar existir, usa `[Campo] <> ''` como filtro neutro.
- **Local de implementação:** `Extracao_SQL_Hubble.bas:100-135, 222-250`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-019 — Fontes fora do orquestrador "Extrair Tudo"
- **Nome:** MOCKUP_RGM e Quick_Data não fazem parte da extração em lote.
- **Descrição:** `Extrair_Todas_as_Bases` (`Auxiliar.bas:4-29`, fora do cluster) chama diretamente
  `Processo_Extrair_Base_Hubble`, `_Consolidada`, `_Ajustes`, `_1009`, `_Other_Income`, `_RGM`, `_Fixed_Rev` —
  mas **não** `Processo_Extrair_Base_MOCKUP_RGM` nem `Processo_Extrair_Base_Quick_Data`.
- **Condição:** execução do botão "Extrair Tudo".
- **Ação executada:** as 7 fontes citadas são atualizadas; MOCKUP_RGM e Quick_Data permanecem com os dados da
  última execução manual isolada.
- **Exceções:** nenhuma.
- **Local de implementação:** `Auxiliar.bas:15-21` (lista de chamadas do orquestrador, fora do cluster, mas
  evidência direta do comportamento do cluster).
- **Evidência:** linhas 15-21 de `Auxiliar.bas`.
- **Nível de confiança:** confirmado (ausência verificada por leitura completa da lista de chamadas).
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se é intencional (fontes usadas só em cenários pontuais,
  ex. simulações) ou lacuna de manutenção que deveria ter sido incluída.

### RN-020 — Rastreabilidade de linha original (LIN_BASE) variável por fonte
- **Nome:** Nível de rastreio da linha de origem.
- **Descrição:** A maioria das fontes grava `LIN_BASE = "-"` (sem rastreio granular); Hubble grava o número da
  linha de configuração usada; Fixed_Revenues grava "Linha_Arq_Fixa" da chave (linha real no arquivo fonte).
- **Condição:** sempre, por fonte.
- **Ação executada:** valor fixo "-" ou valor dinâmico conforme a fonte.
- **Exceções:** MOCKUP_RGM e RGM gravam a linha da chave (`Linha_RGM`) na coluna `LIN_BASE`/similar diretamente
  em `Processo_Extracao_Sheet_Base` (não via `Form_Preenchimento_Generico`).
- **Local de implementação:** `Extracao_SQL_Hubble.bas:491` (`Lin AS Lin_Base`), `Extracao_Fixed_Revenues.bas:385`,
  `Extracao_Base_MOCKUP_RGM.bas:401`, `Extracao_Base_1009.bas:287-289` (fixo "-").
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-021 — Organic informado manualmente em Ajustes
- **Nome:** Ajustes não recalculam Organic — usuário informa diretamente.
- **Descrição:** Diferente das demais 8 fontes (que usam `Form_Organic` para calcular via faixa/regra), Ajustes
  copia o campo ORGANIC diretamente da aba de origem (`Sheet13`), assumindo que o analista já preencheu
  corretamente.
- **Condição:** sempre, para a fonte Ajustes.
- **Ação executada:** `Copiar_Base_Ajustes(..., "ORGANIC", "ORGANIC", ...)` em vez de `Form_Organic`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Sheet_Ajustes.bas:196-199`.
- **Evidência:** linhas citadas (ausência de `Call Form_Organic` no módulo).
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-022 — Vazio vira zero explícito em Ajustes
- **Nome:** Ausência de lançamento tratada como zero.
- **Descrição:** Após copiar os 12 meses de Ajustes, células vazias são substituídas por "0" via
  `Selection.Replace`, tratando ausência de lançamento como valor zero explícito.
- **Condição:** célula de valor mensal vazia após a cópia.
- **Ação executada:** `Replace What:="", Replacement:="0"`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Sheet_Ajustes.bas:142-143`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-023 — Tabela de staging individualizada por usuário (Hubble)
- **Nome:** Isolamento de execuções concorrentes via nome de tabela por usuário.
- **Descrição:** O nome da tabela temporária no SQL Server é sufixado com `Environ("UserName")`, permitindo que
  múltiplos analistas executem extrações simultâneas sem colidir no nome da tabela.
- **Condição:** sempre, na extração Hubble.
- **Ação executada:** `TB_AUX_HUBBLE_QUICK_DATA_<usuário>` usado em `CREATE`/`INSERT`/`SELECT`/`DROP`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_SQL_Hubble.bas:399, 431, 486, 541`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma (mas ver risco de tabela "presa" na Seção D).

### RN-024 — Classificação de linha por padrão do primeiro token (Other Income)
- **Nome:** Distinção entre linha de dado e linha de cabeçalho/subtotal.
- **Descrição:** Em Other Income, uma linha só é mantida como dado se o primeiro token da célula de conta
  começar com dígito e não for "Total" nem conter ".", caso contrário é removida (tratada como
  cabeçalho/subtotal/rubrica agregadora).
- **Condição:** avaliação linha a linha do primeiro token da célula.
- **Ação executada:** mantém e extrai o código de conta; ou `.EntireRow.Delete`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Base_Other_Inco.bas:127-153`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-025 — Enriquecimento de descrição de conta apenas em RGM/Fixed Revenues
- **Nome:** Lookup de descrição textual da Classe de Custo.
- **Descrição:** RGM e Fixed_Revenues chamam `PEGA_A_DESCRICAO_DA_CONTA` (`TK_Functions.bas`, fora do cluster)
  para preencher a coluna adjacente à Classe Custo com sua descrição; as demais fontes não fazem esse
  enriquecimento.
- **Condição:** sempre, nesses 2 módulos, dentro de `Processo_Extracao_Sheet_Base`.
- **Ação executada:** `Sh_Destino.Cells(Lin_Destino, Col_Destino_CC+1) = PEGA_A_DESCRICAO_DA_CONTA(...)`.
- **Exceções:** MOCKUP_RGM (mesma estrutura de módulo) **não** chama essa função.
- **Local de implementação:** `Extracao_Base_RGM.bas:391`, `Extracao_Fixed_Revenues.bas:388`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** por que MOCKUP_RGM não recebe o mesmo enriquecimento.

### RN-026 — Zero explícito removido em Consolidada (oposto de RN-022)
- **Nome:** Linhas com Valor Prévia = 0 são descartadas, não mantidas.
- **Descrição:** Em Base Consolidada, linhas cujo "Valor Prévia" é exatamente 0 são removidas da planilha de
  origem antes da cópia — tratamento oposto ao de Ajustes (RN-022), que preserva e zera.
- **Condição:** `Sh_Origem.Cells(Lin,Col_Verif) = 0`.
- **Ação executada:** `.EntireRow.Delete`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Extracao_Base_Consolidad.bas:102-110`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma (a divergência de tratamento entre fontes é factual, não um erro identificado).

### RN-027 — Segmento não recalculado em MOCKUP_RGM/RGM
- **Nome:** Segmento/Abertura_1 vem direto da chave, sem `Form_Segmentos`.
- **Descrição:** Em MOCKUP_RGM e RGM, a chamada a `Form_Segmentos` está comentada — o campo Segmento/Abertura_1
  é gravado diretamente a partir da tabela de chaves (`Sh_Chaves`) dentro de `Processo_Extracao_Sheet_Base`, ao
  contrário de 1009/Other Income/Consolidada/Ajustes/Fixed_Revenues, que recalculam via `Form_Segmentos`.
- **Condição:** sempre, nesses 2 módulos.
- **Ação executada:** `Sh_Destino.Cells(...,Col_Destino_SEGMENTO) = Sh_Chaves.Cells(Lin,Col_Chave_SEGMENTO)`.
- **Exceções:** Fixed_Revenues (mesma família de módulos) **chama** `Form_Segmentos` ativamente (linha 220).
- **Local de implementação:** `Extracao_Base_MOCKUP_RGM.bas:219` (comentada), `Extracao_Base_RGM.bas:217`
  (comentada), vs. `Extracao_Base_MOCKUP_RGM.bas:399`/`Extracao_Base_RGM.bas:396` (gravação direta).
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se a divergência entre os 3 módulos irmãos é intencional.

### RN-028 — Dois modos de limpeza: "Geral" vs. Fonte/Versão específica
- **Nome:** Parametrização do escopo de limpeza histórica.
- **Descrição:** `Limpar_Base_Historica` (fora do cluster, mas invocada por todo o cluster) aceita um `Tipo` que
  pode ser "Geral" (limpa a partir do cabeçalho, todas as linhas) ou um valor específico de Fonte/Versão
  (limpeza seletiva, com reordenação prévia da Base).
- **Condição:** valor do parâmetro `Tipo`/`Chave` na chamada.
- **Ação executada:** limpeza total ou seletiva, conforme o valor.
- **Exceções:** nenhuma.
- **Local de implementação:** `Auxiliar.bas:496-566` (fora do cluster); chamada por todos os
  `Processo_Limpar_Base_*` do cluster e por `Gerar_Base_Versao_Pre_Closing` (com Campo="Versão").
- **Evidência:** ex. `Extracao_SQL_Hubble.bas:37-38` (Campo="Fonte", Chave="Base_Hubble") vs.
  `Auxiliar.bas:12-13` (Campo="Fonte", Valor="Geral").
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-029 — Duas categorias de linha geradas pelo Pré-Closing
- **Nome:** Rastreabilidade de origem dentro da versão de Pré-Closing.
- **Descrição:** O Pré-Closing grava linhas com `Fonte = "QD AUTOMATIC - SELEÇÃO MANUAL"` (cenários definidos
  por coluna de "Preview") e `Fonte = "QD AUTOMATIC - INPUT AJUSTE"` (ajustes vinculados à versão de
  Pré-Closing) — mesmo mecanismo de filtro/cópia, fontes textuais diferentes para auditoria.
- **Condição:** bloco de cenários (linhas 114-187) vs. bloco de ajustes (linhas 190-244).
- **Ação executada:** grava o literal de Fonte correspondente ao bloco.
- **Exceções:** nenhuma.
- **Local de implementação:** `Gerar_Base_Pre_Closing.bas:180, 240`.
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

### RN-030 — Other Income é a única fonte multi-arquivo (pasta)
- **Nome:** Extração baseada em pasta, não em arquivo único.
- **Descrição:** Das 7 fontes externas do cluster, apenas Other Income usa `GetPasta` (seleção de pasta) e
  itera múltiplos arquivos via `Dir()`; as demais 6 usam `GetArquivo` (seleção de arquivo único).
- **Condição:** estrutural — Other Income tem um arquivo por operadora na mesma pasta.
- **Ação executada:** loop `While Arq <> ""` processando cada arquivo encontrado que casa com o critério
  wildcard.
- **Exceções:** nenhuma.
- **Local de implementação:** `Aux_Leitura_Nome_Arqs.bas:38-43` (`Extrair_Diretorio_Arq_Other_Ico` →
  `GetPasta`), `Extracao_Base_Other_Inco.bas:83-232` (loop `Dir()`).
- **Evidência:** linhas citadas.
- **Nível de confiança:** confirmado.
- **Questão pendente:** nenhuma.

---


## 12.2 Cluster Core / Motor de Cálculo (RN-031 a RN-070)
## B. Regras de Negócio (RN-031 a RN-070)

> Faixa reservada a este cluster (motor de cálculo central). Prioridade total ao motor de rateio (`Form_Segmentos`, RN-044 a RN-046, RN-065) e às reclassificações de combinação (RN-056 a RN-059).

---

**RN-031 — Gate de versão obrigatório do Quick Data**
- **Descrição**: o sistema recusa operar se a versão local não corresponder à versão oficial cadastrada centralmente.
- **Condição**: versão retornada por `SELECT VERSAO FROM [BPAM].[dbo].[TB_HUBBLE_VERSAO_FERRAMENTAS] WHERE FERRAMENTA='QUICK DATA'` ≠ `"3.0"` (literal hardcoded no VBA).
- **Ação executada**: exibe `MsgBox` crítico orientando contato com administradores; encerra a execução via `End`.
- **Exceções**: nenhuma — bloqueio incondicional.
- **Local de implementação**: `Auxiliar.bas` — `Verifica_Versao` (linhas 82-117).
- **Evidência (linha)**: Auxiliar.bas:93-101.
- **Nível de confiança**: Fato.
- **Questão pendente**: quem atualiza a tabela `TB_HUBBLE_VERSAO_FERRAMENTAS` e com que frequência? `[VALIDAR COM O NEGÓCIO]`.

**RN-032 — Modo silencioso obrigatório durante processamento em massa**
- **Descrição**: toda rotina pesada desliga ScreenUpdating/DisplayAlerts/EnableEvents e força cálculo manual antes de operar, religando ao final.
- **Condição**: início de qualquer rotina de extração/reclassificação em massa.
- **Ação executada**: `Desligar_Tudo` no início, `Ativar_Tudo` ao final.
- **Exceções**: `Ativar_Tudo` **não** restaura `Application.Calculation` para automático (ver risco D-9).
- **Local de implementação**: `Auxiliar.bas` — `Desligar_Tudo`/`Ativar_Tudo` (linhas 69-80, 128-135).
- **Evidência (linha)**: Auxiliar.bas:130-135, 75-79.
- **Nível de confiança**: Fato.
- **Questão pendente**: é intencional que o cálculo automático não seja restaurado, ou é um bug histórico? `[VALIDAR COM O NEGÓCIO]`.

**RN-033 — Limpeza seletiva de histórico por Fonte antes de nova extração**
- **Descrição**: antes de inserir dados de uma fonte, remove-se o bloco de linhas já existente daquela mesma fonte, garantindo que reprocessar não duplique dados.
- **Condição**: `Campo_Pesq="Fonte"`, `Tipo=<nome da fonte>` (ex. "Base_Hubble", "Base_1009") ou `"Geral"` (limpa tudo).
- **Ação executada**: ordena a Base, localiza o bloco contíguo com `Fonte=Tipo`, deleta via `EntireRow.Delete`.
- **Exceções**: caso especial "BASE_CONSOLIDADA" (RN-034).
- **Local de implementação**: `Auxiliar.bas` — `Limpar_Base_Historica` (linhas 496-564).
- **Evidência (linha)**: Auxiliar.bas:517-561.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-034 — Identificação do bloco "Base_Consolidada" por múltiplos prefixos**
- **Descrição**: a fonte "Base_Consolidada" não tem um único valor de Fonte — é composta por linhas cujo valor começa com "RPD" OU "Preview ", tratadas como um único bloco lógico.
- **Condição**: `Tipo = "BASE_CONSOLIDADA"` em `Limpar_Base_Historica`, ou `Chave="Base_Consolidada"` em `Calcular_Comb_Meses_Intervalo`.
- **Ação executada**: localiza o menor índice entre o primeiro "RPD*" e o primeiro "Preview *"; soma as contagens dos dois prefixos para determinar o fim do bloco.
- **Exceções**: assume implicitamente que, após ordenação, os dois prefixos ficam adjacentes — não validado explicitamente no código.
- **Local de implementação**: `Auxiliar.bas` — `Limpar_Base_Historica` (linhas 544-558), `Calcular_Comb_Meses_Intervalo` (linhas 428-439).
- **Evidência (linha)**: Auxiliar.bas:546-548, 430-439.
- **Nível de confiança**: Fato (existência da regra) / Inferência (garantia de adjacência) — `[VALIDAR COM O NEGÓCIO]`.
- **Questão pendente**: por que "Base_Consolidada" tem 2 prefixos de fonte em vez de um só nome de Fonte padronizado?

**RN-035 — Cálculo de combinação de meses via fórmula com INDIRECT**
- **Descrição**: certas linhas da Base representam um KPI cujo valor mensal é, na verdade, a soma de um intervalo de outras colunas de mês, indicado como texto (ex. "Jan:Mar") em uma célula de referência.
- **Condição**: aplica-se ao bloco de colunas após "Dez" (4 colunas de distância) até o fim do cabeçalho.
- **Ação executada**: `=SUM(INDIRECT(LEFT(R1C,FIND(":",R1C,1)-1)&ROW()&":"&RIGHT(R1C,LEN(R1C)-FIND(":",R1C,1))&ROW()))`, depois convertida em valor fixo.
- **Exceções**: nenhuma condicional — aplica-se a todo o bloco indicado.
- **Local de implementação**: `Auxiliar.bas` — `Calcular_Comb_Meses` (morto), `Calcular_Comb_Meses_Intervalo`, `Calcular_Comb_Meses_Intervalo_Linha`; `TK_Functions.bas` — `Calcular_Comb_Meses_Intervalo_Linha_TK`.
- **Evidência (linha)**: Auxiliar.bas:234, 453, 482.
- **Nível de confiança**: Fato.
- **Questão pendente**: qual o propósito de negócio exato desse mecanismo de "combinação de meses" (que tipo de KPI usa)? `[VALIDAR COM O NEGÓCIO]`.

**RN-036 — Ordenação da Base por 5 chaves antes de localizar blocos por Campo/Chave**
- **Descrição**: para localizar de forma confiável um bloco contíguo de linhas por um critério, a Base inteira é primeiro ordenada por Campo-alvo, KPI, Versão, Exercício e Tipo Nível 2.
- **Condição**: chamada de `Calcular_Comb_Meses_Intervalo`.
- **Ação executada**: `Sort.SortFields` com as 5 chaves em ordem ascendente, `SortMethod=xlPinYin`.
- **Exceções**: nenhuma.
- **Local de implementação**: `Auxiliar.bas` — `Calcular_Comb_Meses_Intervalo` (linhas 404-421).
- **Evidência (linha)**: Auxiliar.bas:396-421.
- **Nível de confiança**: Fato.
- **Questão pendente**: a reordenação da Base como efeito colateral é aceitável do ponto de vista de negócio (referências externas à ordem original)? `[VALIDAR COM O NEGÓCIO]`.

**RN-037 — Geração da visão contábil "IFRS Itália"**
- **Descrição**: determinadas linhas da Base, quando casam com regras de uma tabela de mapeamento dedicada, são duplicadas para formar uma visão contábil alternativa exigida para a entidade italiana do grupo.
- **Condição**: Classe Custo da linha bate exatamente com "Classe Custo" da regra, OU todos os critérios de Abertura_2-8 preenchidos na regra batem com os valores da linha.
- **Ação executada**: copia a linha inteira; se houver multiplicador ≠ 1, aplica-o aos 12 meses; sobrescreve Abertura_2-8 conforme "PARA_A2..PARA_A8"; marca "Visao"="IFRS Itália".
- **Exceções**: regras cujo Classe Custo e todos os critérios de chave estejam vazios são ignoradas (`GoTo Prox_Lin_IT`).
- **Local de implementação**: `Auxiliar.bas` — `Gerar_Visao_Italia` (linhas 245-375).
- **Evidência (linha)**: Auxiliar.bas:316-363.
- **Nível de confiança**: Fato.
- **Questão pendente**: quais linhas de negócio/contas exigem tratamento IFRS Itália e por quê? `[VALIDAR COM O NEGÓCIO]`.

**RN-038 — Cascata de 4 fallbacks para resolver Diretoria Gerencial N1-N3**
- **Descrição**: a Diretoria Gerencial responsável por cada linha é resolvida tentando, em ordem, 4 fontes cada vez mais genéricas.
- **Condição**: aplica-se a todas as linhas com Diretoria N1/N2/N3 Gerencial a calcular.
- **Ação executada**: (1) INDEX/MATCH em Sh_Ref_2 por Classe Custo + Diretoria N1-N3 do Centro de Custo; se erro, (2) INDEX/MATCH em Sh_Ref_1 por concatenação "Grupo BD 3..8"; se erro, (3) INDEX/MATCH em Sh_Ref_1 por Classe Custo simples; se erro, (4) INDEX/MATCH em Sup_Linhas por Centro de Custo.
- **Exceções**: "Grupo BD 3"/"Grupo BD 8" são pulados da concatenação se não existirem no cabeçalho da Base processada.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (linhas 68-149).
- **Evidência (linha)**: Aux_Formulas_Base.bas:134-141.
- **Nível de confiança**: Fato.
- **Questão pendente**: qual a hierarquia de "confiabilidade" de negócio entre as 4 fontes — por que essa ordem específica? `[VALIDAR COM O NEGÓCIO]`.

**RN-039 — Exceção hardcoded: Classe Custo N203073156 + prefixo CDC "NT" → código BD 382**
- **Descrição**: uma combinação específica de Classe de Custo e prefixo de Centro de Custo é forçada a usar o código de agrupamento 382, ignorando a busca normal por Classe Custo.
- **Condição**: `Classe Custo = "N203073156"` E `Left(Centro de Custo, 2) = "NT"`.
- **Ação executada**: em vez de `MATCH(Classe Custo, ...)`, usa `MATCH(382, Sheet2!Descrição_9 ou Sup_Linhas!Linha_BD)`.
- **Exceções**: nenhuma — é a própria regra que é uma exceção a um comportamento padrão.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Grupo_BD` (linhas 678-681), `Form_Opex_Driven` (linhas 712-715), `Form_Linha_BD` (linhas 742-745) — **regra copiada 3 vezes**.
- **Evidência (linha)**: Aux_Formulas_Base.bas:679, 713, 743.
- **Nível de confiança**: Fato (existência) / motivo de negócio `[VALIDAR COM O NEGÓCIO]`.
- **Questão pendente**: por que essa combinação específica precisa de tratamento especial? Documentar antes de tocar em qualquer uma das 3 cópias.

**RN-040 — Classificação CLASSE com exceção "Labour Cost" → "w/o Fiber"**
- **Descrição**: linhas cuja Abertura_3 seja "Labour Cost" recebem CLASSE fixa "w/o Fiber", ignorando o lookup normal por Centro de Custo.
- **Condição**: `Abertura_3 = "Labour Cost"`.
- **Ação executada**: `CLASSE = "w/o Fiber"`; senão, `INDEX/MATCH` normal em Sup_Linhas.
- **Exceções**: é a própria exceção da regra geral.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Classe` (linhas 152-171).
- **Evidência (linha)**: Aux_Formulas_Base.bas:162-163.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que "Labour Cost" mapeia especificamente para "w/o Fiber"? `[VALIDAR COM O NEGÓCIO]`.

**RN-041 — Resolução de EMPRESA via Centro de Custo, sem fallback**
- **Descrição**: cada linha recebe sua Empresa via lookup direto (INDEX/MATCH) do Centro de Custo em Sup_Linhas, sem tratamento de erro.
- **Condição**: toda linha processada por `Form_Empresa`.
- **Ação executada**: `=INDEX(Sup_Linhas!'EMPRESA 2', MATCH(Centro Custo, Sup_Linhas!CENTROdeCUSTO, 0))`.
- **Exceções**: nenhuma — Centro de Custo não encontrado propaga `#N/A`.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Empresa` (linhas 174-192).
- **Evidência (linha)**: Aux_Formulas_Base.bas:183-184.
- **Nível de confiança**: Fato.
- **Questão pendente**: é intencional a ausência de `IFERROR` aqui, diferente de `Form_Classe`/`Form_IFRS_Contabil`? `[VALIDAR COM O NEGÓCIO]`.

**RN-042 — Resolução de Abertura_1 com fallback fixo "Staff"**
- **Descrição**: Abertura_1 é resolvida via VLOOKUP em DP_Segmento (ou "Rateio" se marcado como tal na origem), com fallback para Sup_Linhas e, na ausência de qualquer resultado, fallback fixo "Staff" para empresas FIBER/INTELIG, ou "Total" para as demais.
- **Condição**: linha não pertence a "Base_Ajustes" já preenchida.
- **Ação executada**: `VLOOKUP` em DP_Segmento por Classe Custo+Centro Custo → `IFERROR` → `INDEX/MATCH` em Sup_Linhas → `IFERROR` → `IF(Empresa começa com "FIBER" ou "INTELIG", "Staff", "Total")`.
- **Exceções**: linhas com Fonte="Base_Ajustes" e valor já preenchido são puladas (RN-063).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 213-279).
- **Evidência (linha)**: Aux_Formulas_Base.bas:231-236.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-043 — Resolução de Segmento com fallback fixo "Others"**
- **Descrição**: mesmo padrão de RN-042, para o campo Segmento, com fallback fixo "Others".
- **Condição**: idem RN-042.
- **Ação executada**: idem RN-042, trocando o fallback final para "Others".
- **Exceções**: idem RN-042.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 284-332).
- **Evidência (linha)**: Aux_Formulas_Base.bas:294-300.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-044 — Identificação de linhas elegíveis a rateio**
- **Descrição**: uma linha é candidata ao motor de rateio quando sua Abertura_1 calculada é exatamente "Rateio".
- **Condição**: `UCase(Abertura_1) = UCase("Rateio")`.
- **Ação executada**: dispara o desdobramento de linha (RN-045/RN-046).
- **Exceções**: nenhuma.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linha 472).
- **Evidência (linha)**: Aux_Formulas_Base.bas:472.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-045 — Rateio fixo vs. rateio variável**
- **Descrição**: o percentual de redistribuição de uma linha de rateio é obtido por VLOOKUP em duas tabelas alternativas: "Rateio Fixo" (percentual constante por Empresa+Abertura_1+Linha_BD) ou "Rateio Variavel" (percentual dependente de KPI+Versão+Data+Empresa+Abertura_1+Linha_BD), com o Fixo tentado primeiro.
- **Condição**: linha elegível a rateio (RN-044), avaliada para cada combinação Empresa×Abertura_1 elegível.
- **Ação executada**: `=IFERROR(Valor * VLOOKUP(chave, Sh_Rateio!Rateio_Fixo:+5, 5, 0), IFERROR(Valor * VLOOKUP(chave+data, Sh_Rateio!Rateio_Variavel:+12, 12, 0), 0))`.
- **Exceções**: se nenhuma das duas tabelas tiver correspondência, o valor rateado é 0.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 482, 506-507).
- **Evidência (linha)**: Aux_Formulas_Base.bas:506.
- **Nível de confiança**: Fato.
- **Questão pendente**: qual é a regra de negócio que decide se um item de custo usa rateio fixo ou variável? `[VALIDAR COM O NEGÓCIO]`.

**RN-046 — Linha de exclusão/estorno compensatória do rateio**
- **Descrição**: para cada linha original rateada, é gerada uma linha adicional de "EXCLUSAO" que estorna (zera) o valor original, condicionada à existência de ao menos uma regra de rateio aplicável (fixa, variável ou por revenue).
- **Condição**: pelo menos uma das linhas de rateio geradas para aquela linha original é ≠ 0, OU existe correspondência nas tabelas de rateio fixo/variável/revenue para a chave da linha.
- **Ação executada**: `=IF(OR(condições de rateio aplicável), -Valor_Original, 0)` mês a mês; Fonte marcada com sufixo "- AJUSTE RATEIO VBA - EXCLUSAO".
- **Exceções**: se nenhuma condição de rateio aplicável for satisfeita, a linha de exclusão fica com valor 0 (efetivamente neutra).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 517-539).
- **Evidência (linha)**: Aux_Formulas_Base.bas:530, 535.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-047 — Filtro de universo de rateio via 5 Slicers**
- **Descrição**: antes de calcular o rateio, o sistema restringe o universo de dados considerado ajustando o estado de 5 segmentações (Slicers) do workbook.
- **Condição**: execução de `Form_Segmentos`.
- **Ação executada**: `Slicer_EMPRESA` restrito a "INTELIG ONGOING"; `Slicer_ABERTURA_1` exclui "TIM FIXO"/"VOIP"; `Slicer_ABERTURA_2` sem filtro ativo (regra de exclusão de "REVENUES" está comentada/desativada); `Slicer_LINHA_BD` restrito a códigos numéricos de 3 dígitos entre 159-240 ou igual a 265; `Slicer_IFRS_CONTABIL` exclui "IFRS 9"/"IFRS 15"/"IFRS 16".
- **Exceções**: a exclusão de "REVENUES" em Abertura_2 está **desativada** (código comentado) — só os outros 4 filtros estão ativos.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 381-424).
- **Evidência (linha)**: Aux_Formulas_Base.bas:384-386, 389-395, 397-403 (comentado), 405-416, 418-424.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que a exclusão de Revenues foi desativada — decisão de negócio ou esquecimento? `[VALIDAR COM O NEGÓCIO]`.

**RN-048 — Ajuste de sinal de Revenues**
- **Descrição**: valores de linhas de receita têm o sinal padronizado (invertido) para manter convenção contábil consistente entre receita e despesa.
- **Condição**: `Abertura_2 = "REVENUES"`.
- **Ação executada**: `mês = mês * -1` para as 12 colunas de mês.
- **Exceções**: apenas o fluxo "Other Income" aplica esta correção (único chamador encontrado).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Acertar_Sinal_Revenues` (linhas 624-649).
- **Evidência (linha)**: Aux_Formulas_Base.bas:640.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que só a base "Other Income" precisa desse ajuste de sinal, e as demais não? `[VALIDAR COM O NEGÓCIO]`.

**RN-049 — Cálculo do FY como soma Jan-Dez**
- **Descrição**: o total anual de cada linha é sempre a soma simples dos 12 meses.
- **Condição**: aplicado a toda linha inserida na Base, por praticamente todos os módulos de extração.
- **Ação executada**: `FY = SUM(Jan:Dez)`.
- **Exceções**: nenhuma.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Calcular_FY` (linhas 22-37).
- **Evidência (linha)**: Aux_Formulas_Base.bas:31.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-050 — Zeramento de meses exceto um**
- **Descrição**: para bases pontuais (evento em um único mês), todos os demais meses são forçados a zero.
- **Condição**: chamada explícita com um `Mes_Excessao`.
- **Ação executada**: `mês = 0` para todos os meses exceto o informado.
- **Exceções**: se `Mes_Excessao` não corresponder a nenhum mês válido, todos os 12 são zerados.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Zerar_Meses_Exceto_Um` (linhas 39-65).
- **Evidência (linha)**: Aux_Formulas_Base.bas:58-61.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-051 — Ajuste de escala (multiplicador)**
- **Descrição**: valores de determinadas fontes são multiplicados por um fator de escala (ex.: milhares → unidades) antes de compor a Base final.
- **Condição**: chamada explícita com um `Multiplicador` pelos módulos Fixed Revenues, RGM, MOCKUP_RGM.
- **Ação executada**: `mês = mês * Multiplicador` para os 12 meses.
- **Exceções**: nenhuma.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Acertar_Escala` (linhas 924-949).
- **Evidência (linha)**: Aux_Formulas_Base.bas:939-940.
- **Nível de confiança**: Fato.
- **Questão pendente**: qual o valor do multiplicador em cada fonte, e por que essas 3 fontes especificamente precisam de ajuste de escala? `[VALIDAR COM O NEGÓCIO]`.

**RN-052 — Classificação Organic (hoje sempre "Ajusted")**
- **Descrição**: a coluna ORGANIC é preenchida com a constante "Ajusted" para todas as linhas processadas — uma regra condicional anterior (baseada em Classe Custo específica) foi substituída por essa constante.
- **Condição**: nenhuma — aplica-se incondicionalmente onde a rotina é chamada.
- **Ação executada**: `ORGANIC = "Ajusted"`.
- **Exceções**: base de Ajustes não passa por esta rotina (chamada comentada em Extracao_Sheet_Ajustes.bas).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Organic` (linhas 771-786).
- **Evidência (linha)**: Aux_Formulas_Base.bas:777-778.
- **Nível de confiança**: Fato (comportamento atual) — mudança de regra não documentada, `[VALIDAR COM O NEGÓCIO]`.
- **Questão pendente**: a regra condicional anterior (`N904015189`→"Reported") deveria ainda estar ativa? Confirmar com o negócio antes de reescrever.

**RN-053 — Peso REF_ORGANIC (1-4) a partir de ORGANIC**
- **Descrição**: atribui peso numérico a cada categoria de ORGANIC, para uso em ordenação/priorização de relatórios.
- **Condição**: `ORGANIC = "Reported"` → 1; `"Recurrent"` → 3; `"Normalized"` → 4; qualquer outro (inclusive "Ajusted") → 2.
- **Ação executada**: fórmula `IF` aninhada, bake-in em valor.
- **Exceções**: nenhuma — todo valor não-mapeado cai no peso 2.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Ref_Organic` (linhas 788-802).
- **Evidência (linha)**: Aux_Formulas_Base.bas:794.
- **Nível de confiança**: Fato.
- **Questão pendente**: como "Reported"/"Recurrent"/"Normalized" são hoje atingidos, se `Form_Organic` só produz "Ajusted"? `[VALIDAR COM O NEGÓCIO]`.

**RN-054 — Classificação IFRS Contábil com fallback "w/o IFRS"**
- **Descrição**: cada linha é classificada quanto à norma IFRS aplicável via VLOOKUP posicional em tabela de contas; se não encontrado, recebe "w/o IFRS".
- **Condição**: toda linha processada por `Form_IFRS_Contabil`.
- **Ação executada**: `=IFERROR(VLOOKUP(Classe Custo, Sheet2!'Contas IFRS':+2, 3, 0), "w/o IFRS")`.
- **Exceções**: nenhuma.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_IFRS_Contabil` (linhas 753-769).
- **Evidência (linha)**: Aux_Formulas_Base.bas:761.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-055 — Grupo BD / Linha_BD / Abertura_2-8 via lookup por Classe Custo**
- **Descrição**: toda a árvore de classificação de P&L (Grupo BD 2-8, Abertura_2-8, Linha_BD) é resolvida por lookup posicional de Classe de Custo nas tabelas Sheet2/Sup_Linhas, sujeita à exceção RN-039.
- **Condição**: toda linha processada por `Form_Grupo_BD`/`Form_Opex_Driven`/`Form_Linha_BD`.
- **Ação executada**: `=IFERROR(INDEX(tabela_Descricao_x, MATCH(Classe Custo, tabela_Conta, 0)), "-")` (com a exceção 382 substituindo o MATCH quando aplicável).
- **Exceções**: RN-039; campos "Grupo BD 3"/"Grupo BD 8" são pulados se não existirem no cabeçalho da Base processada.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Grupo_BD` (652-690), `Form_Opex_Driven` (693-724), `Form_Linha_BD` (727-751).
- **Evidência (linha)**: Aux_Formulas_Base.bas:678-681.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-056 — Motor de reclassificação de combinações (Empresa/IFRS/Proforma)**
- **Descrição**: cada linha da Base recebe, para cada "combinação" pré-definida em DropComb, uma marca "S"/"N" indicando pertencimento, calculada por uma cadeia de `IF(AND(...),"S", ...)` montada dinamicamente a partir da matriz de critérios.
- **Condição**: existe uma linha de critério em DropComb (Sheet9) para cada combinação (colunas prefixadas "Empresa*"/"IFRS_Contabil*"/"Proforma*").
- **Ação executada**: monta e aplica a fórmula ao bloco de colunas correspondente; bake-in em valor.
- **Exceções**: linhas de critério que sejam separadores visuais ("------------"/"============") são ignoradas (não vale para Empresas, que não tem essa checagem).
- **Local de implementação**: `Auxiliar.bas` — `Reclassificar_Combinacoes_Empresas/IFRS_Contabil/Proforma` (606-936); duplicado em `TK_Functions.bas` com sufixo `_TK` (877-953, 716-794, 797-875).
- **Evidência (linha)**: Auxiliar.bas:667, 830, 909 (fórmula ELSE de cada uma).
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma quanto ao mecanismo; ver RN-059/D-1 quanto à duplicação.

**RN-057 — Critério "ELSE" implícito nas reclassificações**
- **Descrição**: se uma linha da Base não bate com nenhum critério explícito da matriz DropComb, a última cláusula da cadeia de `IF` compara o cabeçalho da coluna de combinação diretamente ao valor da linha (autoteste "a combinação tem o mesmo nome que o valor da linha?").
- **Condição**: nenhum critério anterior da cadeia `IF(AND(...))` casou.
- **Ação executada**: `IF(R<LIN_BASE>C = RC<coluna do campo>, "S", "N")` como cláusula final.
- **Exceções**: nenhuma — é o próprio fallback universal.
- **Local de implementação**: `Auxiliar.bas` — `Reclassificar_Combinacoes_Empresas/IFRS_Contabil/Proforma` (linhas 667, 830, 909); replicado nas versões `_TK`.
- **Evidência (linha)**: Auxiliar.bas:667.
- **Nível de confiança**: Fato.
- **Questão pendente**: este fallback é a regra de negócio pretendida ou um "catch-all" técnico sem significado de negócio? `[VALIDAR COM O NEGÓCIO]`.

**RN-058 — Critérios de exclusão "<>" isolados em AND, separados dos de igualdade em OR**
- **Descrição**: dentro de uma mesma combinação, critérios de exclusão explícita (prefixo "<>") são exigidos simultaneamente (AND), enquanto critérios de igualdade (sem "<>") bastam que um seja satisfeito (OR).
- **Condição**: item de filtro contém "<>" (exclusão) vs. não contém (igualdade).
- **Ação executada**: monta `Complemento_1` (AND de exclusões) e `Complemento_2` (OR de igualdades), combinados na cláusula final `AND(critério_principal, Complemento_1, Complemento_2)`.
- **Exceções**: itens contendo "TRIM" não recebem aspas automáticas (tratados como expressão, não literal).
- **Local de implementação**: `Auxiliar.bas` — `Criar_Formula_Filtros` (linhas 939-971).
- **Evidência (linha)**: Auxiliar.bas:949-965.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-059 — Divisão em 2 partes da fórmula de reclassificação de Empresas**
- **Descrição**: apenas a reclassificação de Empresas (não IFRS nem Proforma, e não a versão `_TK` de Empresas) divide a matriz de combinações ao meio e aplica 2 fórmulas separadas a 2 blocos de colunas distintos.
- **Condição**: sempre, ao executar `Reclassificar_Combinacoes_Empresas` (Auxiliar.bas).
- **Ação executada**: localiza a combinação do meio; monta `Formula_Empresa_PARTE_1` (até o meio) e `_PARTE_2` (restante); aplica cada uma ao seu bloco de colunas.
- **Exceções**: a versão `_TK` equivalente (`Reclassificar_Combinacoes_Empresas_TK`) **não** replica essa divisão.
- **Local de implementação**: `Auxiliar.bas` — `Reclassificar_Combinacoes_Empresas` (linhas 625-763).
- **Evidência (linha)**: Auxiliar.bas:634, 757-763; ausência confirmada em TK_Functions.bas:877-953.
- **Nível de confiança**: Fato (existência da divisão) / Inferência (motivo = limite de tamanho/performance de fórmula) `[VALIDAR COM O NEGÓCIO]`.
- **Questão pendente**: a divisão foi necessária para contornar um limite técnico (tamanho de fórmula, tempo de cálculo)? Se sim, a versão `_TK` sem essa divisão está exposta ao mesmo problema.

**RN-060 — Preenchimento automático de CDC ausente via referência SQL, com marcação de cor**
- **Descrição**: quando uma linha não traz Centro de Custo da fonte original ("-"), o sistema busca um valor de referência em uma tabela DE-PARA alimentada via SQL Server, preenche se encontrado, e sinaliza visualmente (cor) a origem do dado.
- **Condição**: `Centro de Custo = "-"`.
- **Ação executada**: monta chave (colunas 31+32+33+8); busca em Sh11; se achado, preenche e colore "com CDC por referencia" (RGB 49407); se não achado, colore "sem CDC" (tema escuro); se já tinha valor, colore "com CDC Real" (RGB 5287936).
- **Exceções**: nenhuma.
- **Local de implementação**: `TK_Functions.bas` — `UPDATE_aplicar_CDC_por_Referencia` (521-574), `Extrair_Base_CDCs_DE_PARA` (585-645), `SET_Cor_CdC` (647-674).
- **Evidência (linha)**: TK_Functions.bas:548-566.
- **Nível de confiança**: Fato.
- **Questão pendente**: procedimento sem chamador identificado no dump — é acionado manualmente/por botão? `[NÃO ACESSÍVEL]`.

**RN-061 — Idempotência de extração por deleção prévia do bloco da fonte**
- **Descrição**: princípio geral que percorre todo o sistema — nenhuma extração de fonte insere dados sem antes remover qualquer bloco pré-existente daquela mesma fonte (RN-033), evitando duplicação ao reprocessar.
- **Condição**: início de qualquer `Processo_Extrair_Base_*`.
- **Ação executada**: `Call Limpar_Base_Historica(Campo, Chave)` antes da inserção.
- **Exceções**: nenhuma identificada.
- **Local de implementação**: padrão observado em todos os módulos `Extracao_Base_*.bas` (fora deste cluster) via `Limpar_Base_Historica` (Auxiliar.bas).
- **Evidência (linha)**: Auxiliar.bas:496-564 (implementação); chamadores em >10 módulos de extração.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-062 — Tratamento de erro por linha em Form_Segmentos**
- **Descrição**: se o cálculo de Abertura_1 de uma linha específica resultar em erro de fórmula, o sistema avisa o usuário via `MsgBox` e pula essa linha, sem interromper o processamento das demais.
- **Condição**: `VarType(Sh_Destino.Cells(Lin, Col_INFO)) = vbError`.
- **Ação executada**: `MsgBox "Atenção, erro encontrado na linha..."`; `GoTo proximaLinha`.
- **Exceções**: nenhuma — toda linha com erro é pulada da mesma forma.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 464-469).
- **Evidência (linha)**: Aux_Formulas_Base.bas:465-468.
- **Nível de confiança**: Fato.
- **Questão pendente**: o log estruturado (`fn_ListAllErrors`) foi planejado para substituir este `MsgBox` mas está desconectado (chamada comentada) — reconectar antes de tornar o processo desatendido/batch.

**RN-063 — Exceção de preenchimento para linhas de Fonte="Base_Ajustes"**
- **Descrição**: linhas cuja Fonte é "Base_Ajustes" e que já têm um valor preenchido em Abertura_1/Segmento não são sobrescritas pela fórmula automática — presume-se que foram ajustadas manualmente e devem ser preservadas.
- **Condição**: `Sh_Destino.Cells(linha, Col_INFO).Value <> "" AND Sh_Destino.Cells(linha, 2).Value = "Base_Ajustes"`.
- **Ação executada**: `GoTo proximaLinha__`/`proximaLinha___` (pula a linha sem alterar).
- **Exceções**: aplica-se apenas a Abertura_1 e Segmento — outros campos calculados (Classe, Grupo BD etc.) não têm essa exceção.
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 241, 314).
- **Evidência (linha)**: Aux_Formulas_Base.bas:241, 314.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que a exceção cobre só Abertura_1/Segmento e não os demais campos calculados da Base_Ajustes? `[VALIDAR COM O NEGÓCIO]`.

**RN-064 — Padronização "Live" → "UBB" pós-cálculo**
- **Descrição**: após calcular e converter em valor a coluna Abertura_1, o sistema substitui todas as ocorrências textuais de "Live" por "UBB" no mesmo intervalo.
- **Condição**: aplicado sempre, ao final do bloco de cálculo de Abertura_1 em `Form_Segmentos`.
- **Ação executada**: `Range.Replace What:="Live", Replacement:="UBB"`.
- **Exceções**: aplicado apenas à coluna Abertura_1 (não ao bloco de Segmento, que tem um `Replace` equivalente comentado/desativado, linha 337 comentada).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linha 258).
- **Evidência (linha)**: Aux_Formulas_Base.bas:258.
- **Nível de confiança**: Fato.
- **Questão pendente**: qual o significado de negócio de "Live"/"UBB" (nomenclatura de produto/tecnologia)? `[VALIDAR COM O NEGÓCIO]`.

**RN-065 — Elegibilidade de rateio por Empresa presente em Lista_Empresa_A1**
- **Descrição**: uma linha de rateio só é desdobrada para as combinações Empresa×Abertura_1 que estejam na lista de elegíveis, construída a partir de duas fontes: o "Index A1" de Sh_Rateio (excluindo marcador "-") e o "Rateio Fixo" com percentual ≠ 0.
- **Condição**: `Sh_Rateio.Cells(Lin, Col_Lista_A1) <> "-"` OU `Sh_Rateio.Cells(Lin, Col_Rateio_Fixo+3) <> 0`.
- **Ação executada**: monta `Lista_Empresa_A1` (chaves únicas "Empresa > Abertura_1"); para cada linha de rateio, gera uma linha nova apenas para as combinações da lista cuja Empresa bate com a da linha original.
- **Exceções**: combinações duplicadas (já presentes na lista) não são adicionadas de novo (checagem via `InStr`).
- **Local de implementação**: `Aux_Formulas_Base.bas` — `Form_Segmentos` (linhas 433-455, 484-513).
- **Evidência (linha)**: Aux_Formulas_Base.bas:439-455, 489.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-066 — Tratamento IFRS16: substituição de CC "DE" por CC "PARA" com estorno**
- **Descrição**: linhas identificadas para tratamento IFRS16 (leasing) têm seu Centro/Classe de Custo original ("DE") substituído por um Centro/Classe de Custo alvo ("PARA") mapeado em Sup_Linhas, e os valores são multiplicados por -1 antes de serem reinseridos na Base como novas linhas.
- **Condição**: `Sh_Extracao.Range("I25") = "Sim"` (flag de ativação do tratamento); linha da Base cuja Fonte começa com o prefixo informado (`vBase_a_procurar`) e cujo CC "DE" existe em Sup_Linhas (coluna 71).
- **Ação executada**: copia a linha para a aba auxiliar IFRS16 (Sheet28); substitui CC pelo valor "PARA" (coluna 73 de Sup_Linhas); para Empresa ≠ "5G"/"METIS_CZ", também resolve CDC "PARA" via chave "CTCEL"+chave composta; multiplica valores por -1 (via `Sup_Linhas!BX1=-1` colado com operação Multiply); reinsere linhas na Base; marca linhas originais como "DELETAR" e remove via `Limpar_Base_Historica`.
- **Exceções**: Empresa "5G" ou "METIS_CZ" não recebe substituição de CDC (só de CC); linhas com CC/CDC/Empresa/Chave ausentes (`#N/A`) são puladas e contadas (RN-067).
- **Local de implementação**: `fx_IFRS16.bas` — `UPDATE_Tratar_IFRS16` (linhas 3-141, fora deste cluster, mas dispara diretamente os procedimentos `_TK` deste cluster).
- **Evidência (linha)**: fx_IFRS16.bas:29, 53, 62-68, 77-84, 103-105.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que "5G"/"METIS_CZ" são excluídas da substituição de CDC? `[VALIDAR COM O NEGÓCIO]`.

**RN-067 — IFRS16: linhas com dado ausente são puladas e contabilizadas**
- **Descrição**: linhas candidatas ao tratamento IFRS16 cujas colunas-chave (CC, CDC, Empresa, ou parte da chave composta) contenham erro `#N/A` são ignoradas, mas contadas para informar o usuário ao final.
- **Condição**: `IsNA(coluna 23) OR IsNA(coluna 35) OR IsNA(coluna 8) OR IsNA(coluna 31)`.
- **Ação executada**: incrementa contador `xx`; `GoTo proximaLinha` (pula o tratamento desta linha).
- **Exceções**: nenhuma.
- **Local de implementação**: `fx_IFRS16.bas` — `UPDATE_Tratar_IFRS16` (linhas 40-43, 133-137).
- **Evidência (linha)**: fx_IFRS16.bas:40-43, 134.
- **Nível de confiança**: Fato.
- **Questão pendente**: nenhuma.

**RN-068 — Sincronização da grade de combinações a partir de DropComb**
- **Descrição**: a lista de "slots" de combinação de Empresa/Proforma exibida no cabeçalho da Base é sincronizada manualmente a partir da matriz de manutenção DropComb, via ação explícita do usuário (ComboBox na aba Extracao).
- **Condição**: usuário seleciona a opção "Combinações" no `ComboBox1` de Sheet8 (aba Extracao).
- **Ação executada**: insere/remove colunas de slot na Base conforme diferença de quantidade; copia nomes de combinação para a linha de cabeçalho.
- **Exceções**: a sincronização de Proforma existe implementada mas não está conectada a nenhuma opção do ComboBox (item 59) — só é acionada via IFRS16.
- **Local de implementação**: `TK_Functions.bas` — `UPDATE_Combinacoes_Empresas` (115-204), `UPDATE_Combinacoes_Proforma` (206-282, órfã).
- **Evidência (linha)**: TK_Functions.bas:142-164; Sheet8.cls:14-15.
- **Nível de confiança**: Fato.
- **Questão pendente**: por que não há opção de sincronização de Proforma no ComboBox, dado que a lógica já existe pronta? `[VALIDAR COM O NEGÓCIO]`.

**RN-069 — Recalcular reclassificação apenas se detectada mudança real**
- **Descrição**: após sincronizar o cabeçalho de combinações, a reclassificação em massa (`_TK`) só é disparada se a comparação de células em branco antes/depois indicar que algo mudou — evita recálculo caro quando a sincronização não alterou nada de fato.
- **Condição**: `qtde_cells_blank <> qtde_cells_selected` (heurística de contagem de células vazias no bloco de dados).
- **Ação executada**: chama `Reclassificar_Combinacoes_Empresas_TK`/`_Proforma_TK`.
- **Exceções**: a heurística não detecta mudança de *conteúdo* dos critérios quando a *quantidade* de combinações permanece igual — nesse caso a reclassificação não é disparada mesmo que devesse ser.
- **Local de implementação**: `TK_Functions.bas` — `UPDATE_Combinacoes_Empresas` (linhas 179-181), `UPDATE_Combinacoes_Proforma` (linhas 266-268).
- **Evidência (linha)**: TK_Functions.bas:179-181, 266-268.
- **Nível de confiança**: Fato.
- **Questão pendente**: ver risco associado em D — a heurística é uma aproximação, não uma comparação de conteúdo real.

**RN-070 — Limite de linha hardcoded (40.837) em atualização manual de combinações**
- **Descrição**: a propagação manual (AutoFill) de fórmulas de combinação de Empresa cobre um range fixo até a linha 40.837 da Base, independentemente do tamanho real dos dados.
- **Condição**: execução de `UPDATE_Combinacoes_Empresas_parte2`.
- **Ação executada**: `Selection.AutoFill Destination:=Range("EC6:GD40837")`.
- **Exceções**: nenhuma — linhas além de 40.837 nunca são preenchidas por esta rotina, sem erro ou aviso.
- **Local de implementação**: `TK_Functions.bas` — `UPDATE_Combinacoes_Empresas_parte2` (linhas 315-340).
- **Evidência (linha)**: TK_Functions.bas:321.
- **Nível de confiança**: Fato.
- **Questão pendente**: a Base já ultrapassou 40.837 linhas em produção? Se sim, esta rotina (se ainda em uso manual) já está gerando dados incompletos silenciosamente. `[VALIDAR COM O NEGÓCIO]` com urgência — ver risco crítico D-2.

---


## 12.3 Cluster Refresh / Validação / IFRS16 (RN-071 a RN-096)
# B) REGRAS DE NEGÓCIO (RN-071 a RN-096)

> Numeração reservada RN-071–RN-100 (26 regras documentadas, intervalo não excedido). Prioridade máxima às regras contábeis de IFRS16 e à cascata de Ref Cruzada, conforme solicitado.

### RN-071 — Escopo condicional do tratamento IFRS16
- **Descrição:** O tratamento contábil de IFRS16 só é executado se a extração estiver configurada para tal.
- **Condição:** `Sh_Extracao!I25 (Sheet8!I25) = "Sim"`.
- **Ação:** executa `UPDATE_Tratar_IFRS16`; senão, executa o caminho alternativo de reclassificação "não-TK" (ver RN-081).
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:12`; complementar em `Extracao_Base_1009.bas:395`.
- **Evidência:** `fx_IFRS16.bas:12`, `Extracao_Base_1009.bas:395`.
- **Confiança:** Confirmado (leitura direta).
- **Questão pendente:** nenhuma.

### RN-072 — Filtro de linhas-alvo do IFRS16 por prefixo de origem
- **Descrição:** Só linhas cujo campo "Fonte" (coluna 2) comece com o valor de `vBase_a_procurar` são candidatas ao tratamento IFRS16.
- **Condição:** `vLinha.Columns(2) Like vBase_a_procurar & "*"`.
- **Ação:** inclui a linha no processamento de IFRS16.
- **Exceções:** nenhuma; no fluxo real observado, `vBase_a_procurar = "Base_1009"` (via `Extracao_Base_1009.bas:5,16`).
- **Local de implementação:** `fx_IFRS16.bas:25,36`.
- **Evidência:** `fx_IFRS16.bas:25,36`; `Extracao_Base_1009.bas:5,16`.
- **Confiança:** Confirmado.
- **Questão pendente:** confirmar se, historicamente, o tratamento já rodou para outras origens (a chamada em `Extracao_SQL_Hubble.bas:16` está comentada — ver RN-081).

### RN-073 — De-Para de Classe Custo (CC) para linhas IFRS16
- **Descrição:** O Centro/Classe de Custo original ("CC DE", coluna 23) de uma linha elegível é substituído pelo Centro/Classe de Custo de destino ("CC PARA") cadastrado em `Sup_Linhas`.
- **Condição:** `vCC_DE` (coluna 23) encontrado em `Sup_Linhas.Columns(71)` via `Match`.
- **Ação:** `vCC_PARA = Sup_Linhas.Cells(vLin_Alvo, 73)`; grava em `Sh_AuxIFRS16` coluna 23.
- **Exceções:** nenhuma condição de negócio adicional — depende só da existência do mapeamento.
- **Local de implementação:** `fx_IFRS16.bas:49,53,62`.
- **Evidência:** `fx_IFRS16.bas:26-27,45,49,53,62`.
- **Confiança:** Confirmado.
- **Questão pendente:** quem mantém `Sup_Linhas` colunas 71/73 (ver risco crítico, seção D).

### RN-074 — Linha sem De-Para de CC é excluída do tratamento
- **Descrição:** Se o CC "DE" não for encontrado em `Sup_Linhas.Columns(71)`, a linha é ignorada (não tratada) e contabilizada em um contador de erro.
- **Condição:** `IsError(Match(vCC_DE, Sup_Linhas.Columns(71), 0))` (via `GoTo proximaLinha`).
- **Ação:** pula a linha (`xx = xx + 1`); ao final, se `xx > 0`, informa ao usuário via MsgBox quantas linhas foram puladas.
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:40-43,133-134`.
- **Evidência:** `fx_IFRS16.bas:40-43,133-137`.
- **Confiança:** Confirmado.
- **Questão pendente:** não há log de qual CC específico falhou — apenas contagem total (risco de auditoria, ver seção D).

### RN-075 — Marcação da linha original como "DELETAR"
- **Descrição:** Toda linha elegível para IFRS16 (mesmo antes de confirmar o De-Para completo) tem a coluna "Fonte" marcada como "DELETAR" na primeira passada, para posterior exclusão.
- **Condição:** CC "DE" encontrado em `Sup_Linhas.Columns(71)` (usando, nesta 1ª passada, `Sup_Linhas.Columns(73)` como referência, `:27`).
- **Ação:** `Sh_Base.Range("B" & vLinha.Row).Value = "DELETAR"`.
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:24-33`.
- **Evidência:** `fx_IFRS16.bas:29`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-076 — Sufixo de rastreabilidade "_IFRS16 Tratado"
- **Descrição:** A cópia da linha tratada recebe um sufixo no campo "Fonte" para indicar que já passou pelo tratamento IFRS16, evitando reprocessamento em execuções futuras.
- **Condição:** linha copiada para `Sh_AuxIFRS16` durante o tratamento.
- **Ação:** `Sh_AuxIFRS16.Cells(...,2) = Sh_AuxIFRS16.Cells(...,2).Value & "_IFRS16 Tratado"`.
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:63`.
- **Evidência:** `fx_IFRS16.bas:63`; confirmado também pelo uso da string `"Base_1009_IFRS16 Tratado"` em `Extracao_Base_1009.bas:37` (rotina de limpeza que remove esse rastro antes de reextrair).
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-077 — De-Para de Centro de Custo (CDC) condicional por Empresa
- **Descrição:** O remapeamento de Centro de Custo (CDC) só é aplicado a empresas que não sejam "5G" nem "METIS_CZ" — essas duas ficam fora do tratamento de CDC do IFRS16 (mesmo estando dentro do tratamento de CC).
- **Condição:** `vEmpresa <> "5G" AND vEmpresa <> "METIS_CZ"`.
- **Ação:** resolve `vCDC_PARA` via `Sup_Linhas.Columns(15)`/`22` e grava em `Sh_AuxIFRS16` coluna 35.
- **Exceções:** Empresas "5G" e "METIS_CZ" — mantêm o CDC original.
- **Local de implementação:** `fx_IFRS16.bas:65-69`.
- **Evidência:** `fx_IFRS16.bas:65`.
- **Confiança:** Confirmado (regra explícita no código).
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** — motivo de negócio para a exceção não está documentado no código; presumível que 5G/METIS_CZ tenham estrutura de CDC própria não coberta pela tabela `Sup_Linhas` no formato "CTCEL"+chave.

### RN-078 — Chave de busca do CDC de destino
- **Descrição:** A busca do CDC de destino usa uma chave composta por um prefixo fixo "CTCEL" concatenado com a Diretoria N1+N2+N3 da linha.
- **Condição:** aplicável apenas quando RN-077 permite o remapeamento.
- **Ação:** `vChave = Columns(31)&Columns(32)&Columns(33)`; `Match("CTCEL" & vChave, Sup_Linhas.Columns(15), 0)`.
- **Exceções:** nenhuma além de RN-077.
- **Local de implementação:** `fx_IFRS16.bas:48,66-67`.
- **Evidência:** `fx_IFRS16.bas:48,66`.
- **Confiança:** Confirmado.
- **Questão pendente:** o prefixo fixo `"CTCEL"` sugere uma convenção de nomenclatura de sistema de origem (SAP?) não documentada — **[VALIDAR COM O NEGÓCIO]** o significado exato de "CTCEL".

### RN-079 — Inversão de sinal dos valores IFRS16 tratados
- **Descrição:** Os valores mensais (colunas de meses, a partir de "AJ") das linhas tratadas por IFRS16 têm o sinal invertido em relação à base original — regra clássica de reclassificação contábil (retirar de uma linha, reclassificar em outra, com estorno).
- **Condição:** aplicável a toda linha copiada para `Sh_AuxIFRS16` no fluxo de tratamento.
- **Ação:** multiplica os valores por `-1`, lido de `Sup_Linhas!BX1` (setado para `-1` no próprio código antes de colar) via `PasteSpecial Operation:=xlMultiply`.
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:76-86`.
- **Evidência:** `fx_IFRS16.bas:77-78,83-84`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma quanto ao mecanismo; **[VALIDAR COM O NEGÓCIO]** a justificativa contábil detalhada (qual lançamento de origem e qual de destino) para documentação formal do 24-seções.

### RN-080 — Substituição (não duplicação) das linhas originais
- **Descrição:** As linhas tratadas por IFRS16 substituem as originais na base histórica — a operação é conceitualmente um "mover", nunca uma duplicação de saldo.
- **Condição:** após reinserção das linhas tratadas ao final de `Sh_Base`.
- **Ação:** `Limpar_Base_Historica(Campo:="Fonte", Chave:="DELETAR")` remove definitivamente as linhas marcadas na RN-075.
- **Exceções:** nenhuma.
- **Local de implementação:** `fx_IFRS16.bas:90-105`.
- **Evidência:** `fx_IFRS16.bas:103-105`, `Auxiliar.bas:496` (definição de `Limpar_Base_Historica`).
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-081 — Exclusividade entre fluxo IFRS16 e fluxo de reclassificação padrão
- **Descrição:** Quando `I25 = "Sim"`, a reclassificação de combinações (Empresas/IFRS_Contabil/Proforma/Meses) ocorre **dentro** de `UPDATE_Tratar_IFRS16`, usando as variantes "_TK". Quando `I25 = "Não"`, a mesma reclassificação ocorre **imediatamente após a extração**, usando as variantes sem sufixo "_TK" — os dois caminhos não coexistem na mesma execução.
- **Condição:** valor de `Sheet8!I25`.
- **Ação:** ramo "Sim" → `fx_IFRS16.bas:125-128` (`Reclassificar_Combinacoes_Empresas_TK`, `Calcular_Comb_Meses_Intervalo_Linha_TK`, `Reclassificar_Combinacoes_Proforma_TK`, `Reclassificar_Combinacoes_IFRS_Contabil_TK`); ramo "Não" → `Extracao_Base_1009.bas:397-400` (`Reclassificar_Combinacoes_Empresas`, `Reclassificar_Combinacoes_IFRS_Contabil`, `Reclassificar_Combinacoes_Proforma`, `Calcular_Comb_Meses_Intervalo_Linha`, sem sufixo "_TK", definidas em `Auxiliar.bas`).
- **Exceções:** nenhuma — os dois ramos são mutuamente exclusivos por construção (`If/Else` sobre o mesmo flag).
- **Local de implementação:** `fx_IFRS16.bas:122-128`; `Extracao_Base_1009.bas:395-402`.
- **Evidência:** `Extracao_Base_1009.bas:395-402` (leitura direta do trecho completo).
- **Confiança:** Confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se as duas famílias de funções (`_TK` vs. sem sufixo, ambas em módulos fora do escopo desta entrega — `TK_Functions.bas`/`Auxiliar.bas`) produzem resultado equivalente ou se divergem em algum detalhe — não verificado nesta análise (fora do escopo dos 9 módulos solicitados).

### RN-082 — Cascata "Ref Cruzada", prioridade 1: Regra 2 - CC e Diretoria
- **Descrição:** Na planilha de Ajustes, a regra de redirecionamento de maior prioridade usa Centro de Custo + Diretoria (via `Sup_Linhas`) para achar uma correspondência em `Ref_Cruzada_2`.
- **Condição:** `RC17 <> "-"` (Centro Custo preenchido) E `RC23 <> "-"` E existe correspondência em `Ref_Cruzada_2!C3` para a chave `CC & Diretoria(N1,N2,N3 via Sup_Linhas)`.
- **Ação:** classifica `Ref Cruzada = "Regra 2 - CC e Diretoria"`.
- **Exceções:** nenhuma além da condição.
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:163-167` (fórmula do campo "Ref Cruzada").
- **Evidência:** `Limpeza_Base_Ajustes.bas:163-165`.
- **Confiança:** Confirmado (fórmula lida integralmente).
- **Questão pendente:** nenhuma.

### RN-083 — Cascata "Ref Cruzada", prioridade 2: Regra 1 - Grupo BD
- **Descrição:** Se a Regra 2 não se aplica, verifica correspondência por "Grupo BD" (coluna 3, chave `AA_Chave`) em `Ref_Cruzada_1`.
- **Condição:** `RC17 <> "-"` E `COUNTIFS(Ref_Cruzada_1!C2, RC3, Ref_Cruzada_1!C2, "<>") > 0`.
- **Ação:** classifica `Ref Cruzada = "Regra 1 - Grupo BD"`.
- **Exceções:** só avaliada se RN-082 não se aplicou (ordem `IF` aninhado).
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:165-166`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:165-166`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-084 — Cascata "Ref Cruzada", prioridade 3: Regra 1 - CC
- **Descrição:** Se as regras 2 e 1 (Grupo BD) não se aplicam, verifica correspondência apenas pelo Centro de Custo em `Ref_Cruzada_1`.
- **Condição:** `RC17 <> "-"` E `COUNTIFS(Ref_Cruzada_1!C3, RC17, Ref_Cruzada_1!C3, "<>") > 0`.
- **Ação:** classifica `Ref Cruzada = "Regra 1 - CC"`.
- **Exceções:** só avaliada se RN-082 e RN-083 não se aplicaram.
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:166-167`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:166-167`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-085 — Cascata "Ref Cruzada", fallback: N.A. / "-"
- **Descrição:** Se Centro de Custo = "-", a linha nem entra na cascata (resultado "-"); se Centro de Custo existe mas nenhuma das 3 regras acima bate, resultado "N.A.".
- **Condição:** `RC17 = "-"` → resultado `"-"`; senão, nenhuma das RN-082/083/084 → `"N.A."`.
- **Ação:** define o valor final do campo "Ref Cruzada", que por sua vez é usado como critério de seleção de fonte em outros campos "DESTINO" (Diretoria N1-N3, ver RN-086 nota).
- **Exceções:** nenhuma.
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:163,167`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:163,167`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-086 — Regra "Labour Cost" força CLASSE DESTINO = "W/o Fiber"
- **Descrição:** Independentemente do lookup padrão em `Sup_Linhas`, se o campo "A3_Destino" da linha for "Labour Cost", a Classe de Destino é forçada para "W/o Fiber".
- **Condição:** `RC[A3_Destino] = "Labour Cost"` (e Centro de Custo ≠ "-").
- **Ação:** `CLASSE DESTINO = "W/o Fiber"`.
- **Exceções:** se Centro de Custo = "-", resultado é "-" (prioridade sobre esta regra).
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:183-185`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:184`.
- **Confiança:** Confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** o racional contábil de "Labour Cost → W/o Fiber" para a documentação formal.

### RN-087 — Regra "Rateio" → "Others" em SEG_N2_DESTINO e A1_DESTINO
- **Descrição:** Quando o lookup em `Sheet21` (DP_Segmento) retorna literalmente o valor "Rateio" para Segmento ou Abertura_1, o sistema substitui por "Others" — linhas classificadas para alocação proporcional entre segmentos não podem aparecer nomeadas como um segmento específico.
- **Condição:** `VLOOKUP(CC&CDC, Sheet21!C1:C10, ..., 0) = "Rateio"`.
- **Ação:** `SEG_N2_DESTINO` (ou `A1_DESTINO`) `= "Others"`.
- **Exceções:** se `SEGMENTO`/`ABERTURA_1` já estiver preenchido manualmente na linha, este lookup nem é avaliado (prioridade ao valor digitado).
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:194-199` (SEG_N2_DESTINO), `:205-210` (A1_DESTINO, mesma lógica).
- **Evidência:** `Limpeza_Base_Ajustes.bas:197,209`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma quanto ao mecanismo.

### RN-088 — Fallback "w/o IFRS" no cálculo de CLASSE CUSTO
- **Descrição:** No cálculo do campo "CLASSE CUSTO" de Ajustes, se o campo "IFRS_Contabil" da linha estiver vazio, o sistema usa o texto literal "w/o IFRS" como parte da chave de busca em `Sup_Linhas`.
- **Condição:** `RC7 = ""` (coluna IFRS_Contabil vazia).
- **Ação:** monta a chave de busca como `"w/o IFRS" & Abertura_2..8` em vez de `IFRS_Contabil & Abertura_2..8`.
- **Exceções:** se já houver valor manual em "Ref Cruzada" (coluna deslocada `RC[23]`), usa esse valor diretamente, ignorando o lookup.
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:140-144`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:141-143`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-089 — Padronização de nomenclatura de Empresa no De-Para
- **Descrição:** Durante o refresh de Ref Cruzadas, nomes técnicos de empresa vindos do arquivo externo são renomeados para os nomes de negócio usados no restante do sistema.
- **Condição:** valor da coluna Empresa em `Sup_Linhas` igual a um dos códigos técnicos.
- **Ação:** `TPAR→TPAR_HO`; `Fiber_OG→Fiber Ongoing`; `INT_OG→Intelig Ongoing`; `INT_RR→Intelig Rural`; `Fiber_RR→Fiber Rural`; `CRC→Bloqueado (CRC)`.
- **Exceções:** nenhuma condicional — substituição direta (`Replace ... LookAt:=xlWhole`).
- **Local de implementação:** `Refresh_De_X_Para.bas:139-150`.
- **Evidência:** `Refresh_De_X_Para.bas:143-148`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-090 — Distinção Fiber WTTx vs. Fiber Live
- **Descrição:** Para linhas cuja Empresa original é "FIBER", o sistema decide entre dois nomes de negócio distintos conforme a Classe do registro.
- **Condição:** Empresa = "FIBER".
- **Ação:** se `Classe = "NO UBB"` → Empresa = "Fiber WTTx"; senão → Empresa = "Fiber Live".
- **Exceções:** só aplicável quando Empresa original = "FIBER" (case-insensitive).
- **Local de implementação:** `Refresh_De_X_Para.bas:128-137`.
- **Evidência:** `Refresh_De_X_Para.bas:129-136`.
- **Confiança:** Confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** o significado de negócio de "NO UBB" como critério de segmentação de produto.

### RN-091 — Limite de 747 itens em listas de validação de Ajustes
- **Descrição:** Ao reconstruir as listas de validação dinâmicas de Ajustes, apenas os primeiros 747 valores distintos (ordenados) são incluídos; o excedente é descartado silenciosamente.
- **Condição:** número de valores distintos de uma coluna em `Sup_Linhas` maior que 747.
- **Ação:** corta a lista em 747 itens (`If Qtd_Filtros > 747 Then Qtd_Filtros = 747`).
- **Exceções:** nenhuma — corte incondicional.
- **Local de implementação:** `Lista_Validacao_Ajustes.bas:122-129`.
- **Evidência:** `Lista_Validacao_Ajustes.bas:123`.
- **Confiança:** Confirmado (limite técnico do Excel para listas de validação por valores).
- **Questão pendente:** confirmar com o negócio se alguma coluna já ultrapassou esse limite na prática (risco de opções ausentes sem aviso).

### RN-092 — "Total" sempre por último nas listas de validação
- **Descrição:** Em todas as listas de validação ordenadas do sistema (Ajustes, Main Results — se ativo), o valor "Total" é forçado para a última posição, independentemente da ordem alfabética.
- **Condição:** presença do valor literal "Total" na lista.
- **Ação:** bubble sort customizado que trata "Total" como sempre "maior" que qualquer outro valor.
- **Exceções:** nenhuma.
- **Local de implementação:** `Lista_Validacao_Ajustes.bas:110-121`; `BackupCodigo_MainResults.bas:190-200`; `Auxiliar.bas:1096` (versão pública, usada por `Form_Exportacao.frm:563`).
- **Evidência:** `Lista_Validacao_Ajustes.bas:113`; `BackupCodigo_MainResults.bas:193`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma quanto à regra; nota de engenharia: lógica duplicada em 3 lugares (ver procedimento nº 20, item 6).

### RN-093 — Validação hierárquica automática condicionada a flag (status: possivelmente inerte)
- **Descrição:** A restrição em cascata dos dropdowns de abertura em "Main Results" só deveria rodar se a célula `L32` ("Ligar validação de linhas automáticas") contiver "SIM".
- **Condição:** `UCase(Me.Range("L32")) = "SIM"`.
- **Ação:** executa a validação em cascata (procedimento nº 18/19).
- **Exceções:** nenhuma condicional adicional — mas ver item de confiança abaixo.
- **Local de implementação:** `BackupCodigo_MainResults.bas:17-18`.
- **Evidência:** `BackupCodigo_MainResults.bas:2,17-18`; ausência de qualquer chamada externa a `Worksheet_Change` (grep completo do diretório) e ausência de `Worksheet_Change` equivalente em qualquer arquivo `.cls` do dump.
- **Confiança:** **[VALIDAR COM O NEGÓCIO]** — a regra em si está corretamente implementada no código, mas o procedimento que a executa está estruturalmente fora do ponto em que o Excel dispararia o evento automaticamente (módulo `.bas` comum, não code-behind de planilha) — ver procedimento nº 18, item 6, para a evidência formal completa.
- **Questão pendente:** confirmar com o negócio se esta trava está de fato ativa hoje no arquivo `.xlsb` em produção.

### RN-094 — Calendário de fechamento hardcoded na lista de KPI/Versão
- **Descrição:** Além das combinações reais vindas do SQL, o sistema sempre oferece um conjunto fixo de versões de acompanhamento de fechamento.
- **Condição:** execução de `Atualizar_Lista_KPI_Versao`.
- **Ação:** adiciona via `UNION SELECT` as versões `ACT > Preview 1` a `Preview 6`, `ACT > Pré-Closing`, `ACT > AJ_Pré-Closing`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Refresh_Sup_Linhas.bas:404-411`.
- **Evidência:** `Refresh_Sup_Linhas.bas:404-411`.
- **Confiança:** Confirmado.
- **Questão pendente:** **[VALIDAR COM O NEGÓCIO]** se o calendário de fechamento (nº de previews) já mudou historicamente e exigiu alteração deste código.

### RN-095 — Cálculo de FY como soma de Jan a Dez
- **Descrição:** O total anual (FY) de cada linha de Ajustes é sempre a soma direta dos 12 meses, sem ajuste adicional.
- **Condição:** aplicável a toda linha de dados de `Sheet13`.
- **Ação:** `FY = SUM(Jan:Dez)`.
- **Exceções:** nenhuma.
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:156-158`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:157`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

### RN-096 — CENTRO CUSTO calculado por Diretoria N1-N3 + Classe
- **Descrição:** O Centro de Custo de uma linha de Ajustes (quando não informado manualmente via Ref Cruzada) é derivado por lookup em `Sup_Linhas` usando a combinação Diretoria N1 + N2 + N3 + Classe como chave.
- **Condição:** campo "Ref Cruzada" (offset `RC[16]`) vazio E existe pelo menos uma das 5 colunas-chave preenchida.
- **Ação:** `INDEX/MATCH` em `Sup_Linhas!C22` pela chave `Diretoria N1&N2&N3&Classe&...` (5 componentes, `RC[-5]` a `RC[-1]`).
- **Exceções:** se todas as 5 colunas-chave estiverem vazias, resultado "-".
- **Local de implementação:** `Limpeza_Base_Ajustes.bas:148-152`.
- **Evidência:** `Limpeza_Base_Ajustes.bas:149-151`.
- **Confiança:** Confirmado.
- **Questão pendente:** nenhuma.

---


## 12.4 Cluster UI / Forms (RN-101 a RN-115)
# (B) REGRAS DE NEGÓCIO / UX

| ID | Nome | Descrição | Condição | Ação | Exceções | Local de implementação | Evidência | Confiança | Questão pendente |
|---|---|---|---|---|---|---|---|---|---|
| RN-101 | Bloqueio de digitação em campos de arquivo (Importação) | Os campos de arquivo/diretório do Form_Importacao são tratados como somente-leitura via aviso, forçando o uso do diálogo nativo. | Usuário pressiona uma tecla com foco em `TB_Arq` ou `TB_Diret`. | Exibe MsgBox crítico orientando a usar o botão de pesquisa; tenta setar `Cancel = True`. | — | Form_Importacao.frm:742-757 | Linhas 744-748, 753-757 | Confirmado (existência do código) / **Dúvida técnica real** sobre eficácia (ver campo 15 dos procedimentos 9-10) | O `Cancel = True` não é parâmetro válido de `KeyPress` de TextBox — bloqueio pode não funcionar de fato. [VALIDAR COM O NEGÓCIO / teste funcional no arquivo original] |
| RN-102 | Confirmação obrigatória antes de importar | Nenhuma importação de layout é executada sem confirmação explícita do usuário. | Usuário clica "OK" no Form_Importacao com ao menos 1 layout selecionado. | MsgBox Sim/Não listando os layouts escolhidos; só prossegue se "Sim". | Se lista vazia, bloqueia antes mesmo de perguntar (mensagem de erro distinta). | Form_Importacao.frm:136-141 | Linhas 136-141 | Confirmado | — |
| RN-103 | Renomeação automática de abas duplicadas na importação | Ao importar uma aba cujo nome já existe no destino, o sistema renomeia automaticamith incrementando um sufixo numérico, em vez de bloquear ou perguntar. | Nome da aba de origem já existe entre as abas do workbook destino (exceto para "AJUSTES"/"BASE", que seguem fluxo próprio). | Incrementa sufixo `" NN"` até achar nome livre. | Não se aplica a "AJUSTES" e ao nome de `Sheet3`. | Form_Importacao.frm:159-171 | Linhas 160-170 | Confirmado | Nenhuma validação de que o novo nome gerado não ultrapasse o limite de 31 caracteres do Excel para nomes de aba — [VALIDAR COM O NEGÓCIO] se já ocorreram nomes muito longos na prática. |
| RN-104 | Tratamento especial para abas "AJUSTES" e "BASE" na importação | O fluxo de importação genérico (cópia de aba completa) não se aplica a duas abas com nome reservado; cada uma tem lógica de mapeamento de campo a campo própria. | Nome da aba de origem = "AJUSTES" (case-insensitive) ou igual ao nome de `Sheet3`. | Direciona para os branches específicos (linhas 477 e 563 de Form_Importacao.frm), em vez de criar uma aba nova. | — | Form_Importacao.frm:477-647 | Linhas 477, 563 | Confirmado | — |
| RN-105 | Exigência de escolha do tipo de exportação | Usuário deve escolher entre "Arquivo único" e "Arquivos separados" antes de exportar. | Clique em "OK" no Form_Exportacao. | Se nenhum dos dois OptionButtons estiver marcado, bloqueia com MsgBox crítico e `Exit Sub`. | — | Form_Exportacao.frm:57-61 | Linhas 57-61 | Confirmado | — |
| RN-106 | Exigência de item selecionado para exportação | Não é possível exportar sem ao menos um Front/versão escolhido. | Clique em "OK" com `LB_Versoes_Select` vazia. | MsgBox crítico, `Exit Sub`. | — | Form_Exportacao.frm:70-74 | Linhas 70-74 | Confirmado | — |
| RN-107 | Sobrescrita silenciosa de arquivo exportado | Se já existir um arquivo com exatamente o mesmo nome gerado (timestamp por minuto) na pasta de destino, ele é apagado e substituído sem aviso. | Nome de arquivo gerado (`... - YYYYMMDD_HHMM h.xlsb`) já existe na pasta escolhida. | `Kill` do arquivo antigo, seguido de `SaveAs` do novo. | — | Form_Exportacao.frm:147, 175, 401 | Linhas 147, 175, 401 | Confirmado — **classificado também como Risco Crítico, ver Seção D** | Deveria haver confirmação ou sufixo com granularidade de segundos/GUID? [VALIDAR COM O NEGÓCIO] |
| RN-108 | Exclusão automática de linhas EBITDA e "IFRS Itália" na exportação de Base | Ao exportar a Base, linhas cujo campo ABERTURA_2 = "EBITDA" ou VISAO = "IFRS ITÁLIA" são sempre removidas do arquivo exportado, independentemente da seleção do usuário. | Modo "Exportar Base" (`Page_Frame.Index = 1`). | Chama `Excluir_Item_Especifico` duas vezes com critérios fixos. | — | Form_Exportacao.frm:240-249 | Linhas 240-249 | Confirmado | Motivo de negócio para excluir sempre esses dois recortes não está documentado no código — [VALIDAR COM O NEGÓCIO]. |
| RN-109 | Inconsistência de interação clique-simples vs. duplo-clique entre forms semelhantes | A ação de "remover item da lista de seleção" é acionada por duplo-clique no Form_Importacao mas por clique único no Form_Exportacao, apesar dos dois forms seguirem o mesmo padrão visual de duas ListBox. | Usuário interage com a lista de itens selecionados em qualquer um dos dois forms. | Form_Importacao: exige duplo-clique (`LB_Layouts_Select_DblClick`). Form_Exportacao: um clique já remove (`LB_Versoes_Select_Click`). | — | Form_Importacao.frm:736-739; Form_Exportacao.frm:479-482 | Comparação direta entre os dois eventos | Confirmado | Padronizar qual comportamento é o "correto" para a nova UI — [VALIDAR COM O NEGÓCIO/UX]. |
| RN-110 | Form_Tratamento_Opcoes calcula mas não exibe | A lista de extrações habilitadas ("Sim") é calculada em tempo de abertura do form mas nunca é apresentada visualmente ao usuário — só é enviada ao `Debug.Print` (Immediate Window do editor VBA). | Form_Tratamento_Opcoes é aberto/inicializado. | Nenhuma ação visível — dado calculado é descartado do ponto de vista do usuário final. | — | Form_Tratamento_Opcoes.frm:53, 89, 112 | Linhas 53, 89, 112 (`Debug.Print dados`) | Confirmado | Este comportamento é intencional (feature incompleta) ou é uma regressão? [VALIDAR COM O NEGÓCIO]. |
| RN-111 | Menu de comandos do ComboBox1 (Sheet8) por correspondência textual | O gatilho de comandos rápidos do painel principal depende do texto do item selecionado conter uma substring específica, não um ID/valor fixo. | Usuário seleciona um item de `ComboBox1` em `Sheet8`. | Se contém "Defined Names" → `CLEAR_Defined_Names`; se contém "Combinações" → `UPDATE_Combinacoes_Empresas`; qualquer outro texto → nenhuma ação, sem aviso. | — | Sheet8.cls:10-18 | Linhas 12-16 | Confirmado | A população atual do combo (texto exato dos itens) está em código comentado (Worksheet_Activate desativado) — origem atual dos itens do combo é [NÃO IDENTIFICADO]. |
| RN-112 | Atalho de abertura de arquivo/pasta por duplo-clique em célula (Sheet8) | Duplo-clique em uma célula logo abaixo do rótulo "Diretório:" ou "Arquivo:" abre diretamente o item referenciado, sem necessidade de copiar o caminho. | Duplo-clique em célula de `Sheet8` cuja célula à esquerda contenha exatamente "Diretório:" ou "Arquivo:". | Abre o arquivo (Excel, modo leitura) ou a pasta (Windows Explorer), conforme o caso; cancela a edição padrão da célula. | Se o caminho estiver desatualizado, o erro nativo do Windows/Excel aparece, sem tratamento customizado. | Sheet8.cls:35-62 | Linhas 39-58 | Confirmado | — |
| RN-113 | Seletor de modo/visão via 3 OptionButtons acoplado a célula fixa de outra aba | A escolha de uma entre 3 opções mutuamente exclusivas em `Sheet3` é persistida diretamente em `Sheet11!AV6`, sem camada de abstração (nome definido/constante). | Clique em qualquer um dos 3 `OptionButton`s de `Sheet3`. | Grava o `Caption` do botão clicado em `Sheet11!AV6`; aplica negrito ao botão ativo. | — | Sheet3.cls:9-28 | Linhas 10, 17, 24 | Confirmado (mecânica) / conteúdo exato dos 3 `Caption`s é [NÃO ACESSÍVEL] | Qual é o significado de negócio de cada uma das 3 opções? [VALIDAR COM O NEGÓCIO]. |
| RN-114 | Referência a procedimento inexistente (`Registrar_Sheet`) | O evento de duplo-clique de `ListBox1` no Form_Exportacao chama um procedimento (`Registrar_Sheet`) que não existe em nenhum módulo do projeto. | Duplo-clique em `ListBox1` do Form_Exportacao com valor não vazio. | Tentativa de chamada a procedimento inexistente — comportamento real depende de o projeto VBA compilar ou não com essa referência pendente. | — | Form_Exportacao.frm:467-471 | Linha 469; ausência confirmada via grep em todo `vba_dump_tmp/` | Confirmado como referência quebrada no código-fonte disponível — **também listado como Risco Crítico #1, Seção D** | `ListBox1` ainda existe/é visível no form em produção? O projeto `.xlsb` original compila sem erro? [NÃO ACESSÍVEL sem o binário — VALIDAR COM O NEGÓCIO/TI]. |
| RN-115 | Mensagem final genérica sem resumo quantitativo | Tanto a importação quanto a exportação encerram com uma única MsgBox de sucesso, sem informar quantos itens/arquivos foram processados nem onde foram salvos. | Conclusão bem-sucedida de `B_Ok_Click` em qualquer um dos dois forms. | Exibe "Processo concluído com sucesso!" (Exportação) / popup de tempo de processamento (Importação, via `PopUp_Tempo_Processamento`). | — | Form_Exportacao.frm:423; Form_Importacao.frm:699 (via `PopUp_Tempo_Processamento`, Auxiliar.bas:138-147) | Linhas citadas | Confirmado | — |

---

