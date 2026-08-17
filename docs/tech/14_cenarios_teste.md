# 21. Cenários de Teste

> Cenários derivados do comportamento efetivamente encontrado no código (seções 10-18), não de suposições. Cada um cita a evidência que o justifica.

## CT-01 — Fluxo normal: atualização de tabelas mestre

- **Objetivo:** confirmar que `Refresh_Base_Aux` reconstrói as 5 tabelas mestre sem erro, na ordem esperada.
- **Pré-condição:** conexão SQL Server e arquivo `Bases_DE_PARA.xlsx` acessíveis.
- **Entrada:** execução normal do botão "Atualizar bases auxiliares".
- **Passos:** rodar a macro; conferir, nesta ordem, que DP_Segmento, Ref_Cruzada_1/2, Sup_Linhas, DropComb e Valid_Lin foram todas atualizadas.
- **Resultado esperado:** as 5 tabelas com dados novos, sem erro registrado em `tk_Lista_de_erros`.
- **Componentes afetados:** os 4 módulos `Refresh_*.bas`.
- **Evidência que justifica:** sequência confirmada de chamadas em `Refresh_Base_Aux` (`Auxiliar.bas`).

## CT-02 — Arquivo de entrada fora do padrão (estrutura divergente) — RGM/MOCKUP RGM

- **Objetivo:** verificar o comportamento quando o arquivo fonte não bate com a tabela de chaves esperada.
- **Pré-condição:** arquivo RGM ou MOCKUP RGM com uma linha renomeada/reordenada em relação ao esperado pela chave.
- **Entrada:** arquivo alterado propositalmente.
- **Passos:** rodar a extração da fonte.
- **Resultado esperado (comportamento atual, não desejável):** `MsgBox` crítico seguido de `End` — o VBA para abruptamente, possivelmente deixando `ScreenUpdating`/`Calculation` desligados.
- **Componentes afetados:** Excel inteiro (não só a extração) até fechar/reabrir o arquivo.
- **Evidência que justifica:** validação estrutural com `End` confirmada em `Extracao_Base_MOCKUP_RGM.bas`/`Extracao_Base_RGM.bas`, risco crítico (seção 18).

## CT-03 — Colunas ausentes — Base Quick Data (extração sistema-a-sistema)

- **Objetivo:** verificar o comportamento quando uma coluna esperada não existe no arquivo fonte.
- **Entrada:** arquivo "Quick Data" sem a coluna "Organic", por exemplo.
- **Resultado esperado (atual):** `fn.Match` retorna `#N/A` não tratado — a macro provavelmente quebra sem mensagem clara.
- **Evidência:** `Extracao_Base_Quick_Data.bas`, ausência de tratamento de erro identificada.

## CT-04 — Tipos de dados inválidos — valores não numéricos em colunas de mês

- **Objetivo:** verificar o que acontece quando uma célula de mês contém texto em vez de número.
- **Entrada:** arquivo fonte (qualquer um dos 7 externos) com uma célula de mês contendo texto.
- **Resultado esperado (atual):** comportamento não determinado pela análise de código — **[NÃO IDENTIFICADO]**, precisa de teste real; `Transformar_Texto_Mes_Em_Valor` trata "-" especificamente, mas não está confirmado que trata qualquer texto arbitrário.
- **Evidência:** `Aux_Formulas_Base.bas`, `Transformar_Texto_Mes_Em_Valor`.

## CT-05 — Volume elevado (linhas além do limite hardcoded)

- **Objetivo:** verificar o comportamento quando a Base ultrapassa 40.837 linhas.
- **Entrada:** Base com mais de 40.837 linhas, seguida de `UPDATE_Combinacoes_Empresas_parte2`.
- **Resultado esperado (atual):** linhas além do limite **não são preenchidas** pela sincronização de combinações — bug silencioso de dados incompletos.
- **Evidência:** `TK_Functions.bas`, destino fixo `"EC6:GD40837"` em `Selection.AutoFill`, risco crítico (seção 18, cluster core).

## CT-06 — Muitos "Fronts"/itens selecionados na Importação/Exportação

- **Objetivo:** verificar o comportamento com uma seleção grande de abas para importar/exportar.
- **Entrada:** selecionar todas as abas disponíveis de um arquivo com muitas planilhas.
- **Resultado esperado (atual):** processamento sequencial, sem paralelismo, com feedback só na barra de status — tempo total não estimado ao usuário; sem opção de cancelar no meio.
- **Evidência:** `Form_Importacao.frm`/`Form_Exportacao.frm`, ausência de barra de progresso real (seção 18, cluster UI).

## CT-07 — Caminho de rede inexistente/inacessível

- **Objetivo:** verificar o comportamento quando o caminho de rede configurado (Hubble ou De-Para) não existe mais/está inacessível.
- **Entrada:** apontar `Sheet24` para um caminho inválido.
- **Resultado esperado (atual):** falha na abertura do arquivo sem tratamento robusto em vários módulos de refresh (`On Error GoTo` comentado) — comportamento exato depende do módulo, mas a expectativa é erro não tratado interrompendo a macro.
- **Evidência:** seção 16.1 (tratamento de erro comentado, recorrente).

## CT-08 — Arquivo bloqueado (aberto por outro usuário)

- **Objetivo:** verificar o comportamento quando o arquivo externo (De-Para, RGM etc.) está aberto por outra pessoa em modo de edição.
- **Entrada:** abrir o arquivo fonte em outra sessão do Excel antes de rodar a extração.
- **Resultado esperado (atual):** **[NÃO IDENTIFICADO]** — os módulos abrem arquivos com `ReadOnly:=True` na maioria dos casos confirmados, o que normalmente evita conflito de bloqueio, mas isso não foi confirmado para 100% dos módulos de extração.
- **Evidência:** confirmado `ReadOnly:=True` em ao menos `Refresh_De_X_Para.bas`; não verificado sistematicamente em todos os `Extracao_*.bas`.

## CT-09 — Execução interrompida no meio (usuário fecha o Excel/cai a rede)

- **Objetivo:** verificar o estado da tabela SQL temporária (Hubble) se a execução for interrompida entre a criação e a exclusão dela.
- **Entrada:** interromper a execução (ex.: encerrar o processo do Excel) durante `Extrair_Base_Hubble`.
- **Resultado esperado (atual):** a tabela temporária `TB_AUX_HUBBLE_QUICK_DATA_<usuário>` fica "presa" no SQL Server — a próxima execução do mesmo usuário pode falhar ou duplicar dados, pois o código assume que a tabela não existe ao tentar criá-la.
- **Evidência:** `Extracao_SQL_Hubble.bas`, `Criar_TB_SQL_AUX`, risco documentado na seção 18 (cluster ETL).

## CT-10 — Reprocessamento (rodar a mesma extração duas vezes seguidas)

- **Objetivo:** confirmar que rodar a mesma extração duas vezes não duplica dados.
- **Entrada:** rodar "Extrair Base Hubble" (ou qualquer fonte) duas vezes seguidas sem alterar nada.
- **Resultado esperado:** a segunda execução deve limpar o histórico da fonte antes de reextrair (`Processo_Limpar_Base_*`), resultando na mesma quantidade final de linhas — comportamento idempotente **por design**, mas depende de `Limpar_Base_Historica` localizar corretamente o bloco de linhas antigo pelo campo "Fonte".
- **Evidência:** padrão confirmado em todos os módulos `Extracao_*` (`Processo_Limpar_Base_*` sempre chamado antes da extração).

## CT-11 — IFRS16 com De-Para incompleto

- **Objetivo:** verificar o comportamento quando uma linha da Base 1009 não tem correspondência nas colunas de De-Para de `Sup_Linhas` (71/73).
- **Entrada:** Base 1009 com um Centro de Custo/Classe Custo não cadastrado no De-Para IFRS16.
- **Resultado esperado (atual):** a linha é pulada e contabilizada num contador de "não tratadas", mostrado ao usuário — mas sem indicar **qual** linha especificamente.
- **Evidência:** `fx_IFRS16.bas`, RN-074 (seção 12, cluster refresh).

## CT-12 — Pré-Closing com matriz de cenários incompleta

- **Objetivo:** verificar comportamento quando a aba "Preview" tem uma coluna de cenário mal preenchida.
- **Entrada:** matriz de cenários com um critério vazio inesperado.
- **Resultado esperado (atual):** **[NÃO IDENTIFICADO]** — a lógica de `Filtrar_Item_Base` trata alguns campos especiais com default, mas o comportamento para uma coluna genuinamente mal formatada não foi confirmado; risco de reaproveitamento incorreto da variável `Lin` entre blocos (ver seção 18, cluster ETL).

## CT-13 — Exportação sobrescrevendo arquivo existente

- **Objetivo:** confirmar que a exportação sobrescreve silenciosamente um arquivo já existente com nome coincidente.
- **Entrada:** gerar um arquivo de exportação, depois gerar outro que colida no nome (mesmo timestamp, cenário raro, ou nome custom se aplicável).
- **Resultado esperado (atual):** o arquivo antigo é apagado (`Kill`) sem confirmação.
- **Evidência:** `Form_Exportacao.frm`, RN-107.

## CT-14 — Importação de aba com nome já existente no destino

- **Objetivo:** confirmar o comportamento de renomeação automática.
- **Entrada:** importar uma aba cujo nome já existe no workbook atual.
- **Resultado esperado (atual):** sufixo numérico incremental automático evita colisão.
- **Evidência:** `Form_Importacao.frm`, RN-103.

## CT-15 — Saneamento de Defined Names em arquivo com nomes legítimos parecidos com "lixo"

- **Objetivo:** garantir que a rotina de limpeza não remove nomes definidos que ainda estão em uso.
- **Entrada:** rodar `RUN_Apagar_defined_names_definitivamente` num arquivo de teste (cópia).
- **Resultado esperado:** apenas os nomes classificados como link quebrado/rede não-P&C são removidos; nomes `CB_*` (usados ativamente) permanecem.
- **Evidência:** classificação já existente na planilha `ListDefinedNames`.
- **Observação:** **esta é uma operação destrutiva — só testar em cópia, nunca no arquivo produtivo** (ver seção 19.7).

## CT-16 — Motor de rateio com percentual variável ausente para o mês/versão

- **Objetivo:** verificar o comportamento quando `Sh_Rateio` não tem percentual cadastrado para a combinação mês/KPI/versão/data da linha sendo rateada.
- **Entrada:** linha marcada para rateio variável sem entrada correspondente na tabela de percentuais.
- **Resultado esperado (atual):** **[NÃO IDENTIFICADO]** — não confirmado se há fallback (ex.: rateio zero, erro, ou uso do último percentual disponível); recomenda-se teste dirigido antes de confiar nessa regra numa reescrita.
- **Evidência:** `Form_Segmentos` (`Aux_Formulas_Base.bas`), RN-045.

## CT-17 — Validação de hierarquia do Main Results (status incerto)

- **Objetivo:** confirmar se `BackupCodigo_MainResults.Worksheet_Change` dispara ao editar uma célula de abertura na planilha Main Results.
- **Pré-condição:** arquivo aberto no Excel real (este teste não pode ser feito por análise estática).
- **Entrada:** editar manualmente uma célula de abertura gerencial na aba Main Results, com `L32 = "SIM"`.
- **Resultado esperado a confirmar:** se a lista de valores válidos da próxima coluna é restringida automaticamente (comportamento ativo) ou se nada acontece (código inerte, como a análise de código sugere).
- **Evidência/motivo do teste:** o procedimento está em um módulo `.bas` comum, não no code-behind da planilha — em VBA, isso normalmente significa que o evento **não dispara automaticamente**. Ver seção 18 (risco alto) e seção 23 (dúvida pendente).
