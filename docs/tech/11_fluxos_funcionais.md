# 11. Fluxos Funcionais e Operacionais

## Fluxo F-01 — Atualização das tabelas mestre (`Refresh_Base_Aux`)

| Campo | Detalhe |
|---|---|
| Objetivo | Reconstruir do zero as tabelas de mapeamento (DP_Segmento, Ref_Cruzada_1/2, Sup_Linhas, DropComb, Valid_Lin) a partir do SQL Server e de um arquivo externo, para que toda extração subsequente use a classificação gerencial mais atual |
| Gatilho | **[NÃO IDENTIFICADO]** — nenhum procedimento do código lido chama `Refresh_Base_Aux`; é quase certo que exista um botão de planilha ("Atualizar bases auxiliares", confirmado como shape na aba Extração) apontando para esta macro, mas a ligação shape→macro não está no texto do código, só é visível no VBE/objeto do shape |
| Entrada | Tabelas SQL Server (`BPAM`), arquivo externo `Bases_DE_PARA.xlsx` |
| Validações | Nenhuma validação de sucesso da consulta SQL antes de prosseguir para a etapa seguinte |
| Etapas | 1) `Refresh_Base_Segmento` → 2) `Refresh_Base_De_Para_Ref_Cruzadas` → 3) `Refresh_Base_Suporte_Linhas` → 4) `Refresh_Drop_Comb_Hubble` → 5) `Extrair_Valid_Lin` (ordem fixa confirmada pela sequência de chamadas dentro de `Refresh_Base_Aux`, em `Auxiliar.bas`) |
| Abas envolvidas | DP_Segmento, Ref_Cruzada_1/2, CC BD, Sup_Linhas, DropComb, Valid_Lin |
| Macros envolvidas | `Refresh_DP_Segmento.bas`, `Refresh_De_X_Para.bas`, `Refresh_Sup_Linhas.bas`, `Refresh_Drop_Comb.bas` |
| Saída | Tabelas mestre atualizadas |
| Mensagens e falhas possíveis | Falha de conexão SQL ou arquivo de rede indisponível interrompe a etapa correspondente sem tratamento de erro robusto (ver seção 16) — as etapas seguintes do pipeline rodam mesmo assim, potencialmente com dados desatualizados/parciais |
| Forma de recuperação | Reexecutar a macro (idempotente por natureza — reconstrói do zero) |
| Pontos de validação manual | Nenhum ponto de confirmação do usuário neste fluxo — roda direto |

## Fluxo F-02 — Extração de uma fonte de dados individual

| Campo | Detalhe |
|---|---|
| Objetivo | Trazer dados de uma fonte específica (Hubble, 1009, RGM, etc.) para a Base |
| Gatilho | Botão individual da aba Extração ("EXTRAIR/IMPORTAR" por fonte) |
| Entrada | Depende da fonte — SQL (Hubble) ou arquivo externo apontado previamente via "INFORMAÇÕES"/diálogo de arquivo |
| Validações | Validação estrutural (comparação com tabela de chaves) nas fontes RGM/MOCKUP/Fixed Revenues; nenhuma validação de estrutura nas demais além de localizar cabeçalho por texto |
| Etapas | Limpar histórico da fonte → ler/abrir dados → copiar para a Base → aplicar enriquecimento padrão (`Form_*`) → atualizar lista de KPI/Versão → (condicionalmente) aplicar IFRS16 → mostrar tempo de processamento |
| Abas envolvidas | Extração (config), Base (destino), tabelas de chave específicas da fonte |
| Macros envolvidas | Módulo `Extracao_*` correspondente |
| Saída | Novas linhas na Base, com `Fonte` = nome da base de origem |
| Mensagens e falhas possíveis | Genéricas ("Ocorreu um erro ao abrir este arquivo!"); em RGM/MOCKUP, falha estrutural gera `MsgBox` crítico + possível `End` abrupto (ver seção 18, risco crítico) |
| Forma de recuperação | Rodar "Limpar base" da fonte específica e reextrair |
| Pontos de validação manual | `MsgBox` de confirmação em alguns módulos (ex.: excluir linhas de "Other Income"/"Net Service Revenues") — bloqueante, exige clique do usuário no meio do processo |

## Fluxo F-03 — Extração de todas as bases (`Extrair_Todas_as_Bases`)

| Campo | Detalhe |
|---|---|
| Objetivo | Rodar o Fluxo F-02 para as 7 fontes extraíveis automaticamente em sequência |
| Gatilho | Botão "Extrair todas as bases" |
| Entrada | Idem F-02, para todas as fontes configuradas |
| Etapas | `Verifica_Versao` → `Desligar_Tudo` → limpar histórico → chamar cada extração em sequência → `Calcular_Comb_Meses` → `Ativar_Tudo` → popup de tempo decorrido |
| Saída | Base totalmente reconstruída para o período |
| Mensagens e falhas possíveis | Se `Verifica_Versao` falhar (versão local diferente da esperada no SQL), o processo é abortado via `End` logo no início |
| Forma de recuperação | Reexecutar |
| Pontos de validação manual | Cada `MsgBox` de confirmação das extrações individuais ainda aparece dentro deste fluxo — **não é possível rodar isso de forma desatendida hoje** |

## Fluxo F-04 — Ajustes manuais

| Campo | Detalhe |
|---|---|
| Objetivo | Permitir correções pontuais inseridas manualmente pelo analista, sem depender de nenhuma fonte externa |
| Gatilho | Digitação direta na planilha Ajustes + botão "EXTRAIR/IMPORTAR" da fonte Ajustes |
| Entrada | Digitação do usuário |
| Validações | Listas de validação (dropdowns) reconstruídas por `Lista_Validacao_Ajustes.bas`; fórmulas de destino reconstruídas por `Limpeza_Base_Ajustes.bas` |
| Etapas | Usuário digita → fórmulas de destino calculam automaticamente (Classe/CC/Diretoria/Segmento destino) → `Extrair_Base_Ajustes` copia para a Base |
| Saída | Linhas na Base, `Fonte` = "Base_Ajustes" |
| Mensagens e falhas possíveis | Se não houver ajustes preenchidos, sai com aviso informativo (ainda exige clique OK) |
| Forma de recuperação | Corrigir a digitação e reextrair |
| Pontos de validação manual | Reconstrução das fórmulas de destino (`Limpar_Ajustes`) deve ser rodada sempre que as tabelas mestre mudarem, para as fórmulas ficarem coerentes — **não há automação garantindo que isso aconteça na ordem certa**; é responsabilidade do usuário lembrar de rodar |

## Fluxo F-05 — Tratamento IFRS16

| Campo | Detalhe |
|---|---|
| Objetivo | Reclassificar custos de leasing conforme a norma IFRS16 |
| Gatilho | Automático, dentro de `Extrair_Base_1009`, condicionado à flag `Sheet8!I25 = "Sim"` |
| Entrada | Linhas da Base com `Fonte` iniciando em "Base_1009"; De-Para mantido manualmente em Sup_Linhas (colunas 71/73/15/22) |
| Validações | Linhas sem mapeamento em Sup_Linhas são puladas (contabilizadas, não bloqueiam o processo) |
| Etapas | Filtrar linhas-alvo → marcar original como "DELETAR" → copiar para staging (Aux_IFRS16) com CC/Fonte ajustados e sinal invertido → reintegrar na Base → remover originais marcadas |
| Saída | Linhas com `Fonte` = "..._IFRS16 Tratado" |
| Mensagens e falhas possíveis | Aviso de quantas linhas ficaram sem tratamento por falta de mapeamento; se a flag estiver desligada, aviso "Linhas IFRS16 não tratadas!" |
| Forma de recuperação | Reexecutar a extração da Base 1009 |
| Pontos de validação manual | A manutenção do De-Para (colunas 71/73/15/22 de Sup_Linhas) é 100% manual e fora de qualquer automação — **[VALIDAR COM O NEGÓCIO]** quem faz essa manutenção hoje |

## Fluxo F-06 — Geração de Pré-Closing

| Campo | Detalhe |
|---|---|
| Objetivo | Simular um fechamento antecipado, recombinando dados já existentes na Base conforme cenários configurados |
| Gatilho | Botão "Gerar base pré-closing!" |
| Entrada | Matriz de cenários na aba "Preview"; dados já presentes na Base |
| Etapas | Para cada coluna de cenário: aplicar `AutoFilter` na Base com os critérios da matriz → copiar linhas visíveis para workbook temporário → zerar meses fora do intervalo do cenário → colar de volta na Base com nova Versão/Fonte |
| Saída | Novas linhas na Base, `Fonte` = "QD AUTOMATIC - SELEÇÃO MANUAL"/"QD AUTOMATIC - INPUT AJUSTE" |
| Mensagens e falhas possíveis | `MsgBox` de confirmação para limpar base histórica da versão antes de gerar |
| Forma de recuperação | Reexecutar |
| Pontos de validação manual | Confirmação de limpeza antes de gerar |

## Fluxo F-07 — Importação/Exportação de Fronts ou Base

| Campo | Detalhe |
|---|---|
| Objetivo | Mover abas de relatório ("Fronts") ou a Base filtrada entre arquivos Quick Data diferentes |
| Gatilho | Botões "Importar/Exportar BASE e/ou FRONT(S)" |
| Entrada | Arquivo externo (importação) ou seleção de itens + pasta de destino (exportação) |
| Validações | Exige ao menos um item selecionado; confirmação Sim/Não antes de executar |
| Etapas | Ver detalhamento dos UserForms na seção 10 |
| Saída | Arquivo(s) `.xlsb` novo(s) (exportação) ou abas copiadas para o workbook atual (importação) |
| Mensagens e falhas possíveis | Mensagens genéricas de erro ao abrir arquivo; sobrescrita silenciosa de arquivo existente na exportação |
| Forma de recuperação | Repetir a operação |
| Pontos de validação manual | Confirmação da lista de itens antes de processar |

## Fluxo F-08 — Saneamento de Defined Names

| Campo | Detalhe |
|---|---|
| Objetivo | Remover nomes definidos quebrados/obsoletos que se acumularam ao longo dos anos |
| Gatilho | Botão "Limpar Defined Names com erro" / opção no combo de comandos rápidos de Sheet8 |
| Entrada | Lista de 166 nomes definidos, classificados pela própria planilha `ListDefinedNames` |
| Etapas | `CLEAR_Defined_Names` classifica por tipo de link → `RUN_Apagar_defined_names_definitivamente` remove os marcados |
| Saída | Nomes definidos removidos do workbook |
| Forma de recuperação | Não há — é uma exclusão definitiva; recomenda-se backup do arquivo antes de rodar (ver seção 19) |
