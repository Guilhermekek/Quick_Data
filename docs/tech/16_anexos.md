# 24. Anexos

## 24.1 Lista completa de componentes do projeto VBA

70 componentes confirmados (fonte: stream `PROJECT` do `vbaProject.bin`):
- 1 `ThisWorkbook`
- 33 módulos de classe de planilha: `Sheet1` a `Sheet33`, com gaps em `Sheet32` (nunca existiu no projeto atual — evidência: ausente do stream `PROJECT`) — ver observação de CodeNames órfãos na seção 7.1
- 3 UserForms: `Form_Importacao`, `Form_Exportacao`, `Form_Tratamento_Opcoes`
- 33 módulos padrão (`.bas`): `Auxiliar`, `TK_Functions`, `Aux_Formulas_Base`, `Extracao_SQL_Hubble`, `Extracao_Base_1009`, `Extracao_Base_Consolidad`, `Extracao_Base_MockUp_RGM`, `Extracao_Base_Other_Inco`, `Extracao_Base_Quick_Data`, `Extracao_Base_RGM`, `Extracao_Fixed_Revenues`, `Extracao_Sheet_Ajustes`, `Aux_Leitura_Nome_Arqs`, `Conexoes`, `Refresh_Sup_Linhas`, `Refresh_De_X_Para`, `Refresh_DP_Segmento`, `Refresh_Drop_Comb`, `Limpeza_Base_Ajustes`, `Lista_Validacao_Ajustes`, `BackupCodigo_MainResults`, `fx_IFRS16`, `Front_Processos`, `Gerar_Base_Pre_Closing`, `Module1` a `Module10` (9 módulos residuais de macro gravada/depuração, ver seção 9.3)

Detalhamento por componente: seção 9.

## 24.2 Relação de procedimentos (índice)

**192 procedimentos confirmados** (contagem mecânica por assinatura de declaração `Sub`/`Function`/`Property`). Distribuição por cluster de documentação (seção 10):

| Cluster | Módulos | Procedimentos documentados |
|---|---|---|
| Extração / ETL | 12 módulos (`Extracao_*`, `Conexoes`, `Aux_Leitura_Nome_Arqs`, `Gerar_Base_Pre_Closing`) | 54 |
| Core / Cálculo | 3 módulos (`Auxiliar`, `Aux_Formulas_Base`, `TK_Functions`) | 74 |
| Refresh / Validação / IFRS16 | 9 módulos (`Refresh_*`, `Limpeza_Base_Ajustes`, `Lista_Validacao_Ajustes`, `Front_Processos`, `BackupCodigo_MainResults`, `fx_IFRS16`) | 23 |
| UI / Forms | 3 UserForms + `Sheet8.cls` + `Sheet3.cls` + `ThisWorkbook.cls` + `Module2.bas` + 30 `Sheet*.cls` vazios | ~35 (mais 30 planilhas confirmadas sem código) |
| **Total** | **70 componentes** | **192** |

O detalhamento completo de cada procedimento (15 campos por entrada) está na seção 10, organizado nos mesmos 4 clusters.

## 24.3 Relação de nomes definidos

166 nomes definidos confirmados (fonte: planilha de auditoria interna `ListDefinedNames`), sendo **42 sinalizados pelo próprio arquivo como candidatos a exclusão**. Lista completa:

| ID | Nome | Classificacao | Apagar? |
|---|---|---|---|
| 1 | _AMO_XmlVersion | Nada identificado | Nao |
| 2 | _Example | Link de Rede não P&C | Sim |
| 3 | _FilterDatabase | Link de SharePoint / Link c/ erro de referência | Sim |
| 4 | _Look | Link de Rede não P&C | Sim |
| 5 | _Order1 | Nada identificado | Nao |
| 6 | _Order2 | Nada identificado | Nao |
| 7 | _Series | Link de Rede não P&C | Sim |
| 8 | _Shading | Link de Rede não P&C | Sim |
| 9 | AAA_DOCTOPS | Nada identificado | Nao |
| 10 | AAA_duser | Nada identificado | Nao |
| 11 | AAB_Addin5 | Nada identificado | Nao |
| 12 | ac | Nada identificado | Nao |
| 13 | ALVARA. | Link de Rede não P&C | Sim |
| 14 | Alvarás. | Link de Rede não P&C | Sim |
| 15 | anscount | Nada identificado | Nao |
| 16 | AS2DocOpenMode | Nada identificado | Nao |
| 17 | at | Nada identificado | Nao |
| 18 | blop | Link de Rede não P&C | Sim |
| 19 | blop15 | Link de Rede não P&C | Sim |
| 20 | BLOP21 | Disco Local do Usuário | Sim |
| 21 | blop9 | Disco Local do Usuário | Sim |
| 22 | BLPH10 | Link de Rede não P&C | Sim |
| 23 | BLPH15 | Link de Rede não P&C | Sim |
| 24 | BLPH19 | Link de Rede não P&C | Sim |
| 25 | BLPH2 | Link de Rede não P&C | Sim |
| 26 | BLPH20 | Link de Rede não P&C | Sim |
| 27 | BLPH21 | Link de Rede não P&C | Sim |
| 28 | BLPH22 | Link de Rede não P&C | Sim |
| 29 | BLPH23 | Link de Rede não P&C | Sim |
| 30 | BLPH3 | Link de Rede não P&C | Sim |
| 31 | BLPH4 | Link de Rede não P&C | Sim |
| 32 | BLPH5 | Link de Rede não P&C | Sim |
| 33 | BLPH9 | Link de Rede não P&C | Sim |
| 34 | BLPo20 | Disco Local do Usuário | Sim |
| 35 | bt | Nada identificado | Nao |
| 36 | CB_ABERTURA_2 | Nada identificado | Nao |
| 37 | CB_Anos | Nada identificado | Nao |
| 38 | CB_Campos_FIltro | Nada identificado | Nao |
| 39 | CB_Classe | Nada identificado | Nao |
| 40 | CB_Colunas | Nada identificado | Nao |
| 41 | CB_Divisor | Nada identificado | Nao |
| 42 | CB_Empresa | Nada identificado | Nao |
| 43 | CB_Empresa_DropComb | Nada identificado | Nao |
| 44 | CB_Formulas | Nada identificado | Nao |
| 45 | CB_IFRS_Contabil | Nada identificado | Nao |
| 46 | CB_IFRS_CONTABIL_DropComb | Nada identificado | Nao |
| 47 | CB_KPI | Nada identificado | Nao |
| 48 | CB_KPI_VERSAO | Nada identificado | Nao |
| 49 | CB_Mes_Ano_Text | Nada identificado | Nao |
| 50 | CB_Mes_Text | Nada identificado | Nao |
| 51 | CB_Meses | Nada identificado | Nao |
| 52 | CB_N1_Gerencial | Nada identificado | Nao |
| 53 | CB_N2_Gerencial_Linha | Nada identificado | Nao |
| 54 | CB_N3_Gerencial_Linha | Nada identificado | Nao |
| 55 | CB_Organic | Nada identificado | Nao |
| 56 | CB_PROFORMA | Nada identificado | Nao |
| 57 | CB_PROFORMA_DropComb | Nada identificado | Nao |
| 58 | CB_Versoes_Col | Nada identificado | Nao |
| 59 | CB_Versoes_Lin | Nada identificado | Nao |
| 60 | CB_Visao | Nada identificado | Nao |
| 61 | Col | Nada identificado | Nao |
| 62 | ColMenos_1 | Nada identificado | Nao |
| 63 | ColMenos_2 | Nada identificado | Nao |
| 64 | ColMenos_3 | Nada identificado | Nao |
| 65 | ColMenos_4 | Nada identificado | Nao |
| 66 | DC_Visao | Nada identificado | Nao |
| 67 | dd | Nada identificado | Nao |
| 68 | ddddc | Nada identificado | Nao |
| 69 | ddddt | Nada identificado | Nao |
| 70 | dep | Link de Rede não P&C | Sim |
| 71 | deposito | Link de Rede não P&C | Sim |
| 72 | DEPOSITO1 | Link de Rede não P&C | Sim |
| 73 | depositoss | Link de Rede não P&C | Sim |
| 74 | ev.Calculation | Nada identificado | Nao |
| 75 | ev.Initialized | Nada identificado | Nao |
| 76 | EV___PARKEDCVW__ | Nada identificado | Nao |
| 77 | EV__DECIMALSYMBOL__ | Nada identificado | Nao |
| 78 | EV__EXPOPTIONS__ | Nada identificado | Nao |
| 79 | EV__LASTREFTIME__ | Nada identificado | Nao |
| 80 | EV__LASTREFTIME__2 | Nada identificado | Nao |
| 81 | EV__LOCKEDCVW__BU_DW | Nada identificado | Nao |
| 82 | EV__LOCKEDCVW__BU_IM | Nada identificado | Nao |
| 83 | EV__LOCKEDCVW__BU_IO | Nada identificado | Nao |
| 84 | EV__LOCKEDCVW__BU_MB | Nada identificado | Nao |
| 85 | EV__LOCKEDCVW__BU_OT | Nada identificado | Nao |
| 86 | EV__LOCKEDCVW__BU_TG | Nada identificado | Nao |
| 87 | EV__LOCKEDCVW__FINANCIALS | Nada identificado | Nao |
| 88 | EV__LOCKEDCVW__FLASH | Nada identificado | Nao |
| 89 | EV__LOCKEDCVW__ICMATCHING | Nada identificado | Nao |
| 90 | EV__LOCKEDCVW__OWNERSHIP | Nada identificado | Nao |
| 91 | EV__LOCKEDCVW__RATE | Nada identificado | Nao |
| 92 | EV__LOCKSTATUS__ | Nada identificado | Nao |
| 93 | EV__MAXEXPCOLS__ | Nada identificado | Nao |
| 94 | EV__MAXEXPROWS__ | Nada identificado | Nao |
| 95 | EV__MEMORYCVW__ | Nada identificado | Nao |
| 96 | EV__WBEVMODE__ | Nada identificado | Nao |
| 97 | EV__WBREFOPTIONS__ | Nada identificado | Nao |
| 98 | EV__WBVERSION__ | Nada identificado | Nao |
| 99 | EV__WSINFO__ | Nada identificado | Nao |
| 100 | final | Link de Rede não P&C | Sim |
| 101 | Formula_Delta_Abs | Nada identificado | Nao |
| 102 | Formula_Delta_Perc | Nada identificado | Nao |
| 103 | Formula_Soma_Spot | Nada identificado | Nao |
| 104 | Formula_Somatorio | Nada identificado | Nao |
| 105 | hn.ExtDb | Nada identificado | Nao |
| 106 | hn.ModelType | Nada identificado | Nao |
| 107 | hn.ModelVersion | Nada identificado | Nao |
| 108 | hn.NoUpload | Nada identificado | Nao |
| 109 | hn.RolledForward | Nada identificado | Nao |
| 110 | HTML_CodePage | Nada identificado | Nao |
| 111 | HTML_Control | Nada identificado | Nao |
| 112 | HTML_Controlc | Nada identificado | Nao |
| 113 | HTML_Controlt | Nada identificado | Nao |
| 114 | HTML_Description | Nada identificado | Nao |
| 115 | HTML_Email | Nada identificado | Nao |
| 116 | HTML_Header | Nada identificado | Nao |
| 117 | HTML_LastUpdate | Nada identificado | Nao |
| 118 | HTML_LineAfter | Nada identificado | Nao |
| 119 | HTML_LineBefore | Nada identificado | Nao |
| 120 | HTML_Name | Nada identificado | Nao |
| 121 | HTML_OBDlg2 | Nada identificado | Nao |
| 122 | HTML_OBDlg4 | Nada identificado | Nao |
| 123 | HTML_OS | Nada identificado | Nao |
| 124 | HTML_PathFile | Disco Local do Usuário | Sim |
| 125 | HTML_Title | Nada identificado | Nao |
| 126 | IFRS_Contabil | Nada identificado | Nao |
| 127 | IsColHidden | Nada identificado | Nao |
| 128 | IsLTMColHidden | Nada identificado | Nao |
| 129 | jikjik | Link de Rede não P&C | Sim |
| 130 | jkj | Link de Rede não P&C | Sim |
| 131 | jkjj | Link de Rede não P&C | Sim |
| 132 | jmkjk | Link de Rede não P&C | Sim |
| 133 | just | Link de Rede não P&C | Sim |
| 134 | K2___PARKEDCVW__ | Nada identificado | Nao |
| 135 | K2_ISWBINITED | Nada identificado | Nao |
| 136 | K2_WBEVMODE | Nada identificado | Nao |
| 137 | K2_WBHASINITMODE | Nada identificado | Nao |
| 138 | kjk | Link de Rede não P&C | Sim |
| 139 | kjlhn | Link de Rede não P&C | Sim |
| 140 | limcount | Nada identificado | Nao |
| 141 | LinMenos_1 | Nada identificado | Nao |
| 142 | LinMenos_2 | Nada identificado | Nao |
| 143 | ListOffset | Nada identificado | Nao |
| 144 | lkjhb | Link de Rede não P&C | Sim |
| 145 | lkkljk | Link de Rede não P&C | Sim |
| 146 | llll | Nada identificado | Nao |
| 147 | nkjlç | Link de Rede não P&C | Sim |
| 148 | pacotes20032 | Nada identificado | Nao |
| 149 | pacotes20032c | Nada identificado | Nao |
| 150 | pacotes20032t | Nada identificado | Nao |
| 151 | SAPBEXhrIndnt | Nada identificado | Nao |
| 152 | SAPBEXrevision | Nada identificado | Nao |
| 153 | SAPBEXsysID | Nada identificado | Nao |
| 154 | SAPBEXwbID | Nada identificado | Nao |
| 155 | sencount | Nada identificado | Nao |
| 156 | Slicer_ABERTURA_1 | Nada identificado | Nao |
| 157 | Slicer_ABERTURA_2 | Nada identificado | Nao |
| 158 | Slicer_EMPRESA | Nada identificado | Nao |
| 159 | Slicer_IFRS_CONTABIL | Nada identificado | Nao |
| 160 | Slicer_LINHA_BD | Nada identificado | Nao |
| 161 | SPWS_WBID | Nada identificado | Nao |
| 162 | tai | Link de Rede não P&C | Sim |
| 163 | TesteB | Nada identificado | Nao |
| 164 | TMG | Link de Rede não P&C | Sim |
| 165 | ttttttttt | Nada identificado | Nao |
| 166 | W | Nada identificado | Nao |

## 24.4 Relação de conexões

| Conexão | Tipo | Servidor/Catálogo | Usado por |
|---|---|---|---|
| Conexão principal | SQL Server via ADODB | `SNEPDB24V` / `BPAM` | `Conexoes.bas` (`AbreConexao`), usada por todos os módulos que consultam Hubble e as tabelas mestre |
| Conexão de fallback | SQL Server via ADODB | `SNEPDB24V` / `InfoGER` | `Conexoes.bas` (`AbreConexao`, ramo de fallback — só ativado se a linha `On Error GoTo` correspondente estiver descomentada, ver seção 16.1) |
| Conexão de nível de workbook (Excel) | Nativa do Excel | `connections.bin` — conteúdo não decodificado | [NÃO ACESSÍVEL] — provavelmente ligada ao cache das tabelas dinâmicas de `DP_Rateio` |
| Links externos (3) | Referência de arquivo externo | `externalLink1/2/3.bin` — conteúdo não decodificado | [NÃO ACESSÍVEL] |

Detalhe completo: seção 15.

## 24.5 Índice de regras de negócio

112 regras de negócio catalogadas (RN-001 a RN-115, com faixas reservadas por cluster — ver seção 12 para o detalhe completo de cada uma):

### Cluster ETL / Extração (RN-001 a RN-030)
- RN-001 — Identificador de Fonte por origem de extração
- RN-002 — Limpeza prévia inconsistente entre módulos
- RN-003 — Escala de valores (multiplicador de milhão)
- RN-004 — Centro de Custo inferido por nome de arquivo (Other Income)
- RN-005 — Granularidade de Centro de Custo configurável no Hubble
- RN-006 — Wildcard implícito em filtros de Visão/IFRS
- RN-007 — Tratamento IFRS16 seletivo por fonte
- RN-008 — Formato combinado "KPI > Versão"
- RN-009 — Exclusão interativa de receita para evitar dupla contagem
- RN-010 — Versão default "Divulgado" no Pré-Closing
- RN-011 — Filtros binários S/N reaproveitando o valor como nome de campo
- RN-012 — Zeragem de meses fora do período do cenário (Pré-Closing)
- RN-013 — Visão fixa "IFRS Brasil" nas fontes externas
- RN-014 — Proforma default "w/o Proforma" (com inconsistência de capitalização)
- RN-015 — Validação estrutural obrigatória (RGM/MOCKUP_RGM/Fixed_Revenues)
- RN-016 — Processamento seletivo por tipo de linha de chave ("INF"/"Fórmula")
- RN-017 — Chave de coluna por concatenação Ano+Mês (Fixed Revenues)
- RN-018 — Composição lógica OR/AND configurável via tabela de chaves (Hubble)
- RN-019 — Fontes fora do orquestrador "Extrair Tudo"
- RN-020 — Rastreabilidade de linha original (LIN_BASE) variável por fonte
- RN-021 — Organic informado manualmente em Ajustes
- RN-022 — Vazio vira zero explícito em Ajustes
- RN-023 — Tabela de staging individualizada por usuário (Hubble)
- RN-024 — Classificação de linha por padrão do primeiro token (Other Income)
- RN-025 — Enriquecimento de descrição de conta apenas em RGM/Fixed Revenues
- RN-026 — Zero explícito removido em Consolidada (oposto de RN-022)
- RN-027 — Segmento não recalculado em MOCKUP_RGM/RGM
- RN-028 — Dois modos de limpeza: "Geral" vs. Fonte/Versão específica
- RN-029 — Duas categorias de linha geradas pelo Pré-Closing
- RN-030 — Other Income é a única fonte multi-arquivo (pasta)

### Cluster Core / Cálculo (RN-031 a RN-070)
- RN-031 — Gate de versão obrigatório do Quick Data
- RN-032 — Modo silencioso obrigatório durante processamento em massa
- RN-033 — Limpeza seletiva de histórico por Fonte antes de nova extração
- RN-034 — Identificação do bloco "Base_Consolidada" por múltiplos prefixos
- RN-035 — Cálculo de combinação de meses via fórmula com INDIRECT
- RN-036 — Ordenação da Base por 5 chaves antes de localizar blocos por Campo/Chave
- RN-037 — Geração da visão contábil "IFRS Itália"
- RN-038 — Cascata de 4 fallbacks para resolver Diretoria Gerencial N1-N3
- RN-039 — Exceção hardcoded: Classe Custo N203073156 + prefixo CDC "NT" → código BD 382
- RN-040 — Classificação CLASSE com exceção "Labour Cost" → "w/o Fiber"
- RN-041 — Resolução de EMPRESA via Centro de Custo, sem fallback
- RN-042 — Resolução de Abertura_1 com fallback fixo "Staff"
- RN-043 — Resolução de Segmento com fallback fixo "Others"
- RN-044 — Identificação de linhas elegíveis a rateio
- RN-045 — Rateio fixo vs. rateio variável
- RN-046 — Linha de exclusão/estorno compensatória do rateio
- RN-047 — Filtro de universo de rateio via 5 Slicers
- RN-048 — Ajuste de sinal de Revenues
- RN-049 — Cálculo do FY como soma Jan-Dez
- RN-050 — Zeramento de meses exceto um
- RN-051 — Ajuste de escala (multiplicador)
- RN-052 — Classificação Organic (hoje sempre "Ajusted")
- RN-053 — Peso REF_ORGANIC (1-4) a partir de ORGANIC
- RN-054 — Classificação IFRS Contábil com fallback "w/o IFRS"
- RN-055 — Grupo BD / Linha_BD / Abertura_2-8 via lookup por Classe Custo
- RN-056 — Motor de reclassificação de combinações (Empresa/IFRS/Proforma)
- RN-057 — Critério "ELSE" implícito nas reclassificações
- RN-058 — Critérios de exclusão "<>" isolados em AND, separados dos de igualdade em OR
- RN-059 — Divisão em 2 partes da fórmula de reclassificação de Empresas
- RN-060 — Preenchimento automático de CDC ausente via referência SQL, com marcação de cor
- RN-061 — Idempotência de extração por deleção prévia do bloco da fonte
- RN-062 — Tratamento de erro por linha em Form_Segmentos
- RN-063 — Exceção de preenchimento para linhas de Fonte="Base_Ajustes"
- RN-064 — Padronização "Live" → "UBB" pós-cálculo
- RN-065 — Elegibilidade de rateio por Empresa presente em Lista_Empresa_A1
- RN-066 — Tratamento IFRS16: substituição de CC "DE" por CC "PARA" com estorno
- RN-067 — IFRS16: linhas com dado ausente são puladas e contabilizadas
- RN-068 — Sincronização da grade de combinações a partir de DropComb
- RN-069 — Recalcular reclassificação apenas se detectada mudança real
- RN-070 — Limite de linha hardcoded (40.837) em atualização manual de combinações

### Cluster Refresh / Validação / IFRS16 (RN-071 a RN-096)
- RN-071 — Escopo condicional do tratamento IFRS16
- RN-072 — Filtro de linhas-alvo do IFRS16 por prefixo de origem
- RN-073 — De-Para de Classe Custo (CC) para linhas IFRS16
- RN-074 — Linha sem De-Para de CC é excluída do tratamento
- RN-075 — Marcação da linha original como "DELETAR"
- RN-076 — Sufixo de rastreabilidade "_IFRS16 Tratado"
- RN-077 — De-Para de Centro de Custo (CDC) condicional por Empresa
- RN-078 — Chave de busca do CDC de destino
- RN-079 — Inversão de sinal dos valores IFRS16 tratados
- RN-080 — Substituição (não duplicação) das linhas originais
- RN-081 — Exclusividade entre fluxo IFRS16 e fluxo de reclassificação padrão
- RN-082 — Cascata "Ref Cruzada", prioridade 1: Regra 2 - CC e Diretoria
- RN-083 — Cascata "Ref Cruzada", prioridade 2: Regra 1 - Grupo BD
- RN-084 — Cascata "Ref Cruzada", prioridade 3: Regra 1 - CC
- RN-085 — Cascata "Ref Cruzada", fallback: N.A. / "-"
- RN-086 — Regra "Labour Cost" força CLASSE DESTINO = "W/o Fiber"
- RN-087 — Regra "Rateio" → "Others" em SEG_N2_DESTINO e A1_DESTINO
- RN-088 — Fallback "w/o IFRS" no cálculo de CLASSE CUSTO
- RN-089 — Padronização de nomenclatura de Empresa no De-Para
- RN-090 — Distinção Fiber WTTx vs. Fiber Live
- RN-091 — Limite de 747 itens em listas de validação de Ajustes
- RN-092 — "Total" sempre por último nas listas de validação
- RN-093 — Validação hierárquica automática condicionada a flag (status: possivelmente inerte)
- RN-094 — Calendário de fechamento hardcoded na lista de KPI/Versão
- RN-095 — Cálculo de FY como soma de Jan a Dez
- RN-096 — CENTRO CUSTO calculado por Diretoria N1-N3 + Classe

### Cluster UI / Forms (RN-101 a RN-115)
- RN-101 — Bloqueio de digitação em campos de arquivo (Importação) 
- RN-102 — Confirmação obrigatória antes de importar 
- RN-103 — Renomeação automática de abas duplicadas na importação 
- RN-104 — Tratamento especial para abas "AJUSTES" e "BASE" na importação 
- RN-105 — Exigência de escolha do tipo de exportação 
- RN-106 — Exigência de item selecionado para exportação 
- RN-107 — Sobrescrita silenciosa de arquivo exportado 
- RN-108 — Exclusão automática de linhas EBITDA e "IFRS Itália" na exportação de Base 
- RN-109 — Inconsistência de interação clique-simples vs. duplo-clique entre forms semelhantes 
- RN-110 — Form_Tratamento_Opcoes calcula mas não exibe 
- RN-111 — Menu de comandos do ComboBox1 (Sheet8) por correspondência textual 
- RN-112 — Atalho de abertura de arquivo/pasta por duplo-clique em célula (Sheet8) 
- RN-113 — Seletor de modo/visão via 3 OptionButtons acoplado a célula fixa de outra aba 
- RN-114 — Referência a procedimento inexistente (`Registrar_Sheet`) 
- RN-115 — Mensagem final genérica sem resumo quantitativo 

