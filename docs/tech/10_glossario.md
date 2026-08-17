# 22. Glossário

| Termo | Significado | Evidência/confiança |
|---|---|---|
| KPI | Indicador/métrica reportada (ex.: Receita, EBITDA) | Confirmado — usado em toda a estrutura da Base e das telas de configuração |
| Versão | Cenário orçamentário (Atual, Budget, Forecast, Preview 1-6, Pré-Closing, AJ_Pré-Closing) | Confirmado — lista hardcoded em `Atualizar_Lista_KPI_Versao` (`Refresh_Sup_Linhas.bas`) |
| Centro de Custo (CDC/CC) | Unidade organizacional/contábil de origem do custo | Confirmado |
| Classe Custo | Classificação contábil do tipo de custo | Confirmado |
| Abertura_1 a Abertura_8 | Hierarquia de dimensões gerenciais de classificação de P&L | Confirmado — colunas presentes em Sup_Linhas e na Base |
| Diretoria Gerencial N1/N2/N3 | Hierarquia organizacional (3 níveis) | Confirmado |
| Grupo BD / Linha_BD | Agrupamento de linhas de base de dados para relatórios | Confirmado |
| Organic | Classificação Ajusted/Reported/Recurrent/Normalized, com peso numérico de referência (1-4) | Confirmado — `Form_Organic`/`Form_Ref_Organic` |
| Proforma | Indicador de resultado proforma (ajustado para comparabilidade) | Confirmado |
| IFRS_Contábil | Norma contábil aplicável à linha (IFRS 9/15/16, "w/o IFRS") | Confirmado |
| Rateio | Alocação proporcional de custos indiretos entre empresas/segmentos | Confirmado — `Form_Segmentos` |
| Pré-Closing | Versão de fechamento antecipado, simulada a partir de dados já existentes na Base antes dos números oficiais existirem | Confirmado — `Gerar_Base_Pre_Closing.bas` |
| **Front** | **[VALIDAR COM O NEGÓCIO — parcialmente inferido]**: pelo uso no código (`Form_Importacao`/`Form_Exportacao` tratam "Fronts" como sinônimo de abas/planilhas de relatório inteiras que podem ser copiadas entre arquivos), "Front" parece significar uma aba de relatório/layout — possivelmente a "camada de apresentação" do sistema (em contraste com a "Base", que é a camada de dados). O termo não é definido explicitamente em nenhum comentário do código lido; esta é a melhor inferência possível a partir do uso, não uma definição confirmada. |
| Hubble | Data warehouse/sistema corporativo de origem da extração automatizada via SQL | Confirmado — `LISTA_ARQ_AUX`, `Extracao_SQL_Hubble.bas` |
| BPAM / InfoGER | Nomes dos bancos de dados SQL Server acessados | Confirmado — `Conexoes.bas` |
| De-Para | Tabela de tradução/mapeamento entre uma classificação de origem e uma de destino (ex.: Conta Contábil → Abertura_2) | Confirmado — termo usado recorrentemente no código e nos nomes de arquivo (`Bases_DE_PARA.xlsx`) |
| Ref Cruzada (Referência Cruzada) | Regra de redirecionamento excepcional de responsabilidade entre diretorias/centros de custo, aplicada em cascata de prioridade | Confirmado — `Ref_Cruzada_1`/`Ref_Cruzada_2`, `Limpeza_Base_Ajustes.bas` |
| CDC | Sinônimo de Centro de Custo, usado alternando com "CC" no código | Confirmado |
| LIN_BASE | Marcador de texto usado para localizar dinamicamente a linha de cabeçalho da planilha Base | Confirmado — usado em `Auxiliar.bas`, `Aux_Formulas_Base.bas` |
| DropComb | Nome da aba/estrutura que armazena as combinações válidas de Empresa/IFRS/Proforma usadas em segmentações (slicers) | Confirmado |
| Visão Itália / IFRS Itália | Visão contábil alternativa gerada duplicando linhas da Base conforme regras específicas de uma entidade | Confirmado — `Gerar_Visao_Italia` |
