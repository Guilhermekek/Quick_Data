# 20. Matriz de Rastreabilidade

> Rastreabilidade em nível de fluxo (seção 11), que é o nível prático de uso deste documento para planejamento de teste/impacto. Para o detalhe de cada regra de negócio individual (Regra → procedimento/fórmula/aba exatos), consulte a seção 12 — cada entrada RN-xxx já traz "Local de implementação" e "Evidência" próprios, então repetir isso linha a linha aqui seria redundante com a seção 12 (a especificação pede rastreabilidade, não duplicação de conteúdo já rastreável).

| Fluxo | Regras de negócio relacionadas | Abas envolvidas | Macros/módulos envolvidos | Arquivo de entrada | Saída | Cenário(s) de teste |
|---|---|---|---|---|---|---|
| F-01 Atualização de tabelas mestre | RN-094 (calendário hardcoded) | DP_Segmento, Ref_Cruzada_1/2, CC BD, Sup_Linhas, DropComb, Valid_Lin | `Refresh_DP_Segmento.bas`, `Refresh_De_X_Para.bas`, `Refresh_Sup_Linhas.bas`, `Refresh_Drop_Comb.bas` | SQL Server (BPAM), `Bases_DE_PARA.xlsx` | Tabelas mestre atualizadas | CT-01, CT-07, CT-08 |
| F-02 Extração de fonte individual | RN-001 a RN-030 (parcial, conforme fonte), RN-056 a RN-059 (reclassificação, todas as fontes) | Extração, Base, tabela de chave da fonte | Módulo `Extracao_*` correspondente | Arquivo externo/SQL da fonte | Linhas na Base | CT-02, CT-03, CT-04, CT-09 |
| F-03 Extração de todas as bases | Todas as RN-001 a RN-030, RN-031 (gate de versão) | Todas as envolvidas em F-02, para as 7 fontes | `Extrair_Todas_as_Bases` (`Auxiliar.bas`) + todos os `Extracao_*` | Todas as fontes externas | Base totalmente reconstruída | CT-05, CT-10 |
| F-04 Ajustes manuais | RN-021, RN-022, RN-082 a RN-090 (cascata de destino) | Ajustes, Sup_Linhas, Ref_Cruzada_1/2, DP_Segmento | `Extracao_Sheet_Ajustes.bas`, `Limpeza_Base_Ajustes.bas`, `Lista_Validacao_Ajustes.bas` | Digitação manual | Linhas na Base (Fonte=Base_Ajustes) | CT-06 |
| F-05 Tratamento IFRS16 | RN-071 a RN-081 | Base, Sup_Linhas, Aux_IFRS16 | `fx_IFRS16.bas` | Base já extraída (fonte Base_1009) | Linhas reclassificadas na Base | CT-11 |
| F-06 Geração de Pré-Closing | RN-010, RN-012, RN-029 | Preview (não coberta em detalhe nesta análise), Base | `Gerar_Base_Pre_Closing.bas` | Base já extraída | Novas linhas na Base (Fonte=QD AUTOMATIC) | CT-12 |
| F-07 Importação/Exportação | RN-101 a RN-109 | Base, planilhas de layout ("Fronts") | `Form_Importacao.frm`, `Form_Exportacao.frm` | Arquivo `.xlsb` externo | Arquivo(s) `.xlsb` gerado(s) ou abas importadas | CT-13, CT-14 |
| F-08 Saneamento de Defined Names | — (manutenção, não regra de negócio funcional) | ListDefinedNames | `TK_Functions.bas` (`CLEAR_Defined_Names`) | — | Nomes definidos removidos | CT-15 |
| Motor de rateio (transversal a F-02/F-03) | RN-044 a RN-047, RN-065 | Base, DP_Rateio (Sh_Rateio) | `Form_Segmentos` (`Aux_Formulas_Base.bas`) | Base + tabela de percentuais | Linhas desdobradas + linha de estorno | CT-16 |
| Validação de hierarquia (Main Results) | RN-093 | Main Results, Valid_Lin | `BackupCodigo_MainResults.bas` (status incerto) | — | Lista de validação aplicada à célula | CT-17 |
