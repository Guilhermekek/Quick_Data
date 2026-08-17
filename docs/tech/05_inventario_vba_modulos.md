# 9. Inventário do Projeto VBA

**Fatos confirmados de estrutura geral (evidência: stream `PROJECT` do `vbaProject.bin`, decompilação via `oletools`):**
- 70 componentes de código no total: 33 módulos de classe de planilha (`Sheet1`–`Sheet33`, com gaps — ver 7.1), 1 `ThisWorkbook`, 3 UserForms (`Form_Importacao`, `Form_Exportacao`, `Form_Tratamento_Opcoes`), e 33 módulos padrão (`.bas`).
- 192 procedimentos (`Sub`/`Function`/`Property`) no total, contados mecanicamente por assinatura de declaração.
- O projeto está protegido por senha para visualização no VBE (evidência: campos `CMG=`/`DPB=`/`GC=` no stream `PROJECT`), mas isso não impediu a extração do código-fonte por ferramenta de parsing direto (ver seção 4.3).

## 9.1 ThisWorkbook

| Campo | Valor |
|---|---|
| Nome | ThisWorkbook |
| Tipo | Módulo de documento (workbook) |
| Responsabilidade | Eventos globais do workbook |
| Procedimentos existentes | `Workbook_Open` (1 procedimento, **vazio** — nenhuma instrução dentro dele) |
| Dependências | Nenhuma |
| Variáveis globais/públicas | Nenhuma |
| Possíveis efeitos colaterais | Nenhum (evento vazio) |
| Ponto de entrada | `Workbook_Open` dispararia automaticamente ao abrir o arquivo, mas está vazio — **não há nenhuma rotina de inicialização automática confirmada** (sem checagem de versão automática, sem splash screen, sem carregamento de estado inicial) |

## 9.2 Módulos de classe de planilha (code-behind)

| CodeName | Aba correspondente (ver seção 7) | Procedimentos | Responsabilidade |
|---|---|---|---|
| Sheet8 | Extracao [confirmado] | `ComboBox1_Change`, `Worksheet_BeforeDoubleClick` (2) | Painel/home do sistema: combo de "comandos rápidos" (dispara `CLEAR_Defined_Names` ou `UPDATE_Combinacoes_Empresas` conforme o texto selecionado) e atalho de duplo-clique para abrir arquivo/pasta referenciada em célula |
| Sheet3 | Base [confirmado] | 3 `OptionButton_Click` (3) | Seletor de "modo/visão" simples — grava a legenda escolhida em `Sheet11!AV6` e alterna negrito entre os três botões |
| Sheet1, Sheet2, Sheet4, Sheet5, Sheet6, Sheet7, Sheet9, Sheet10, Sheet11, Sheet12, Sheet13, Sheet14, Sheet15, Sheet16, Sheet17, Sheet18, Sheet19, Sheet20, Sheet21, Sheet22, Sheet23, Sheet24, Sheet25, Sheet26, Sheet27, Sheet28, Sheet29, Sheet30, Sheet31, Sheet33 | Ver mapeamento parcial na seção 7 | 0 (módulos vazios) | Nenhuma — nenhum evento de planilha implementado; interatividade dessas abas (se houver) vem de fórmulas/validações/botões ligados a módulos `.bas`, não de código de evento próprio |

## 9.3 Módulos padrão (.bas)

| Módulo | Linhas | Procedimentos | Responsabilidade (resumo — detalhe completo na seção 10) |
|---|---|---|---|
| Auxiliar | 1.120 | 28 | Orquestrador central: macro-mestre de extração completa, controle de ciclo de vida (versão, otimizações), reclassificação de combinações Empresa/IFRS/Proforma, geração da visão Itália |
| TK_Functions | 953 | 26 | Biblioteca de utilitários e manutenção: log de erros, limpeza de defined names/estilos, sincronização de combinações a partir de DropComb, preenchimento de CDC via SQL — contém versões `_TK` duplicadas de regras já existentes em `Auxiliar.bas` |
| Aux_Formulas_Base | 960 | 20 | Fábrica de fórmulas da planilha Base: resolve dimensões gerenciais via lookup, implementa o motor de rateio de custos indiretos (`Form_Segmentos`) |
| Form_Importacao | 782 | 11 | Código-behind do UserForm de importação de Fronts (ver seção 9.4) |
| Form_Exportacao | 615 | 11 | Código-behind do UserForm de exportação de Fronts/Base (ver seção 9.4) |
| Extracao_SQL_Hubble | 604 | 9 | Extração automatizada via SQL dinâmico contra o data warehouse Hubble |
| Refresh_Sup_Linhas | 537 | 4 | Reconstrói a tabela mestre Sup_Linhas e listas de KPI/Versão a partir do SQL Server |
| Extracao_Base_1009 | 482 | 3 | Importação e normalização da "Base 1009" (relatório contábil externo) |
| Extracao_Base_MOCKUP_RGM | 435 | 4 | Extração da base de simulação RGM |
| Extracao_Base_Other_Inco | 429 | 5 | Extração de "Other Income" — varre pasta com múltiplos arquivos por operadora |
| Extracao_Base_RGM | 428 | 4 | Extração da base RGM "real" |
| Extracao_Fixed_Revenues | 424 | 4 | Extração de receitas de serviços fixos |
| Gerar_Base_Pre_Closing | 383 | 2 | Reclassificação interna: gera versão sintética de fechamento antecipado |
| Extracao_Base_Consolidad | 358 | 5 | Importação da "Base Consolidada" (prévia contábil por Conta Contábil × CC) |
| Extracao_Base_Quick_Data | 332 | 3 | Importação de base já pré-formatada no layout Quick Data (extração sistema-a-sistema) |
| Extracao_Sheet_Ajustes | 309 | 4 | Extração dos ajustes manuais da planilha Ajustes — único módulo 100% interno, sem I/O externo |
| Refresh_De_X_Para | 268 | 3 | Importa tabelas de responsabilidade cruzada de um arquivo externo |
| Limpeza_Base_Ajustes | 264 | 4 | Reset e reconstrução de fórmulas da planilha Ajustes |
| BackupCodigo_MainResults | 211 | 3 | Validação automática em cascata de dropdowns do Main Results — **[VALIDAR COM O NEGÓCIO]**: nome do módulo e estrutura sugerem que está desativado (ver seção 18) |
| fx_IFRS16 | 174 | 3 | Tratamento contábil de leasing sob IFRS16 |
| Front_Processos | 159 | 2 | Rotinas de manutenção reutilizáveis chamadas a partir da planilha ativa (reaplicar fórmula padrão, atualizar validação — esta última **incompleta**, ver seção 18) |
| Lista_Validacao_Ajustes | 145 | 2 | Reconstrói listas de validação (dropdowns) da planilha Ajustes |
| Aux_Leitura_Nome_Arqs | 132 | 9 | Utilitário de diálogos de seleção de arquivo/pasta — grava caminho escolhido nas células de configuração da aba Extração |
| Module2 | 60 | 4 | [NÃO IDENTIFICADO — não coberto em profundidade nesta análise; conteúdo aparenta ser utilitário menor] |
| Conexoes | 59 | 2 | Módulo de conexão SQL Server — contém credenciais em texto plano (ver seção 17 e 18, achado crítico de segurança) |
| Refresh_DP_Segmento | 90 | 1 | Atualiza a tabela DP_Segmento a partir do SQL Server |
| Refresh_Drop_Comb | 41 | 1 | Copia a aba DropComb de um arquivo externo |
| Module8 | 21 | 2 | Macro gravada (`Macro1`, `Macro2`) — código de teste/depuração deixado no projeto; insere fórmula `VLOOKUP` fixa em célula específica |
| Module10 | 15 | 1 | Macro gravada (`Macro3`) — código de teste/depuração; insere fórmula `HLOOKUP`/`VLOOKUP` fixa em células específicas |
| Module1 | 11 | 1 | Macro gravada (`seleciona_pra_baixo`) — código de teste/depuração; seleciona um range fixo |
| Module3 | 11 | 1 | Macro gravada (`teste_apagar_linha`) — código de teste/depuração; limpa conteúdo de uma linha fixa (linha 20130, número hardcoded) |
| Module5 | 4 | 1 | `Sub listDefinedNames()` — **vazia**, sem instruções |
| Module6 | 2 | 0 | Módulo vazio (só `Attribute VB_Name`) |
| Module7 | 2 | 0 | Módulo vazio |
| Module4 | 1 | 0 | Módulo vazio |
| Module9 | 1 | 0 | Módulo vazio |

**Observação transversal:** os módulos `Module1`, `Module3`, `Module4`, `Module5`, `Module6`, `Module7`, `Module8`, `Module9`, `Module10` (9 módulos, ~68 linhas no total) são, com alta confiança, **resíduo de gravação de macro e/ou depuração** — evidência: comentários padrão do gravador de macro (`' Macro1 Macro`), atributo `VB_Invoke_Func` característico de macro gravada, nomes genéricos (`teste_apagar_linha`, `listDefinedNames` vazia), e ausência de qualquer referência a eles a partir dos módulos de produção (a confirmar via grep cruzado na seção 10). Não representam funcionalidade do sistema em produção, mas permanecem no projeto — risco de manutenção classificado na seção 18.

## 9.4 UserForms

| UserForm | Procedimentos | Responsabilidade |
|---|---|---|
| Form_Importacao | 11 | Interface para localizar arquivo externo e importar uma ou mais abas ("Fronts") para dentro do workbook atual |
| Form_Exportacao | 11 | Interface para exportar Fronts ou a Base filtrada para arquivo(s) externo(s) |
| Form_Tratamento_Opcoes | 3 | Tela de consulta somente leitura da configuração de extrações habilitadas — **funcionalidade aparentemente incompleta** (dados calculados nunca chegam à interface visível, só a `Debug.Print`) |

Detalhamento completo de campos/controles: ver catálogo de procedimentos (seção 10) e seção 4.4 quanto à limitação de leitura da definição visual binária (`.frx`).

## 9.5 Referências e bibliotecas

Ver seção 17 (Configurações e Dependências Técnicas) — extraídas mecanicamente do stream `VBA/dir` do `vbaProject.bin`.
