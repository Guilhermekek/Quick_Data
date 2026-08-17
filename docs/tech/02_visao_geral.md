# 5. Visão Geral do Quick Data 3.23

## 5.1 Objetivo aparente da solução

**[INFERIDO — evidência: nomes de abas, macros e estrutura de dados]** O Quick Data 3.23 é uma ferramenta de consolidação e reporting financeiro/gerencial. Ela reúne dados contábeis vindos de múltiplas fontes internas e externas (um data warehouse corporativo chamado "Hubble" via SQL Server, e uma série de relatórios em Excel produzidos por outros sistemas/times), os traduz para uma estrutura gerencial comum (Diretoria, Segmento, Abertura contábil, Classe de Custo), aplica regras contábeis específicas (rateio de custos indiretos entre empresas, tratamento de leasing sob IFRS16, geração de uma versão de fechamento antecipado) e disponibiliza o resultado consolidado para relatórios internos (planilha "Main Results" e um painel com tabelas dinâmicas, "DP_Rateio").

## 5.2 Entradas

**[CONFIRMADO — evidência: `Aux_Leitura_Nome_Arqs.bas`, módulos `Extracao_*`]**

| Entrada | Mecanismo |
|---|---|
| Base Hubble | Consulta SQL Server automatizada (sem seleção manual de arquivo) |
| Base 1009 | Arquivo `.xlsx` externo, selecionado manualmente pelo usuário |
| Base RGM | Arquivo externo, seleção manual |
| Base MOCKUP RGM | Arquivo externo, seleção manual |
| Base Fixed Revenues | Arquivo externo, seleção manual |
| Base Other Income | Pasta com múltiplos arquivos, seleção manual da pasta |
| Base Consolidada | Arquivo externo, seleção manual |
| Base Quick Data (sistema-a-sistema) | Arquivo externo já no layout de destino, seleção manual |
| Ajustes manuais | Digitação direta na planilha "Ajustes" pelo usuário |
| Tabelas mestre de mapeamento (Sup_Linhas, DP_Segmento, Ref_Cruzada, DropComb) | SQL Server + um arquivo externo `Bases_DE_PARA.xlsx` |

## 5.3 Processamentos

**[CONFIRMADO/INFERIDO — evidência: módulos `Aux_Formulas_Base.bas`, `Auxiliar.bas`, `fx_IFRS16.bas`, `Gerar_Base_Pre_Closing.bas`]** Após a extração, cada linha bruta passa por: (1) resolução de dimensões gerenciais via cascatas de lookup contra as tabelas mestre; (2) desdobramento de rateio de custos indiretos entre empresas/segmentos; (3) reclassificação contábil condicional (combinações Empresa/IFRS/Proforma, marcadas como "S"/"N" via fórmula gerada dinamicamente); (4) tratamento específico de IFRS16 (reclassificação de custo de leasing com inversão de sinal), condicionado a uma flag configurável; (5), sob demanda, geração de uma versão sintética "Pré-Closing" recombinando dados já existentes na base.

## 5.4 Saídas

**[CONFIRMADO]** A saída primária é a própria planilha "Base" consolidada (usada por relatórios internos ao workbook: "Main Results", "DP_Rateio"). O sistema também permite **exportar** a Base filtrada ou abas de relatório inteiras ("Fronts") para novos arquivos `.xlsb`, e **importar** Fronts de arquivos externos para dentro do workbook atual (via `Form_Exportacao`/`Form_Importacao`).

## 5.5 Usuários ou perfis envolvidos

**[VALIDAR COM O NEGÓCIO]** Não há, no código analisado, nenhum mecanismo de login, permissão por usuário, ou distinção de papéis (não há tabela de usuários nem checagem de identidade além de `Environ("Username")` usado só para nomear uma tabela SQL temporária e para registrar quem gerou um erro no log `tk_Lista_de_erros`). É razoável inferir, pela natureza da ferramenta (planejamento & controle financeiro), que existam ao menos dois perfis de uso — quem extrai/consolida dados e quem consome os relatórios — mas o arquivo não impõe essa distinção tecnicamente. Isso precisa ser confirmado com o negócio.

## 5.6 Fluxo operacional principal

**[INFERIDO — evidência: nomes e ordem de dependência dos botões da aba "Extração", ver seção 6]**

1. Usuário abre o arquivo e navega até a aba de controle ("Extração").
2. Atualiza as tabelas mestre de mapeamento ("Atualizar bases auxiliares").
3. Aponta manualmente (via diálogo de arquivo) o caminho de cada fonte externa que for usar no mês.
4. Extrai cada fonte (botões individuais) ou todas de uma vez ("Extrair todas as bases").
5. Insere ajustes manuais na planilha "Ajustes", se necessário.
6. Opcionalmente gera a versão "Pré-Closing".
7. Consulta os relatórios ("Main Results", "DP_Rateio") ou exporta os dados/relatórios para outro arquivo.
