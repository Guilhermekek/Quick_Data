# 13. Arquivos de Entrada

**Nota geral:** para as 6 fontes abaixo cujo caminho é apontado manualmente pelo usuário, não há um **nome de arquivo obrigatório** — o sistema aceita qualquer nome que o usuário selecione no diálogo. A identificação do arquivo correto depende inteiramente do usuário escolher o arquivo certo; o VBA não confere o nome do arquivo contra um padrão esperado.

## 13.1 Base 1009

| Campo | Valor |
|---|---|
| Nome/padrão esperado | Nenhum padrão de nome exigido pelo código; conteúdo esperado: relatório contábil por Classe/Centro de Custo (**[VALIDAR COM O NEGÓCIO]** para o nome real do sistema de origem) |
| Extensão | `.xlsx`/`.xls` (implícito — sem checagem explícita de extensão no código) |
| Abas obrigatórias | Aba com cabeçalho localizável por busca textual `Cells.Find("Jan*")` |
| Colunas obrigatórias | Centro Custo, Classe Custo, Exercício, meses (Jan-Dez) |
| Ordem das colunas | Não é fixa por posição — localizada por busca de texto do cabeçalho |
| Tipos de dados | Meses devem ser numéricos; há tratamento de encoding quebrado hardcoded (`"Mar‡o"` → `"Março"`) |
| Regras de preenchimento | [NÃO IDENTIFICADO] |
| Limites identificados | Nenhum limite de linha hardcoded identificado neste módulo especificamente |
| Tratamento de linhas vazias/duplicadas/inválidas | Linhas zeradas são removidas (`Apagar_Linhas_Zeradas`); linhas de "Other Income" são removidas mediante confirmação do usuário |
| Comportamento fora do padrão | Se o cabeçalho "Jan*" não for encontrado, `Cells.Find` retorna `Nothing` — comportamento resultante não tratado explicitamente (risco de erro 91, objeto não definido) |
| Validações existentes | Nenhuma validação de estrutura antes de processar, além da busca do cabeçalho |
| Validações ausentes | Não valida se as colunas de meses contêm apenas números; não valida se o arquivo é realmente da fonte "1009" |
| Evidência | `Extracao_Base_1009.bas`, subs `Processo_Extrair_Base_1009` |

## 13.2 RGM / MOCKUP RGM / Fixed Revenues

| Campo | Valor |
|---|---|
| Nome/padrão esperado | Nenhum padrão de nome exigido |
| Extensão | `.xlsx`/`.xls` |
| Abas obrigatórias | Abas com nome composto dinamicamente por KPI (sufixos `_Act`, `_Bdg`, `_Fcst##`) — **se o KPI não mapear em nenhum dos `If` de tradução, a aba não é encontrada e o erro só aparece depois, com mensagem genérica** |
| Colunas obrigatórias | Definidas por uma tabela de "chaves" interna (Sheet7/Sheet14 para RGM, Sheet22 para MOCKUP, Sheet19/Sheet20 para Fixed Revenues) que mapeia linha a linha do arquivo fonte |
| Ordem das colunas | Localizada por chave/rótulo, não por posição fixa de coluna, mas a **linha** de cada chave é fixa/validada estruturalmente |
| Tipos de dados | Numérico nas colunas de valor; escala aplicada (×1.000.000 — valores no arquivo fonte estão em milhões) |
| Regras de preenchimento | [NÃO IDENTIFICADO] |
| Limites identificados | Array de chaves de tamanho fixo `Campos_Chaves(10000)` — mais de 10.000 chaves estoura o array (erro de índice) |
| Tratamento de linhas vazias/duplicadas/inválidas | [NÃO IDENTIFICADO] |
| Comportamento fora do padrão | **RGM e MOCKUP RGM**: se o rótulo da linha no arquivo fonte não bater com o esperado pela tabela de chaves, o sistema dispara `MsgBox` crítico e executa `End` — interrompe todo o VBA abruptamente (ver risco crítico, seção 18). **Fixed Revenues**: mesma validação, mas sem o `End` (inconsistência entre módulos "irmãos" no tratamento do mesmo tipo de erro) |
| Validações existentes | Validação estrutural linha a linha contra a tabela de chaves (comparação de texto exato — frágil a espaço extra) |
| Validações ausentes | Nenhum tratamento de erro ao abrir o arquivo em si (arquivo ausente/corrompido) |
| Evidência | `Extracao_Base_RGM.bas`, `Extracao_Base_MOCKUP_RGM.bas`, `Extracao_Fixed_Revenues.bas`, sub privada `Processo_Extracao_Sheet_Base` (implementada de forma quase idêntica e duplicada nos 3 módulos) |

## 13.3 Other Income

| Campo | Valor |
|---|---|
| Nome/padrão esperado | Um arquivo por operadora/entidade; Centro de Custo atribuído por 8 padrões de nome de arquivo hardcoded (TBP, CTCEL, TPAR, FIBER_RS, FIBER_RR, FIBER_OG, METIS_CZ, INTELIG_OG; fallback para CDC `XPHR999` se nenhum padrão bater) |
| Extensão | `.xlsx` |
| Abas obrigatórias | Aba "Contas de resultados" |
| Colunas obrigatórias | Coluna com Classe Custo (primeiro token antes do espaço) e colunas de meses |
| Ordem das colunas | Localizada por busca de cabeçalho ("Contas de resultados") |
| Tipos de dados | [NÃO IDENTIFICADO] |
| Regras de preenchimento | [NÃO IDENTIFICADO] |
| Limites identificados | Nenhum limite de linha hardcoded identificado neste módulo |
| Tratamento de linhas vazias/duplicadas/inválidas | Remove linhas de totais/não numéricas; remove linhas de classificações específicas (Market/Process/Labour/Volume Driven Costs) e, mediante confirmação, "Net Service Revenues" |
| Comportamento fora do padrão | Se o nome do arquivo não casar com nenhum dos 8 padrões, o Centro de Custo é atribuído silenciosamente ao valor padrão `XPHR999` — **classificação incorreta sem aviso ao usuário** |
| Validações existentes | Nenhuma validação de que a pasta selecionada contém arquivos válidos além do teste de padrão de nome |
| Validações ausentes | Sem tratamento de erro se um arquivo da pasta não tiver a aba/estrutura esperada |
| Evidência | `Extracao_Base_Other_Inco.bas` |

## 13.4 Base Consolidada

| Campo | Valor |
|---|---|
| Nome/padrão esperado | Nenhum padrão exigido |
| Extensão | `.xlsx`/`.xls` |
| Colunas obrigatórias | Conta Contábil, Centro de Custo, Valor Prévia |
| Regras de preenchimento | Linhas com Conta Contábil/CC/Valor em branco são removidas |
| Tratamento de linhas vazias/duplicadas/inválidas | Remove linhas com valor zero (loop decrescente linha a linha); mediante confirmação, remove linhas de "Product Revenues"/"Net Service Revenues" (evitar dupla contagem com outras fontes) |
| Comportamento fora do padrão | `On Error Resume Next` cobre 3 operações de exclusão de linhas em branco sem checar sucesso individualmente (ver seção 16) |
| Evidência | `Extracao_Base_Consolidad.bas` |

## 13.5 Quick Data (extração sistema-a-sistema)

| Campo | Valor |
|---|---|
| Nome/padrão esperado | Nenhum padrão exigido; espera-se estrutura de colunas já compatível com a Base (Empresa, Centro Custo, Classe Custo, Exercício, Tipo Nível 2, Organic, Proforma, Abertura_1, Segmento, meses) |
| Extensão | `.xlsx`/`.xls` |
| Colunas obrigatórias | As listadas acima, localizadas por cabeçalho "FONTE" |
| Comportamento fora do padrão | Se uma coluna esperada não existir no arquivo fonte, `fn.Match` retorna erro `#N/A` **não tratado** — quebra a macro sem mensagem clara |
| Evidência | `Extracao_Base_Quick_Data.bas` — descrito no próprio código como o módulo mais simples do conjunto |

## 13.6 Ajustes (entrada interna)

| Campo | Valor |
|---|---|
| Origem | Digitação direta do usuário na planilha Ajustes (Sheet13), não é um arquivo externo |
| Colunas obrigatórias | Versão, Exercício, Centro Custo, Classe Custo, meses, e colunas de destino calculadas automaticamente (`A1_DESTINO`, `SEG_N2_DESTINO` etc.) |
| Tratamento de vazios | Célula vazia em mês é convertida para "0" via `Selection.Replace` |
| Validações existentes | Listas de validação (dropdowns) reconstruídas por `Lista_Validacao_Ajustes.bas`, com limite hardcoded de 747 itens por lista (limite técnico do Excel) — itens excedentes são **descartados silenciosamente** |
| Validações ausentes | Nenhuma validação de que o restante da linha seja numérico além da troca de vazio por "0" |
| Evidência | `Extracao_Sheet_Ajustes.bas`, `Lista_Validacao_Ajustes.bas`, `Limpeza_Base_Ajustes.bas` |

---

# 14. Arquivos e Dados de Saída

## 14.1 Exportação de Front(s)

| Campo | Valor |
|---|---|
| Nome/padrão | `FRONT_QD - <data_hora>.xlsb` (arquivo único) ou um arquivo por aba selecionada |
| Destino | Pasta escolhida pelo usuário via diálogo nativo |
| Conteúdo | Abas selecionadas, com fórmulas convertidas em valores, campos de menu de seleção removidos |
| Processo responsável | `Form_Exportacao` (modo "Front") |
| Regras de sobrescrita | **Sobrescreve arquivo existente silenciosamente** (`Kill` antes de salvar, sem aviso ao usuário) |
| Condições de sucesso/falha | `MsgBox` final único "Processo concluído com sucesso!", sem resumo de quantos arquivos/onde |

## 14.2 Exportação de Base

| Campo | Valor |
|---|---|
| Nome/padrão | `BASE_QD - <data_hora>.xlsb` (arquivo único) ou um arquivo por combinação KPI/Versão/Ano |
| Destino | Pasta escolhida pelo usuário |
| Conteúdo | Base filtrada/ordenada pelas combinações escolhidas, colunas desnecessárias removidas, linhas de EBITDA/IFRS Itália removidas, meses convertidos em valores |
| Processo responsável | `Form_Exportacao` (modo "Base") |
| Regras de sobrescrita | Mesma observação do item 14.1 |

## 14.3 Log de erros (saída interna)

| Campo | Valor |
|---|---|
| Destino | Planilha `tk_Lista_de_erros`, dentro do próprio workbook |
| Conteúdo | Função, linha VBA, aba/linha/coluna Excel, valor de erro, usuário (`Environ("Username")`), timestamp |
| Processo responsável | `fn_ListAllErrors` (`TK_Functions.bas`) |
| Dependências | Nenhuma — é gravação local, não gera arquivo externo |
