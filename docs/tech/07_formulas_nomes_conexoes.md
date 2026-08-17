# 15. Fórmulas, Nomes Definidos e Conexões

## 15.1 Fórmulas críticas

**[CONFIRMADO/INFERIDO — evidência: código-fonte de `Aux_Formulas_Base.bas`, `Auxiliar.bas`, `Limpeza_Base_Ajustes.bas`]** A grande maioria das fórmulas de negócio do sistema **não são fórmulas estáticas de planilha** — são geradas dinamicamente em tempo de execução como strings VBA (`FormulaR1C1`), aplicadas em massa a um range, calculadas, e então "congeladas" em valores (`.Copy` → `PasteSpecial xlPasteValues`). Isso significa que abrir o arquivo e olhar uma célula da Base **não mostra a fórmula original que a gerou** — ela já foi convertida em valor. As fórmulas-modelo mais importantes, reconstruídas a partir do código VBA que as gera:

| Campo calculado | Padrão de fórmula (reconstruído do VBA) | Módulo gerador |
|---|---|---|
| Diretoria Gerencial N1-N3 | Cascata de 4 níveis `IFERROR(INDEX/MATCH(...), IFERROR(INDEX/MATCH(...), IFERROR(INDEX/MATCH(...), INDEX(...))))` | `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (`Aux_Formulas_Base.bas`) |
| Classe Custo (Ajustes) | `INDEX/MATCH` em Sup_Linhas por `IFRS_Contabil & Abertura_2..8`, fallback `"w/o IFRS"` | `Comando_Refazer_Formula` (`Limpeza_Base_Ajustes.bas`) |
| Reclassificação de combinação (Empresa/IFRS/Proforma) | `IF(AND(critério1,critério2,...),"S","N")` construído dinamicamente a partir de uma tabela de critérios | `Criar_Formula_Filtros` (`Auxiliar.bas`) |
| FY (total anual) | Soma simples `=SUM(Jan:Dez)` | `Form_Calcular_FY` (`Aux_Formulas_Base.bas`) |

## 15.2 Fórmulas voláteis

**[CONFIRMADO]** Uso confirmado de `INDIRECT()` dentro de fórmulas geradas em `Calcular_Comb_Meses`/`Calcular_Comb_Meses_Intervalo`/`Calcular_Comb_Meses_Intervalo_Linha` (`Auxiliar.bas`) — função volátil (recalcula a cada alteração de qualquer célula da planilha, não só das suas dependências diretas), aplicada sobre a Base inteira antes de ser convertida em valor. Isso é uma causa direta de lentidão durante o cálculo, mesmo sendo temporário.

## 15.3 Nomes definidos

**[CONFIRMADO — evidência: planilha de auditoria própria `ListDefinedNames`, 166 linhas]** O workbook contém **166 nomes definidos**. A própria planilha de auditoria classifica cada um por tipo de link e sinaliza **42 deles (25%) como candidatos a exclusão** ("APAGAR? = Sim"), nas categorias:
- "Link de Rede não P&C" — a maioria, apontando para caminhos de rede de sistemas/pastas de usuários específicos, muitos datados de 1998 a 2014 (ex.: referências a "Stock Performance.xls", relatórios de "Depósito Cível", planilhas ligadas a um sistema OLAP antigo com prefixo `EV__`/`K2_` — nomenclatura típica de complementos Essbase/Hyperion, hoje descontinuados).
- "Link de SharePoint" com erro de referência.
- "Disco Local do Usuário" (caminhos `C:\...` de máquinas específicas de ex-funcionários).

Nomes definidos usados ativamente pelo sistema hoje (não sinalizados para exclusão) seguem o padrão `CB_*` (ex.: `CB_Empresa`, `CB_Meses`, `CB_KPI`, `CB_N1_Gerencial`) — alimentam listas de validação (dropdowns) via fórmulas `OFFSET`/`COUNTA` contra as tabelas mestre.

## 15.4 Links externos

**[CONFIRMADO — quantidade; NÃO ACESSÍVEL — conteúdo]** O workbook registra **3 links externos** (`externalLink1.bin`, `externalLink2.bin`, `externalLink3.bin` no pacote). O conteúdo binário (para qual arquivo cada um aponta) não foi decodificado nesta análise — ver seção 4.4.

## 15.5 Consultas e conexões

**[CONFIRMADO]** O workbook tem uma definição de conexões a nível de Excel (`connections.bin`, conteúdo binário não decodificado — usada tipicamente para alimentar a fonte de uma tabela dinâmica). As conexões efetivamente usadas pelo VBA para consultar dados **não passam por essa camada declarativa do Excel** — são abertas e fechadas em tempo de execução via `ADODB.Connection` dentro de `Conexoes.bas` (`AbreConexao`/`FechaConexao`), string de conexão fixa no código-fonte, contra o servidor `SNEPDB24V` (bancos `BPAM` e `InfoGER`, fallback).

## 15.6 Tabelas dinâmicas

**[CONFIRMADO — existência; NÃO ACESSÍVEL — configuração de campos]** 2 tabelas dinâmicas (`pivotTable1.bin`, `pivotTable2.bin`), ambas associadas à aba `DP_Rateio` (via relacionamento no pacote OOXML), com cache definido em `pivotCacheDefinition1.bin`. Os filtros/segmentações (slicers) associados foram identificados pelos nomes dos shapes em `drawing7.xml`: `EMPRESA`, `ABERTURA_1`, `LINHA_BD`, `ABERTURA_2`, `IFRS_CONTABIL`.

## 15.7 Dependências e riscos ligados a este item

- O acúmulo de 166 nomes definidos (42 já reconhecidos como lixo pelo próprio time) é causa clássica de lentidão de abertura/salvamento e, em casos extremos, corrupção de arquivos `.xlsb`/`.xlsx` antigos — ver seção 18.
- O padrão "fórmula → calcular → congelar em valor" significa que **auditar uma regra de negócio exige ler o código VBA, não a planilha** — a fórmula que gerou o valor atual de uma célula específica não está mais lá.
- `INDIRECT()` volátil sobre a Base inteira é um contribuinte direto de lentidão em qualquer recálculo, mesmo fora do fluxo de extração.
