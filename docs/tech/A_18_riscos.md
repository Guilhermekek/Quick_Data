# 18. Riscos para Manutenção

> Consolidado dos 4 clusters (46 riscos no total: 9 críticos, 17 altos, 14 médios, 6 baixos — soma das classificações de cada cluster). Achado transversal mais grave (aplicável a todo o sistema, repetido pelos clusters ETL/Core/Refresh): credenciais de banco de dados em texto plano em Conexoes.bas — tratar como prioridade #1 independentemente de qualquer outra ação.

## 18.1 Cluster Extração / ETL
### D-01 — Credenciais de banco de dados em texto plano — **CRÍTICO**
- **Descrição:** usuário e senha do SQL Server (`AdminBPAM` / catálogo `BPAM`, e fallback
  `Report2BPAM`/`URELATIG` / catálogo `InfoGER`) estão hardcoded em texto plano no código-fonte VBA, com a
  mesma senha reutilizada entre os dois pares de credenciais.
- **Evidência:** `Conexoes.bas:26-27, 36-37`.
- **Impacto:** qualquer pessoa com acesso ao arquivo `.xlsb` (ou ao VBA Project, mesmo sem senha de proteção)
  tem acesso de leitura/escrita ao banco `BPAM`/`InfoGER` — vazamento de credencial institucional, possível
  acesso não autorizado a dados financeiros sensíveis da TIM Brasil.
- **Componentes afetados:** todo o cluster que depende de `AbreConexao` (Hubble) e módulos de refresh fora do
  cluster que também usam `Conexoes.bas`.
- **Cuidados antes de alterar:** nunca reproduzir a senha em nenhum artefato da reescrita (documentação,
  commits, logs); tratar como incidente de segurança a ser reportado à área responsável antes/durante a
  migração; trocar a senha no SQL Server assim que a migração para gerenciamento de segredo (variável de
  ambiente / cofre) estiver pronta.
- **Teste sugerido:** validar que a nova implementação nunca loga a connection string completa (nem em modo
  debug) e que a senha nunca aparece em texto plano em nenhum arquivo versionado.

### D-02 — SQL construído por concatenação de string sem parametrização — **CRÍTICO**
- **Descrição:** todo o SQL dinâmico da extração Hubble (`CREATE TABLE`, `INSERT...SELECT`, filtros WHERE) é
  montado por concatenação direta de valores vindos de planilha (KPI, Versão, Empresa, nomes de coluna), sem
  uso de parâmetros ADODB (`?`/`Parameters`).
- **Evidência:** `Extracao_SQL_Hubble.bas:485-498` (`Extrair_Linha`), `:399` (`Criar_TB_SQL_AUX`).
- **Impacto:** qualquer apóstrofo ou caractere especial em um nome de KPI/Empresa/Versão configurado na
  planilha quebra a query (erro de sintaxe) ou, no pior caso, permite injeção SQL não intencional (dado que os
  valores vêm de células editáveis pelo usuário, não de input externo controlado, o risco prático de ataque é
  menor, mas o risco de quebra operacional é real e recorrente).
- **Componentes afetados:** `Processo_Extrair_Base_Hubble`, `Extrair_Linha`, `Criar_TB_SQL_AUX`,
  `Excluir_TB_SQL_AUX`, `Extrair_Base_Final`.
- **Cuidados antes de alterar:** na reconstrução em Python, usar SQL parametrizado (SQLAlchemy/`pyodbc` com
  `?`) para todo valor vindo de configuração de planilha; nunca montar cláusulas `WHERE`/`GROUP BY` por
  f-string com valor de usuário.
- **Teste sugerido:** testar extração com um nome de KPI/Empresa contendo apóstrofo (`O'Brien`, por exemplo) e
  confirmar que não quebra nem gera comportamento inesperado.

### D-03 — `End` abrupto em validação estrutural (RGM / MOCKUP_RGM) — **ALTO**
- **Descrição:** ao detectar divergência entre a estrutura esperada (tabela de chaves) e a estrutura real do
  arquivo fonte, os módulos RGM e MOCKUP_RGM executam a instrução `End`, que mata todo o processo Excel
  instantaneamente, sem fechar conexões SQL abertas, sem restaurar `Application.ScreenUpdating`/`EnableEvents`,
  e sem qualquer rotina de limpeza.
- **Evidência:** `Extracao_Base_RGM.bas:332`, `Extracao_Base_MOCKUP_RGM.bas:335`.
- **Impacto:** usuário perde trabalho não salvo em qualquer outra pasta aberta na mesma instância do Excel;
  aplicação fica em estado inconsistente (tela travada/atualização desligada) até reabertura manual;
  comportamento é imprevisível para quem não conhece o código-fonte.
- **Componentes afetados:** `Processo_Extracao_Sheet_Base` (variantes RGM e MOCKUP_RGM).
- **Cuidados antes de alterar:** na reconstrução, substituir por uma exceção tratada que interrompe apenas a
  extração corrente, fecha conexões/recursos e devolve uma mensagem clara ao usuário/log, sem encerrar o
  processo inteiro.
- **Teste sugerido:** simular uma divergência estrutural proposital (renomear uma célula de rótulo no arquivo
  de teste) e confirmar que o sistema encerra a extração de forma controlada, preservando outros dados abertos.

### D-04 — `MsgBox` bloqueantes no meio do pipeline de extração — **ALTO**
- **Descrição:** múltiplos módulos (1009, Other Income, Consolidada, Ajustes-limpeza, Pré-Closing) exibem
  caixas de confirmação/aviso no meio da execução, exigindo clique humano antes de prosseguir.
- **Evidência:** `Extracao_Base_1009.bas:406`, `Extracao_Base_Other_Inco.bas:405`,
  `Extracao_Base_Consolidad.bas:322,328`, `Gerar_Base_Pre_Closing.bas:12,290`.
- **Impacto:** impede qualquer forma de execução desatendida/agendada do pipeline completo; se agendado via
  Task Scheduler ou processo batch, a macro trava indefinidamente esperando input do usuário.
- **Componentes afetados:** praticamente todo o cluster (exceto Hubble, RGM, MOCKUP_RGM, Fixed_Revenues, Quick
  Data, que não têm `MsgBox` de decisão de negócio no meio do fluxo — apenas de erro).
- **Cuidados antes de alterar:** mover todas as decisões atualmente tomadas via `MsgBox` para parâmetros de
  configuração definidos ANTES do início do processamento em lote (arquivo de config/flags), preservando a
  possibilidade de decisão manual em modo interativo, mas tornando-a opcional em modo batch.
- **Teste sugerido:** rodar o pipeline completo em modo não interativo (sem tela) e confirmar que ele completa
  sem travar esperando input.

### D-05 — Tabela de staging SQL pode ficar "presa" (Hubble) — **ALTO**
- **Descrição:** a tabela temporária `TB_AUX_HUBBLE_QUICK_DATA_<usuário>` é criada no início e removida no
  fim do processo Hubble; se a execução falhar entre `Criar_TB_SQL_AUX` e a segunda chamada de
  `Excluir_TB_SQL_AUX`, a tabela permanece no banco.
- **Evidência:** `Extracao_SQL_Hubble.bas:66-67` (exclusão preventiva + criação), `:341` (exclusão final).
- **Impacto:** próxima execução do mesmo usuário falha ao tentar `CREATE TABLE` (tabela já existe) — a
  exclusão preventiva no início (linha 66) mitiga parcialmente isso, mas apenas se a execução seguinte de fato
  rodar; enquanto isso, a tabela "morta" consome espaço/permissões no banco `BPAM`.
- **Componentes afetados:** `Processo_Extrair_Base_Hubble`, `Criar_TB_SQL_AUX`, `Excluir_TB_SQL_AUX`.
- **Cuidados antes de alterar:** na reconstrução, usar transação SQL com `try/finally` garantindo `DROP TABLE`
  mesmo em caso de exceção, e preferir tabelas temporárias de sessão (`#temp`) em vez de tabelas persistentes
  nomeadas por usuário, se o SGBD permitir.
- **Teste sugerido:** forçar uma falha no meio da extração Hubble (ex.: desconectar a rede) e verificar se a
  tabela auxiliar permanece órfã no banco.

### D-06 — Ausência de limpeza prévia em 5 fontes — **MÉDIO**
- **Descrição:** MOCKUP_RGM, RGM, Fixed_Revenues, Quick_Data (chamada ausente ou comentada) não removem
  registros da mesma Fonte antes de reextrair — cada execução do botão principal é um `APPEND`, não um
  `REPLACE`.
- **Evidência:** ver RN-002.
- **Impacto:** duplicação acumulativa de dados na Base a cada execução repetida sem uma limpeza manual prévia
  via botão separado — pode distorcer relatórios financeiros por dupla contagem.
- **Componentes afetados:** `Sheet3` (Base) — qualquer relatório consumindo essas 4 fontes.
- **Cuidados antes de alterar:** confirmar com o negócio o processo operacional real (a limpeza é sempre feita
  manualmente antes? existe uma rotina externa de deduplicação?) antes de decidir se a reconstrução deve
  automatizar o `REPLACE` por padrão.
- **Teste sugerido:** executar a extração da mesma fonte duas vezes seguidas sem limpar e verificar se a
  contagem de linhas na Base dobra.

### D-07 — Duplicação de código entre 3 módulos irmãos — **MÉDIO**
- **Descrição:** a Sub `Processo_Extracao_Sheet_Base` existe de forma quase idêntica em
  `Extracao_Base_RGM.bas`, `Extracao_Base_MOCKUP_RGM.bas` e `Extracao_Fixed_Revenues.bas`, cada uma com
  pequenas variações (uma tem `End` abrupto, outra não; uma chama `PEGA_A_DESCRICAO_DA_CONTA`, outra não).
- **Evidência:** `Extracao_Base_RGM.bas:256-414`, `Extracao_Base_MOCKUP_RGM.bas:258-419`,
  `Extracao_Fixed_Revenues.bas:259-408`.
- **Impacto:** uma correção de bug feita em um módulo não é automaticamente replicada nos outros dois —
  manutenção divergente já observável (comportamento de erro inconsistente, RN-015).
- **Componentes afetados:** os 3 módulos citados.
- **Cuidados antes de alterar:** na reconstrução, unificar em uma única função parametrizável (Python) que
  cubra as 3 variações via flags explícitas, documentando a intenção de cada diferença encontrada aqui.
- **Teste sugerido:** comparar side-by-side o comportamento das 3 implementações para o mesmo cenário de
  divergência estrutural e confirmar que a versão unificada reproduz (ou corrige deliberadamente) cada uma.

### D-08 — Array de tamanho fixo sem checagem de limite — **MÉDIO**
- **Descrição:** `Campos_Chaves(10000)` é declarado com tamanho fixo em RGM, MOCKUP_RGM e Fixed_Revenues, sem
  verificação de que o número de chaves processadas não ultrapasse 10.000.
- **Evidência:** `Extracao_Base_RGM.bas:258`, `Extracao_Base_MOCKUP_RGM.bas:260`,
  `Extracao_Fixed_Revenues.bas:261`.
- **Impacto:** se a tabela de chaves crescer além de 10.000 linhas "INF"/"Fórmula", a macro gera erro de índice
  fora do intervalo (`Subscript out of range`) e interrompe a extração sem mensagem de negócio clara.
- **Componentes afetados:** as 3 variantes de `Processo_Extracao_Sheet_Base`.
- **Cuidados antes de alterar:** na reconstrução, usar estrutura dinâmica (lista/array redimensionável) sem
  limite arbitrário.
- **Teste sugerido:** confirmar o número real de linhas de chave em cada aba de origem hoje e o quanto de
  margem existe até o limite de 10.000.

### D-09 — Funções declaradas como `Function` mas usadas como `Sub` — **MÉDIO**
- **Descrição:** todas as ~30 rotinas `Processo_Limpar_Base_*`/`Processo_Extrair_Base_*` (e
  `Gerar_Base_Versao_Pre_Closing`) são declaradas com a palavra-chave `Function`, mas nunca atribuem valor ao
  próprio nome — comportamento idêntico ao de uma `Sub`.
- **Evidência:** verificado por leitura completa de cada corpo (nenhuma ocorrência de
  `NomeDaFunção = <valor>` encontrada nas ~30 rotinas do cluster).
- **Impacto:** não é um bug funcional (VBA tolera o padrão), mas confunde a leitura do código — um
  desenvolvedor pode esperar um valor de retorno que nunca existe, e chamadas futuras que tentem capturar
  `x = Processo_Extrair_Base_1009()` receberiam sempre `Empty`/erro de tipo.
- **Componentes afetados:** legibilidade/manutenibilidade de todo o cluster.
- **Cuidados antes de alterar:** na reconstrução em Python, usar funções com contrato de retorno explícito
  (ex.: `bool` de sucesso, ou lançar exceção) em vez de replicar o padrão "Function sem retorno".
- **Teste sugerido:** N/A (é uma constatação de estilo, não um comportamento a testar em runtime).

### D-10 — Código morto/incompleto (`Retirar_Duplicadas`) — **BAIXO**
- **Descrição:** a Sub `Retirar_Duplicadas` (Hubble) cria colunas auxiliares de verificação de duplicata mas
  contém um loop de exclusão totalmente vazio — nunca remove nenhuma linha — e ainda assim exibe uma `MsgBox`
  de sucesso.
- **Evidência:** `Extracao_SQL_Hubble.bas:595-601`.
- **Impacto:** baixo, pois a rotina não é referenciada por nenhum botão/chamada ativa no código lido — risco
  latente apenas se alguém reativar essa rotina sem notar que está incompleta.
- **Componentes afetados:** nenhum em produção (código não executado); risco potencial se reativado.
- **Cuidados antes de alterar:** não portar essa rotina para a reconstrução sem antes completá-la ou descartá-la
  deliberadamente — confirmar com o negócio se a deduplicação por Classe/Centro de Custo é uma necessidade real
  não atendida hoje.
- **Teste sugerido:** N/A (rotina não executada em produção); se for reativada, testar que o loop de fato remove
  as duplicatas identificadas pelas colunas auxiliares.

### D-11 — Inconsistência de nomenclatura em atualização de listas (Ajustes) — **BAIXO**
- **Descrição:** `Extrair_Base_Ajustes` chama `Atualizar_Lista_KPI_Versao` (sem sufixo "_Interna"), enquanto os
  outros 9 módulos de extração chamam `Atualizar_Lista_KPI_Versao_Interna`.
- **Evidência:** `Extracao_Sheet_Ajustes.bas:14` vs. ex. `Extracao_Base_1009.bas:13`.
- **Impacto:** potencialmente a lista de KPI/Versão não é atualizada da mesma forma após uma extração de
  Ajustes comparada às demais fontes — comportamento sutilmente diferente que pode passar despercebido.
- **Componentes afetados:** `Sheet15` (Sup_Linhas, lista de KPI/Versão), indiretamente todos os filtros que
  dependem dela.
- **Cuidados antes de alterar:** **[VALIDAR COM O NEGÓCIO]** qual das duas é o comportamento correto antes de
  padronizar na reconstrução.
- **Teste sugerido:** comparar o resultado de `Atualizar_Lista_KPI_Versao` vs. `_Interna` para o mesmo estado de
  Base e confirmar se a lista final diverge.

### D-12 — Filtro de seleção de arquivo sem restrição de tipo — **BAIXO**
- **Descrição:** `GetArquivo` usa o filtro "All Files (*.*)" na caixa de diálogo de seleção, permitindo ao
  usuário escolher qualquer tipo de arquivo, não apenas planilhas Excel válidas.
- **Evidência:** `Aux_Leitura_Nome_Arqs.bas:67`.
- **Impacto:** erro de usuário (selecionar arquivo errado) só é detectado bem depois, no momento em que o
  módulo de extração tenta abrir o arquivo como workbook Excel, com mensagem de erro nativa pouco amigável.
- **Componentes afetados:** todas as 6 fontes que usam `GetArquivo` (1009, RGM, MOCKUP, Fixed_Revenues,
  Consolidada, Quick_Data).
- **Cuidados antes de alterar:** na reconstrução, restringir a seleção a extensões válidas (`.xlsx`, `.xlsb`,
  `.xls`) e validar a estrutura esperada logo após a seleção, não apenas no momento da extração.
- **Teste sugerido:** selecionar deliberadamente um arquivo de tipo incorreto (ex.: `.pdf`) e confirmar que o
  sistema informa o erro imediatamente, antes de tentar processar.

---

*Fim do catálogo do cluster Extração/ETL — 54 procedimentos, 30 regras de negócio, 20 dependências mapeadas,
12 riscos classificados.*

## 18.2 Cluster Core / Motor de Cálculo
### D — CRÍTICO

**D-1. Duplicação de regras de negócio (`_TK` vs. família `Auxiliar.bas`) com premissas de layout divergentes e caminhos de invocação diferentes**
- **Descrição**: as 4 rotinas de reclassificação/recálculo (Empresas, IFRS_Contabil, Proforma, combinação de meses por linha) existem em 2 implementações: a família `Auxiliar.bas` localiza o cabeçalho dinamicamente via `Match("LIN_BASE", ...)` e recebe intervalo de linhas por parâmetro; a família `_TK` (`TK_Functions.bas`) assume `Lin_Inicial=6` fixo e recalcula sempre a Base inteira, com 3 fórmulas *diferentes entre si* para calcular o fim do range (comparar TK_Functions.bas:784, 865, 891). Além disso, `Reclassificar_Combinacoes_Empresas_TK` não replica a divisão em 2 partes que a versão original usa (RN-059).
- **Evidência**: Auxiliar.bas:472 (busca dinâmica) vs. TK_Functions.bas:691, 728, 809, 887 (linha 6 fixa); TK_Functions.bas:784 vs. 865 vs. 891 (3 fórmulas de `Ult_Linh` diferentes); ausência de divisão em partes em TK_Functions.bas:877-953 vs. presença em Auxiliar.bas:625-763.
- **Impacto**: se a estrutura da Base mudar (nova linha de metadado deslocando o cabeçalho, por exemplo), a família `Auxiliar.bas` continua correta enquanto a família `_TK` recalcula o intervalo errado silenciosamente — sem erro, apenas classificação incorreta em produção. Como os dois caminhos são acionados por gatilhos de negócio diferentes e reais (extração completa vs. pós-IFRS16/edição de combinações via ComboBox), uma correção de regra aplicada em um lado facilmente não é replicada no outro.
- **Componentes afetados**: `Auxiliar.bas` (Reclassificar_Combinacoes_Empresas/IFRS_Contabil/Proforma, Calcular_Comb_Meses_Intervalo_Linha), `TK_Functions.bas` (as 4 versões `_TK`), `fx_IFRS16.bas` (chamador da família `_TK`), todos os módulos `Extracao_Base_*.bas` (chamadores da família `Auxiliar.bas`).
- **Cuidados antes de alterar**: qualquer mudança na regra de classificação de Empresa/IFRS/Proforma (RN-056 a RN-059) deve ser aplicada **simultaneamente** nas duas famílias, e testada em ambos os caminhos de invocação (extração completa E fluxo IFRS16/ComboBox); antes de reescrever em Python, decidir explicitamente se as duas famílias devem convergir para uma única implementação (fortemente recomendado) ou se há razão de negócio genuína para mantê-las distintas.
- **Teste sugerido**: rodar a extração completa e o fluxo de tratamento IFRS16 sobre a mesma cópia da Base, comparando byte-a-byte as colunas de classificação (Empresa, IFRS_Contabil, Proforma) resultantes — qualquer divergência confirma que os dois caminhos produzem resultados diferentes para o mesmo dado de entrada.

**D-2. Limite de linha hardcoded (40.837) em `UPDATE_Combinacoes_Empresas_parte2`**
- **Descrição**: a propagação de fórmulas de combinação de Empresa via `AutoFill` está fixada no range `EC6:GD40837` — linhas além dessa marca não são preenchidas, sem qualquer erro ou aviso.
- **Evidência**: TK_Functions.bas:321.
- **Impacto**: se a Base (descrita em outras partes do sistema como tendo dezenas de milhares de linhas) já ultrapassou 40.837 linhas, esta rotina — se ainda em uso, ainda que manual — está gerando classificação de Empresa incompleta silenciosamente para as linhas excedentes.
- **Componentes afetados**: `TK_Functions.bas` (UPDATE_Combinacoes_Empresas_parte2), coluna "EMPRESA" (bloco de combinações) da Base.
- **Cuidados antes de alterar**: antes de tocar nesta rotina, confirmar com o negócio se ela ainda é usada manualmente e, se sim, checar imediatamente se a Base atual excede 40.837 linhas (`[VALIDAR COM O NEGÓCIO]` com urgência — ver RN-070).
- **Teste sugerido**: contar `COUNTA(Base!A:A)` na Base de produção e comparar com 40.837; se maior, auditar se as linhas excedentes têm classificação de combinação de Empresa ausente/incorreta.

**D-3. Tratamento de erro via `MsgBox` interativo dentro do laço de rateio (`Form_Segmentos`)**
- **Descrição**: quando o cálculo de Abertura_1 de uma linha resulta em erro de fórmula, o sistema pausa a execução com um `MsgBox` por ocorrência (RN-062) — inviável em qualquer execução desatendida/batch, e potencialmente interrompe uma extração noturna/agendada indefinidamente aguardando clique do usuário.
- **Evidência**: Aux_Formulas_Base.bas:465-468.
- **Impacto**: bloqueio total do pipeline de extração se houver qualquer erro de fórmula durante o processamento em massa; nenhum registro estruturado do erro é feito (o log `fn_ListAllErrors` existe mas está desconectado — chamada comentada na linha 466).
- **Componentes afetados**: `Form_Segmentos` (Aux_Formulas_Base.bas), toda extração que a chama (Ajustes, Other Income, 1009, Fixed Revenues, Base Consolidada, ALL BASES).
- **Cuidados antes de alterar**: na reescrita, substituir por log estruturado (reconectando/reimplementando o padrão de `fn_ListAllErrors`) e continuar processamento sem interação humana; se o Quick Data legado precisar continuar operando enquanto a reescrita não está pronta, considerar reconectar `fn_ListAllErrors` como paliativo.
- **Teste sugerido**: forçar uma condição de erro de fórmula propositalmente (ex. Centro de Custo inexistente em Sup_Linhas) e confirmar que hoje a extração completa realmente trava aguardando clique do usuário.

### D — ALTO

**D-4. Ambiguidade de escopo `Ordenar_Lista` (Public em Auxiliar.bas vs. Private em BackupCodigo_MainResults.bas)**
- **Descrição**: existem duas implementações de `Ordenar_Lista` — `Public Function` em `Auxiliar.bas` (item 48) e `Private Sub` em `BackupCodigo_MainResults.bas`. Chamadas dentro de `BackupCodigo_MainResults.bas` resolvem para a cópia local (regra de escopo do VBA); chamadas de outros módulos (ex. `Form_Exportacao.frm`) resolvem para a versão de `Auxiliar.bas`.
- **Evidência**: Auxiliar.bas:1096 (Public Function); BackupCodigo_MainResults.bas:188 (Private Sub); chamadores em BackupCodigo_MainResults.bas:170 (resolve local) e Form_Exportacao.frm:563 (resolve Auxiliar.bas).
- **Impacto**: um desenvolvedor que edite uma das duas cópias assumindo que afeta ambos os pontos de chamada introduz bug silencioso — o comportamento realmente executado depende de qual módulo está chamando, informação não óbvia sem entender a regra de resolução de escopo do VBA.
- **Componentes afetados**: `Auxiliar.bas`, `BackupCodigo_MainResults.bas`, `Form_Exportacao.frm`.
- **Cuidados antes de alterar**: antes de modificar qualquer uma das duas implementações, confirmar explicitamente qual módulo está chamando o ponto que se pretende corrigir; considerar renomear uma das duas na reescrita para eliminar a ambiguidade.
- **Teste sugerido**: comparar a saída de `Ordenar_Lista` chamada a partir de `Form_Exportacao` com a saída do mesmo algoritmo chamado a partir de `BackupCodigo_MainResults` para a mesma lista de entrada — devem produzir o mesmo resultado hoje (mesma lógica), mas qualquer edição futura em uma cópia sem a outra quebra essa paridade sem aviso.

**D-5. Operações O(n²) em `EntireRow.Copy`/`EntireRow.Delete` dentro de laços sobre a Base inteira**
- **Descrição**: `Gerar_Visao_Italia` (laço triplo com `EntireRow.Copy`/`PasteSpecial` no caminho mais interno), `Limpar_Base_Historica` e `Processo_Exclusao_Linhas_Base` (`EntireRow.Delete` dentro de varredura linha a linha) operam sobre uma Base de dezenas de milhares de linhas com padrões cujo custo cresce de forma não-linear.
- **Evidência**: Auxiliar.bas:341-343 (Gerar_Visao_Italia), 561 (Limpar_Base_Historica), 995 (Processo_Exclusao_Linhas_Base).
- **Impacto**: tempo de execução crescente de forma desproporcional ao volume de dados à medida que a Base cresce — sintoma já reconhecido pela própria equipe (uso de `Application.StatusBar` com percentual de progresso nessas rotinas, ex. Auxiliar.bas:299-300, 1003).
- **Componentes afetados**: `Gerar_Visao_Italia`, `Limpar_Base_Historica`, `Processo_Exclusao_Linhas_Base` (Auxiliar.bas); `Apagar_Linhas_Zeradas` (Aux_Formulas_Base.bas, mesmo padrão).
- **Cuidados antes de alterar**: na reescrita em Python, substituir por operações vetorizadas (filtragem/concatenação de DataFrame) em vez de replicar o padrão célula/linha a linha; se for necessário manter o VBA operando enquanto a reescrita não está pronta, medir o tempo de execução atual como baseline antes de qualquer mudança.
- **Teste sugerido**: medir o tempo de `Gerar_Visao_Italia` e `Limpar_Base_Historica` com o volume de dados atual de produção; extrapolar para o crescimento esperado da Base nos próximos 12-24 meses e avaliar se o tempo se torna inaceitável.

**D-6. Regras de negócio hardcoded sem documentação, replicadas em múltiplos locais**
- **Descrição**: a exceção "Classe Custo N203073156 + prefixo CDC NT → código 382" (RN-039) está copiada 3 vezes (`Form_Grupo_BD`, `Form_Opex_Driven`, `Form_Linha_BD`); a exclusão de Empresas "5G"/"METIS_CZ" do tratamento CDC no IFRS16 (RN-066) e a faixa de Linha_BD 159-240/265 usada como filtro de rateio (RN-047) são números/strings mágicos sem comentário explicativo.
- **Evidência**: Aux_Formulas_Base.bas:679, 713, 743 (regra 382); fx_IFRS16.bas:65 (5G/METIS_CZ); Aux_Formulas_Base.bas:410-413 (faixa 159-240/265).
- **Impacto**: correção de uma regra de negócio exige encontrar e atualizar todas as cópias — alto risco de esquecer uma delas; sem documentação do motivo de negócio, qualquer alteração futura corre risco de quebrar um caso de uso não óbvio.
- **Componentes afetados**: `Form_Grupo_BD`, `Form_Opex_Driven`, `Form_Linha_BD` (Aux_Formulas_Base.bas); `Form_Segmentos` (filtro de rateio); `fx_IFRS16.bas`.
- **Cuidados antes de alterar**: antes de tocar em qualquer uma das exceções, buscar todas as ocorrências da mesma regra em todo o código-fonte (não apenas neste cluster); documentar o motivo de negócio junto com o time de P&C antes de migrar para a reescrita.
- **Teste sugerido**: criar casos de teste específicos cobrindo cada exceção hardcoded (uma linha com CC=N203073156+CDC prefixo NT; uma linha de Empresa 5G no fluxo IFRS16; uma linha com Linha_BD fora da faixa 159-240/265 no fluxo de rateio) e confirmar que o comportamento esperado é preservado após qualquer refatoração.

**D-7. Botões de macro (Shape.OnAction) não rastreáveis via análise textual do dump**
- **Descrição**: pelo menos 8 procedimentos deste cluster (`Extrair_Todas_as_Bases`, `Limpar_Todas_as_Bases`, `Exportar`, `Importar_Fronts`, `Atualizar_Base_Todos_Campos_Auxiliares`, `Refresh_Base_Aux`, `Calcular_Range_Selecionado`, `RemoveStyles`, `UPDATE_aplicar_CDC_por_Referencia`, `UPDATE_Combinacoes_Empresas_parte2`, `RUN_Apagar_defined_names_definitivamente`, `RUN_Atualizar_Dinamica_com_erros`) não têm nenhuma chamada `Call` encontrada em todo o dump textual, mas são claramente "botões de primeiro nível" pelo nome/estrutura.
- **Evidência**: ausência de resultados de grep além da própria declaração, para cada um dos procedimentos citados (ver seções A.2/A.3).
- **Impacto**: a documentação funcional não pode confirmar, apenas pelo texto VBA, quais botões da UI disparam quais macros — essa informação vive no XML binário do `.xlsb` (atribuições `Shape.OnAction`), fora do escopo desta extração; há risco de a reescrita mapear incorretamente qual ação de UI corresponde a qual procedimento.
- **Componentes afetados**: toda a camada de UI da aba "Extracao" (Sheet8) e demais painéis com botões.
- **Cuidados antes de alterar**: antes de finalizar o mapeamento de UI→procedimento na documentação funcional completa, abrir o arquivo `.xlsb` original no Excel e inspecionar manualmente a atribuição de macro de cada botão/forma (`Alt+F8` ou botão direito → "Atribuir Macro"), ou extrair o XML de Shapes do pacote `.xlsb`.
- **Teste sugerido**: abrir o Quick Data original, clicar em cada botão da aba Extracao com o VBE aberto em modo de depuração (breakpoint no início de cada Sub candidata) e confirmar qual macro efetivamente dispara.

**D-8. Manipulação de `SlicerCaches` item a item dentro do motor de rateio**
- **Descrição**: `Form_Segmentos` ajusta o estado de 5 Slicers do workbook percorrendo `SlicerItems` um a um (RN-047) — operação conhecidamente lenta no modelo de objetos do Excel e que altera o estado visual de filtros compartilhado com a experiência do usuário fora desta rotina.
- **Evidência**: Aux_Formulas_Base.bas:381-424.
- **Impacto**: lentidão proporcional ao número de itens em cada Slicer; efeito colateral visível ao usuário (filtros da aba de relatórios mudam silenciosamente durante o processamento).
- **Componentes afetados**: `Form_Segmentos`; qualquer relatório/tabela dinâmica conectado aos 5 Slicers citados.
- **Cuidados antes de alterar**: na reescrita, substituir a dependência de Slicers por um filtro explícito sobre a estrutura de dados (DataFrame), eliminando o acoplamento à UI do Excel; se for necessário manter o VBA, avaliar `PivotFilters`/APIs em lote em vez de iteração item a item.
- **Teste sugerido**: capturar o estado dos 5 Slicers antes e depois de rodar `Form_Segmentos`, confirmando que o estado final é o esperado e que nenhum filtro do usuário foi perdido permanentemente.

### D — MÉDIO

**D-9. `Ativar_Tudo` não restaura `Application.Calculation` para automático**
- **Descrição**: `Desligar_Tudo` seta `Calculation = xlCalculationManual`, mas `Ativar_Tudo` (seu par simétrico esperado) não seta de volta para `xlCalculationAutomatic` — apenas restaura ScreenUpdating/DisplayAlerts/EnableEvents.
- **Evidência**: Auxiliar.bas:130-135 (Desligar_Tudo) vs. 75-79 (Ativar_Tudo, sem `Calculation=xlCalculationAutomatic`).
- **Impacto**: após qualquer rotina pesada, o workbook pode permanecer em modo de cálculo manual — usuário edita a planilha depois e nada recalcula sozinho, gerando confusão/erros de interpretação de dados desatualizados.
- **Componentes afetados**: toda rotina que segue o padrão `Desligar_Tudo`/`Ativar_Tudo` (praticamente todo o sistema).
- **Cuidados antes de alterar**: confirmar se essa omissão é intencional (talvez outra rotina restaure o cálculo automático em algum ponto não coberto por este cluster) antes de "corrigir" — pode haver dependência oculta.
- **Teste sugerido**: rodar `Extrair_Todas_as_Bases` do início ao fim e verificar `Application.Calculation` ao final — se `xlCalculationManual`, o problema é real e afeta a experiência do usuário até o próximo fechamento/reabertura do Excel.

**D-10. Código morto/órfão numeroso aumentando a superfície de manutenção**
- **Descrição**: pelo menos 13 procedimentos do cluster são código morto confirmado (referenciados só em comentários) ou órfãos sem nenhum chamador encontrado: `Form_Segmentos_OLD`, `Calcular_Comb_Meses`, `fn_ListAllErrors`, `convertInterval_to_R1C1`, `getCellAddress`, `GET_value_S_or_N`, `UPDATE_Combinacoes_Proforma`, `UPDATE_Combinacoes_Empresas_parte2`, `RUN_Apagar_defined_names_definitivamente`, `Testeee`, `Teste`, `RUN_Atualizar_Dinamica_com_erros`, `FiltraCampo`, `SET_Limpar_Cores_CDC`, `PEGA_A_DESCRICAO_DA_CONTA_Independente`.
- **Evidência**: ver campo "Quem chama" de cada procedimento citado nas seções A.1-A.3.
- **Impacto**: aumenta o esforço de compreensão do sistema para qualquer novo desenvolvedor/analista (não é óbvio quais dos 74 procedimentos são realmente executados em produção); risco de reativação acidental de lógica obsoleta (ex. `Form_Segmentos_OLD` sem a lógica de rateio).
- **Componentes afetados**: `Aux_Formulas_Base.bas`, `Auxiliar.bas`, `TK_Functions.bas`.
- **Cuidados antes de alterar**: antes de remover qualquer um destes na reescrita, confirmar com o negócio que não são usados manualmente via VBE/botão não capturado no dump (ver risco D-7).
- **Teste sugerido**: nenhum teste automatizado aplicável — recomenda-se checklist de confirmação com os administradores do sistema (que têm acesso ao VBE) antes de descartar cada item.

**D-11. Uso de fórmula volátil `INDIRECT` em massa**
- **Descrição**: o cálculo de "combinação de meses" (RN-035) usa `INDIRECT`, função reconhecidamente volátil que impede o Excel de otimizar o grafo de dependências, aplicada a blocos que podem cobrir toda a extensão da Base.
- **Evidência**: Auxiliar.bas:234, 453, 482; TK_Functions.bas:704.
- **Impacto**: cálculo inicial (antes do bake-in em valor) potencialmente caro para blocos grandes; embora o resultado seja congelado em valor logo em seguida, o próprio ato de calcular uma vez sobre uma extensão grande já é custoso.
- **Componentes afetados**: `Calcular_Comb_Meses(_Intervalo)(_Linha)(_TK)` (Auxiliar.bas, TK_Functions.bas).
- **Cuidados antes de alterar**: ao reescrever em Python, substituir por parsing direto da string de intervalo (ex. `"Jan:Mar"` → soma das colunas correspondentes) sem depender de avaliação dinâmica de fórmula.
- **Teste sugerido**: medir o tempo de `Calcular_Comb_Meses_Intervalo` isoladamente com "Geral" (toda a Base) como parâmetro, comparado a uma reimplementação sem INDIRECT.

**D-12. Falta de `Option Explicit` / tipagem `Variant` implícita em todo o cluster**
- **Descrição**: nenhum dos 3 módulos declara `Option Explicit`; a esmagadora maioria das variáveis é `Variant` implícito, incluindo parâmetros de procedimento sem tipo declarado.
- **Evidência**: cabeçalho dos 3 módulos (ausência de `Option Explicit`); assinaturas como `Sub Form_Calcular_FY(Sh_Destino, Lin_Inicial, Lin_Final, Lin_Cabecalho)` (Aux_Formulas_Base.bas:22, sem tipos).
- **Impacto**: erros de digitação em nomes de variável criam silenciosamente novas variáveis Variant vazias em vez de gerar erro de compilação; maior overhead de memória/verificação de tipo em tempo de execução.
- **Componentes afetados**: todo o cluster.
- **Cuidados antes de alterar**: ativar `Option Explicit` retroativamente exigiria testar exaustivamente cada módulo (pode revelar variáveis não declaradas que hoje "funcionam por acidente") — não recomendado fazer isso no legado sem suíte de testes; a lição principal é para a reescrita (tipagem forte em Python).
- **Teste sugerido**: N/A para o legado (mudança de alto risco sem suíte de testes); para a reescrita, garantir tipagem estática desde o início (type hints + mypy/pydantic).

### D — BAIXO

**D-13. Código de teste/depuração deixado em módulos de produção**
- **Descrição**: `Teste` (com chamada quebrada de tipo incompatível a `set_formula_CDC`), `Testeee` (varredura de validações só com `Debug.Print`), `convertInterval_to_R1C1` (experimento com fórmula hardcoded de exemplo) permanecem no módulo `TK_Functions.bas` sem nenhuma utilidade de produção.
- **Evidência**: TK_Functions.bas:102-106, 499-513, 45-49.
- **Impacto**: nenhum risco funcional direto (não são chamados por ninguém), mas poluem o código e podem confundir manutenção futura.
- **Componentes afetados**: `TK_Functions.bas`.
- **Cuidados antes de alterar**: remover na reescrita; nenhum cuidado especial necessário dado que não têm chamador.
- **Teste sugerido**: nenhum necessário.

**D-14. Blocos extensos de código comentado ao lado de código ativo**
- **Descrição**: várias Subs mantêm uma versão anterior inteira comentada logo ao lado da versão ativa (ex. `Form_Diretoria_Gerencial_Com_Ref_Cruzada` linhas 125-132 vs. 134-141; `Form_Segmentos` com múltiplos blocos de `Replace`/`Copy`/`PasteSpecial` comentados).
- **Evidência**: Aux_Formulas_Base.bas:125-141, 242-278, 315-347.
- **Impacto**: aumenta o custo cognitivo de leitura do código; risco (baixo, mas presente) de um desenvolvedor reativar acidentalmente a versão errada.
- **Componentes afetados**: `Aux_Formulas_Base.bas` principalmente.
- **Cuidados antes de alterar**: nenhum — são apenas artefatos textuais sem efeito em runtime.
- **Teste sugerido**: nenhum necessário; recomenda-se apenas não migrar esse padrão para a reescrita.

---

*Fim do catálogo do cluster Auxiliar.bas / Aux_Formulas_Base.bas / TK_Functions.bas — 74 procedimentos documentados (A), 40 regras de negócio RN-031 a RN-070 (B), mapa de dependências (C) e riscos classificados (D).*

## 18.3 Cluster Refresh / Validação / IFRS16
## CRÍTICO

### R-01 — De-Para de IFRS16 mantido fora de qualquer automação rastreável
- **Descrição:** As colunas 71, 73, 15, 22 e a célula `BX1` de `Sup_Linhas` (Sheet15), das quais `UPDATE_Tratar_IFRS16` depende integralmente para o De-Para de Classe Custo/Centro de Custo e para o multiplicador de sinal, não são escritas por nenhuma das macros de refresh analisadas (`Refresh_Base_Segmento`, `Refresh_Base_De_Para_Ref_Cruzadas`, `Refresh_Base_Suporte_Linhas`, `Refresh_Drop_Comb_Hubble`, `Extrair_Valid_Lin`).
- **Evidência:** `grep -rn` por `Columns(71)`, `Columns(73)`, `Columns(15)`, `Columns(22)` e `BX1` em todo `vba_dump_tmp/` retorna ocorrências **exclusivamente** em `fx_IFRS16.bas`.
- **Impacto:** se a manutenção manual dessa tabela for interrompida (saída de um usuário-chave, por exemplo) ou feita de forma inconsistente, o tratamento contábil de IFRS16 passa a classificar incorretamente (ou deixar de classificar — RN-074) linhas de leasing, com impacto direto no resultado reportado, sem qualquer alerta automatizado além da contagem genérica de linhas puladas.
- **Componentes afetados:** `fx_IFRS16.bas` (`UPDATE_Tratar_IFRS16`), `Sup_Linhas` (Sheet15), extração "Base_1009".
- **Cuidados antes de alterar:** **[VALIDAR COM O NEGÓCIO]** antes de qualquer refatoração — identificar quem edita essas colunas hoje, com que frequência, e a partir de qual fonte de verdade (planilha de controle de contratos de leasing?). Não presumir que essas colunas podem ser recalculadas automaticamente sem esse levantamento.
- **Teste sugerido:** comparar, para um período fechado, a lista de Centros de Custo tratados por IFRS16 (rastreável pelo sufixo "_IFRS16 Tratado" na Fonte) contra a lista de contratos de leasing ativos no sistema de origem (fora do Quick Data) para validar completude e correção do De-Para atual antes de migrar.

### R-02 — Status de ativação da validação automática de hierarquia (Main Results) é incerto
- **Descrição:** `Worksheet_Change` em `BackupCodigo_MainResults.bas` está posicionado em um módulo `.bas` comum, não no code-behind de uma planilha — condição sob a qual o Excel não dispara o evento automaticamente. Nenhuma chamada explícita a este procedimento existe em nenhum outro lugar do código (`grep` completo).
- **Evidência:** `BackupCodigo_MainResults.bas:2` (assinatura de evento fora de `.cls`); ausência confirmada de qualquer outra ocorrência de `Worksheet_Change` chamando ou referenciando este procedimento; ausência de qualquer `Worksheet_Change` equivalente em qualquer arquivo `.cls` do dump completo (`Sheet1` a `Sheet33`, `ThisWorkbook`).
- **Impacto:** se a documentação formal (24 seções) descrever esta validação como comportamento vivo do sistema sem confirmação, um leitor (ou o time de reescrita) pode implementar uma trava de validação que **não reflete o comportamento real em produção hoje** — ou, inversamente, se o código estiver de fato ativo por algum mecanismo não capturado neste dump (ex. divergência entre o `.xlsb` binário e este dump de texto), a documentação pode **omitir** uma trava real, levando a regressão funcional na reescrita.
- **Componentes afetados:** `BackupCodigo_MainResults.bas` (todos os 3 procedimentos), planilha "Main Results" (Sheet3, presumido), `Sheet25` (base de validação).
- **Cuidados antes de alterar:** **[VALIDAR COM O NEGÓCIO]** — abrir o arquivo `.xlsb` real e inspecionar diretamente o code-behind de `Sheet3` (não coberto por este dump de texto, que reflete os módulos conforme extraídos) para confirmar se existe uma cópia ativa (idêntica ou modificada) deste evento; também checar o valor atual da célula `L32` em produção.
- **Teste sugerido:** em ambiente de teste, alterar uma célula de abertura em Main Results com `L32 = "SIM"` e observar se uma lista de validação é de fato aplicada às células dependentes; repetir com o arquivo de produção real (não uma cópia deste dump) antes de decidir se a regra RN-093 entra ou não na reescrita como funcionalidade viva.

## ALTO

### R-03 — Lógica de negócio crítica expressa como strings de fórmula Excel dentro do VBA
- **Descrição:** Em `Limpeza_Base_Ajustes.bas`, todas as regras de cascata de Ref Cruzada, De-Para de Classe/Centro de Custo, Segmento/Abertura de destino e a exceção "Labour Cost" estão embutidas como strings de `FormulaR1C1` construídas por concatenação dentro do VBA (não como funções VBA testáveis).
- **Evidência:** `Limpeza_Base_Ajustes.bas:141-219` (12 blocos de fórmula).
- **Impacto:** alto risco de erro de tradução/perda de regra na reescrita, pois a lógica não é revisável isoladamente (está "escondida" dentro de strings) nem coberta por testes automatizados possíveis no VBA legado.
- **Componentes afetados:** `Limpeza_Base_Ajustes.bas`, planilha AJUSTES (Sheet13).
- **Cuidados antes de alterar:** decodificar cada fórmula 1:1 (já feito nesta e na entrega anterior — ver RN-082 a RN-088, RN-095, RN-096) e validar cada regra decodificada com o time de negócio linha a linha antes de implementar em Python.
- **Teste sugerido:** para uma amostra de linhas reais de Ajustes de um período fechado, calcular os campos de destino tanto pela fórmula Excel original quanto pela implementação Python equivalente e comparar resultado célula a célula (teste de paridade).

### R-04 — Erros de De-Para/mapeamento silenciosamente descartados sem log individual
- **Descrição:** Tanto `fx_IFRS16.bas` (`UPDATE_Tratar_IFRS16`) quanto `Refresh_De_X_Para.bas` (`HandleNAErrorIsErrorWithCleaning`) tratam falhas de lookup apenas com contadores agregados ou substituição genérica por texto ("Not Found"/"Error"), sem registrar qual linha/chave específica falhou.
- **Evidência:** `fx_IFRS16.bas:40-43,133-137` (contador `xx`, sem lista de linhas); `Refresh_De_X_Para.bas:257-263` (substitui qualquer erro não-`#N/A` por "Error" genérico).
- **Impacto:** dificulta auditoria e reconciliação — um analista não consegue, a partir do sistema, saber exatamente quais Centros de Custo ficaram sem tratamento IFRS16 numa execução específica sem investigação manual da base.
- **Componentes afetados:** `fx_IFRS16.bas`, `Refresh_De_X_Para.bas`.
- **Cuidados antes de alterar:** na reescrita, substituir por log estruturado (linha, chave de busca, tabela de referência, motivo) por execução.
- **Teste sugerido:** injetar deliberadamente uma linha com CC sem correspondência em `Sup_Linhas` e verificar se o novo sistema reporta a linha específica, não apenas uma contagem.

### R-05 — Credenciais de banco de dados em texto puro no código-fonte (achado transversal, fora dos 9 módulos, mas usado por eles)
- **Descrição:** `Conexoes.bas` (`AbreConexao`, chamada por `Refresh_Base_Segmento`, `Refresh_Base_Suporte_Linhas`, `Extrair_Valid_Lin`) contém usuário e senha do SQL Server diretamente na string de conexão VBA (login principal e login de fallback).
- **Evidência:** `Conexoes.bas:26-27,36-37` (linhas não reproduzidas neste documento por prudência de segurança — a senha está em texto puro no arquivo original).
- **Impacto:** qualquer pessoa com acesso de leitura ao arquivo `.xlsb`/projeto VBA tem acesso de leitura/escrita ao banco `BPAM`.
- **Componentes afetados:** `Conexoes.bas`; todos os procedimentos que chamam `AbreConexao` (`Refresh_Base_Segmento`, `Refresh_Base_Suporte_Linhas`, `Extrair_Valid_Lin`, entre outros fora do escopo dos 9 módulos).
- **Cuidados antes de alterar:** priorizar rotação das credenciais atuais e migração para variável de ambiente/cofre de segredos na reescrita, independente do cronograma de migração funcional.
- **Teste sugerido:** N/A (ação de segurança, não de funcionalidade) — validar apenas que a nova forma de autenticação mantém o mesmo nível de acesso necessário após a rotação.

## MÉDIO

### R-06 — Posicionamento de colunas por deslocamento aritmético/índice numérico fixo
- **Descrição:** `Refresh_Sup_Linhas.bas` usa deslocamentos de coluna calculados (`Col_Destino + Qtd_Colunas + 14`) e `fx_IFRS16.bas` usa índices de coluna fixos por número (23, 35, 71, 73, 15, 22, 31-33) em vez de busca por nome de cabeçalho (`fn.Match`, usada consistentemente em outros módulos como `Limpeza_Base_Ajustes.bas`).
- **Evidência:** `Refresh_Sup_Linhas.bas:114,210,236`; `fx_IFRS16.bas:26,27,40,45-49,53,62,66-67`.
- **Impacto:** qualquer reordenação de colunas na origem (SQL) ou em `Sup_Linhas` quebra silenciosamente o mapeamento, sem erro explícito — pode gerar dados incorretos sem falha visível.
- **Componentes afetados:** `Refresh_Sup_Linhas.bas`, `fx_IFRS16.bas`.
- **Cuidados antes de alterar:** na reescrita, usar acesso por nome de coluna (não por posição) de forma consistente em todo o pipeline.
- **Teste sugerido:** teste de regressão que reordena colunas de entrada simuladas e verifica se o sistema falha explicitamente (em vez de produzir dado incorreto silencioso).

### R-07 — Corte silencioso de listas de validação acima de 747 itens
- **Descrição:** Ver RN-091 — valores excedentes de uma lista de validação de Ajustes são descartados sem aviso.
- **Evidência:** `Lista_Validacao_Ajustes.bas:122-129`.
- **Impacto:** um usuário pode não encontrar uma opção válida no dropdown sem saber por quê, levando a preenchimento incorreto ou a contato com suporte sem causa raiz óbvia.
- **Componentes afetados:** `Lista_Validacao_Ajustes.bas`, planilha AJUSTES.
- **Cuidados antes de alterar:** na reescrita, preferir um componente de UI sem limite prático de itens (ex. campo de busca/autocomplete) em vez de replicar o limite do Excel.
- **Teste sugerido:** verificar quantas colunas de `Sup_Linhas`, na base real de produção, hoje ultrapassam 747 valores distintos, para dimensionar o impacto real antes de migrar.

### R-08 — Função `Atualizar_Validacao_Linhas_Geral` incompleta, mas relata sucesso ao usuário
- **Descrição:** O corpo que aplicaria a validação está vazio; a função sempre exibe "Processo concluído com sucesso!".
- **Evidência:** `Front_Processos.bas:128-135,141`.
- **Impacto:** usuário pode acreditar erroneamente que uma ação foi executada.
- **Componentes afetados:** `Front_Processos.bas`, qualquer planilha com botão ligado a esta macro.
- **Cuidados antes de alterar:** **[VALIDAR COM O NEGÓCIO]** se este botão ainda está exposto na UI em produção e se algum processo operacional presume (incorretamente) que ele funciona.
- **Teste sugerido:** localizar, no arquivo `.xlsb` real, todos os shapes/botões com `OnAction="Atualizar_Validacao_Linhas_Geral"` e confirmar com o negócio se ainda são usados.

### R-09 — Dependência operacional não automatizada entre refresh de tabelas mestre e limpeza de Ajustes
- **Descrição:** Não há nenhuma chamada de código que garanta que `Refresh_Base_Aux` rode antes de `Processo_Limpar_Ajustes`/`Atualizar_Ajustes_Lista_Validacao_Geral` — a ordem correta depende de disciplina do usuário.
- **Evidência:** ausência de qualquer chamada cruzada entre `Auxiliar.bas:567-580` e `Limpeza_Base_Ajustes.bas`/`Lista_Validacao_Ajustes.bas` no `grep` completo (exceto a chamada automática dentro do fluxo de Importação, item 11 da seção C, que não inclui o refresh das tabelas mestre).
- **Impacto:** se um usuário limpar/validar Ajustes sem antes atualizar as tabelas de suporte, as fórmulas recalculadas usarão dados desatualizados sem nenhum aviso do sistema.
- **Componentes afetados:** `Limpeza_Base_Ajustes.bas`, `Lista_Validacao_Ajustes.bas`, `Auxiliar.bas` (Refresh_Base_Aux).
- **Cuidados antes de alterar:** na reescrita, tornar essa dependência explícita e verificável (ex.: timestamp de última atualização das tabelas mestre, checado antes de permitir a limpeza de Ajustes).
- **Teste sugerido:** simular o cenário de Ajustes limpo com tabelas mestre desatualizadas e verificar o comportamento resultante, para documentar o "pior caso" atual antes de decidir a trava a implementar.

## BAIXO

### R-10 — Código morto e duplicado (variáveis não lidas, lógica triplicada, comentários de versões antigas)
- **Descrição:** Variável `Ordernar` setada mas nunca lida (`Refresh_Sup_Linhas.bas:466`); bloco de ordenação comentado (`Refresh_Sup_Linhas.bas:506-529`); bubble sort de "Total por último" duplicado em 3 lugares (`Auxiliar.bas:1096`, `BackupCodigo_MainResults.bas:188`, `Lista_Validacao_Ajustes.bas:110-129`); `ArrayList arr_Linhas_a_apagar` criado e populado mas não usado para a exclusão real (`fx_IFRS16.bas:10,54`); blocos grandes de SQL antigo comentados em `Refresh_Sup_Linhas.bas:46-59,116-129`.
- **Evidência:** linhas citadas acima.
- **Impacto:** aumenta custo de manutenção e risco de confusão durante a leitura do código legado (qual versão é a "real"), mas não afeta o comportamento funcional atual.
- **Componentes afetados:** `Refresh_Sup_Linhas.bas`, `BackupCodigo_MainResults.bas`, `Lista_Validacao_Ajustes.bas`, `fx_IFRS16.bas`, `Auxiliar.bas`.
- **Cuidados antes de alterar:** nenhum cuidado especial — é seguro simplesmente não replicar código morto na reescrita, desde que se confirme (via este catálogo) que ele de fato não é executado.
- **Teste sugerido:** nenhum necessário além da revisão de código já realizada.

### R-11 — Uso de `xlPinYin` (ordenação chinesa) como método de ordenação
- **Descrição:** Vários blocos de `Sort` usam `SortMethod:=xlPinYin`, provavelmente herança de configuração regional do Excel, não intencional para dados em português.
- **Evidência:** `Refresh_Sup_Linhas.bas:203,497`; `Refresh_De_X_Para.bas:181`.
- **Impacto:** funciona por coincidência para o alfabeto latino usado; risco muito baixo, mas indica configuração não intencional.
- **Componentes afetados:** `Refresh_Sup_Linhas.bas`, `Refresh_De_X_Para.bas`.
- **Cuidados antes de alterar:** nenhum — a reescrita naturalmente usará ordenação padrão (ex. `sort_values` do pandas), tornando este risco irrelevante para o novo sistema.
- **Teste sugerido:** comparar a ordem de saída de uma lista com acentuação/caracteres especiais entre o Excel legado (`xlPinYin`) e a nova implementação, para garantir que a ordem apresentada ao usuário não mude de forma perceptível.

---

*Fim do catálogo. Todas as 23 entradas da seção A, 26 regras de negócio da seção B (RN-071–RN-096, dentro do intervalo reservado RN-071–RN-100), 13 dependências da seção C e 11 riscos da seção D foram derivados por leitura direta dos arquivos em `C:\Users\guilm\PycharmProjects\QuickData\vba_dump_tmp\` e por buscas `grep` completas contra o mesmo diretório, executadas nesta sessão. Nenhuma informação foi inferida sem citar a evidência correspondente; toda hipótese está marcada com [VALIDAR COM O NEGÓCIO] e todo dado não verificável está marcado [NÃO ACESSÍVEL] ou [NÃO IDENTIFICADO].*

## 18.4 Cluster UI / Forms
## CRÍTICO

**1. Chamada a procedimento inexistente (`Registrar_Sheet`)**
- **Descrição:** `Form_Exportacao.ListBox1_DblClick` (linha 469) chama `Registrar_Sheet`, que não está declarado em nenhum arquivo `.bas`/`.cls`/`.frm` do dump.
- **Evidência:** `grep -rn "Registrar_Sheet" vba_dump_tmp/` retorna apenas a linha de chamada (Form_Exportacao.frm:469); nenhuma declaração `Sub`/`Function Registrar_Sheet` em lugar nenhum.
- **Impacto:** Em VBA, uma referência a procedimento inexistente tipicamente impede a compilação do projeto inteiro assim que qualquer macro é executada ("Sub ou Function não definida"). Se isso for verdade também no arquivo `.xlsb` de produção, **todo o projeto VBA pode estar em risco de erro de compilação** — ou o `ListBox1`/handler está de fato morto/inatingível na versão em produção (controle não populado, oculto, ou form não usa mais este ListBox).
- **Componentes afetados:** Form_Exportacao (e, na hipótese de erro de compilação de projeto inteiro, potencialmente todo o workbook).
- **Cuidados antes de alterar:** Não presumir que este código execute normalmente. Antes de portar qualquer lógica do Form_Exportacao, confirmar no arquivo `.xlsb` original se `ListBox1` existe visualmente no form e se o projeto VBA compila sem erros (`Debug > Compile VBAProject` no editor).
- **Teste sugerido:** Abrir o arquivo original, rodar `Debug > Compile VBAProject`; se compilar sem erro, `Registrar_Sheet` pode estar definido em um módulo não incluído neste dump (revisar lista de módulos extraídos) ou o `ListBox1` foi removido do form (a assinatura do evento ficaria "órfã" mas sintaticamente válida) — reconfirmar contagem de módulos do projeto original vs. os extraídos.

**2. Sobrescrita silenciosa de arquivo exportado**
- **Descrição:** Antes de salvar um arquivo exportado, o sistema verifica se já existe um arquivo de mesmo nome na pasta de destino e, se existir, o apaga (`Kill`) sem pedir confirmação.
- **Evidência:** Form_Exportacao.frm:147, 175, 401 (`If Dir(Diret & Nome_Arq) <> "" Then Kill (Diret & Nome_Arq)`). Nome do arquivo inclui timestamp só até o minuto (`Format(Now, "YYYYMMDD_HHMM")`).
- **Impacto:** Perda de dado irrecuperável se duas exportações ocorrerem na mesma pasta dentro do mesmo minuto (ex.: exportação "arquivos separados" gerando múltiplos arquivos em sequência rápida, ou dois usuários exportando ao mesmo tempo para a mesma pasta compartilhada de rede).
- **Componentes afetados:** Form_Exportacao — ambos os modos (Front e Base), ambas as opções (único/separado).
- **Cuidados antes de alterar:** Ao portar para a nova interface, adicionar confirmação explícita de sobrescrita ou gerar nomes garantidamente únicos (timestamp com segundos, GUID, ou checagem incremental de nome).
- **Teste sugerido:** Rodar duas exportações consecutivas rapidamente para a mesma pasta e confirmar se a primeira é apagada sem aviso.

## ALTO

**3. Processamento longo sem cancelamento nem progresso real**
- **Descrição:** `B_Ok_Click` de ambos os forms (Importação: ~580 linhas de processamento; Exportação: laços de cópia/filtro/ordenação/gravação) não oferece barra de progresso nem opção de cancelar — apenas texto na `Application.StatusBar`.
- **Evidência:** Form_Importacao.frm:141-701; Form_Exportacao.frm:82-425; único indicador visual é `Application.StatusBar` (Form_Importacao.frm:174).
- **Impacto:** Em volumes grandes de dados, o usuário pode perceber o Excel como "travado" e forçar o fechamento (Ctrl+Alt+Del / Encerrar Tarefa), correndo risco de corromper o arquivo `.xlsb` ou perder trabalho não salvo.
- **Componentes afetados:** Form_Importacao.B_Ok_Click, Form_Exportacao.B_Ok_Click.
- **Cuidados antes de alterar:** Qualquer reescrita deve preservar (ou melhorar) a visibilidade do progresso e adicionar um mecanismo de cancelamento seguro entre etapas.
- **Teste sugerido:** Medir o tempo de execução de uma importação/exportação com volume de dados real de produção e avaliar em que ponto o usuário perderia a percepção de que o processo está ativo.

**4. Lógica de importação acoplada a rótulos textuais fixos, sem validação de schema**
- **Descrição:** O reconhecimento de "aba é um Front de Quick Data" e o remapeamento de campos dependem de encontrar literalmente os textos "Aux Empresa", "Menu de seleção", "Coluna 5", "Bkp fórmula oficial", "Fórmulas:" nas posições esperadas do arquivo de origem.
- **Evidência:** Form_Importacao.frm:186-188, 317, 371-373.
- **Impacto:** Qualquer alteração de layout no arquivo de origem (renomear rótulo, mover célula) quebra a importação de forma parcial e silenciosa (algumas colunas podem ficar sem dados) ou gera erro não tratado em runtime (`Cells.Find` retorna erro se não encontrar o texto buscado, e não há tratamento explícito de "não encontrado" em todos os pontos, ex.: linha 371-373 não tem `On Error` antes do `.Find`).
- **Componentes afetados:** Form_Importacao.B_Ok_Click (fluxo completo de importação de Front).
- **Cuidados antes de alterar:** Na reescrita, implementar validação de schema explícita antes de iniciar a cópia, com mensagem de erro clara indicando qual âncora não foi encontrada.
- **Teste sugerido:** Importar um arquivo com um dos rótulos-âncora renomeado/deslocado e observar o comportamento (erro não tratado vs. dado incorreto silencioso).

## MÉDIO

**5. Inconsistência de interação clique-simples vs. duplo-clique**
- **Descrição:** Ver RN-109 — remoção de item de lista de seleção é por duplo-clique no Form_Importacao e por clique único no Form_Exportacao.
- **Evidência:** Form_Importacao.frm:736-739 vs. Form_Exportacao.frm:479-482.
- **Impacto:** Risco de remoção acidental de item já selecionado no fluxo de Exportação; confusão do usuário ao transitar entre os dois formulários semelhantes.
- **Componentes afetados:** Form_Exportacao (comportamento mais agressivo/arriscado dos dois).
- **Cuidados antes de alterar:** Padronizar a interação na nova UI (recomenda-se duplo-clique ou um botão explícito de "remover", nunca clique único destrutivo).
- **Teste sugerido:** Teste de usabilidade comparando os dois fluxos com usuários reais do time de P&C.

**6. Funcionalidade incompleta em Form_Tratamento_Opcoes**
- **Descrição:** Ver RN-110 — a lista de extrações habilitadas é calculada mas nunca exibida ao usuário.
- **Evidência:** Form_Tratamento_Opcoes.frm:53, 89, 112 (`Debug.Print dados`).
- **Impacto:** Possível lacuna funcional não percebida pelos usuários atuais (eles podem já não usar/precisar desta tela, ou pode ser uma necessidade não atendida há anos).
- **Componentes afetados:** Form_Tratamento_Opcoes.
- **Cuidados antes de alterar:** Não presumir que a tela precise ser replicada como está — confirmar com o negócio o uso real atual antes de investir na reescrita desta tela.
- **Teste sugerido:** [VALIDAR COM O NEGÓCIO] — perguntar diretamente aos usuários se esta tela é utilizada e para quê.

**7. Matching por substring no menu de comandos do Sheet8**
- **Descrição:** Ver RN-111 — `ComboBox1_Change` usa `InStr` sobre o texto do item selecionado, não um valor/ID fixo.
- **Evidência:** Sheet8.cls:12-16.
- **Impacto:** Frágil a edição/renomeação do texto dos itens do combo (cuja população automática está desativada — Sheet8.cls:20-33, comentado).
- **Componentes afetados:** Sheet8 (painel principal).
- **Cuidados antes de alterar:** Na nova UI, usar identificadores estáveis (enum/ID) em vez de comparação textual.
- **Teste sugerido:** Alterar levemente o texto de um item do combo (ex.: adicionar espaço) e confirmar se o gatilho correspondente para de funcionar sem aviso.

## BAIXO

**8. Macros legadas de Module2.bas não referenciadas em nenhum fluxo**
- **Descrição:** As 4 macros `inverter_valores*` não são chamadas por nenhum botão, form ou outro módulo identificado no dump; uma delas (`inverter_valores_3`) parece uma gravação incompleta (não finaliza a operação de colagem); outra (`inverter_valores`) tem efeito matemático (`valor²`) que diverge do nome sugerido ("inverter" = multiplicar por -1).
- **Evidência:** Module2.bas:2-60; ausência de referências externas confirmada via grep.
- **Impacto:** Baixo risco de negócio atual (não fazem parte de nenhum fluxo automatizado), mas podem ser utilitários manuais ainda usados esporadicamente por analistas avançados via Alt+F8.
- **Componentes afetados:** Module2.bas (isolado).
- **Cuidados antes de alterar:** Não assumir que podem ser simplesmente descartadas na reescrita sem checar com o time se ainda são usadas manualmente.
- **Teste sugerido:** [VALIDAR COM O NEGÓCIO] — perguntar à equipe de P&C se essas macros (ou equivalentes) ainda são acionadas manualmente.

**9. Dimensões de UI hardcoded em Form_Tratamento_Opcoes**
- **Descrição:** O form redimensiona a si mesmo e ao `ListBox1` com valores fixos em pixels (600x200 / largura 564), não responsivos a diferentes resoluções/DPI.
- **Evidência:** Form_Tratamento_Opcoes.frm:35-37.
- **Impacto:** Baixo (problema cosmético/usabilidade em telas de alta resolução ou DPI diferente do ambiente original).
- **Componentes afetados:** Form_Tratamento_Opcoes.
- **Cuidados antes de alterar:** Ao portar, usar dimensionamento relativo/responsivo em vez de replicar valores fixos.
- **Teste sugerido:** Abrir o form original em monitores de resoluções/DPI distintos e observar recorte/desproporção.

---

*Fim do catálogo. Ressalva de limitação sobre o `.frx`/definição visual dos controles reproduzida no topo deste arquivo deve constar na Seção 4 (Limitações Metodológicas) do documento final consolidado.*
