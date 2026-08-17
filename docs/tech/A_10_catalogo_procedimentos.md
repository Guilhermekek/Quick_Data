# 10. Catálogo de Procedimentos VBA

> Os 192 procedimentos do sistema, documentados em 4 frentes por cluster de módulos (mesmo formato de 15 campos em todas). Ver seção 24.2 para o índice quantitativo.

## 10.1 Cluster Extração / ETL (54 procedimentos — 12 módulos)
# A) CATÁLOGO POR PROCEDIMENTO

## Módulo `Conexoes.bas`

### 1. AbreConexao

1. **Nome completo:** `AbreConexao`
2. **Módulo:** `Conexoes.bas`
3. **Tipo:** Function (retorno não usado — ver nota transversal)
4. **Escopo:** `Public` — `Conexoes.bas:9`
5. **Objetivo (negócio):** Abrir a conexão com o SQL Server corporativo (banco `BPAM`) que alimenta a extração
   automatizada do Hubble e diversas rotinas de refresh — é o ponto único de acesso a dados via SQL de todo o
   sistema.
6. **Quem chama (evidência):** Dentro do cluster: `Extracao_SQL_Hubble.bas:367,419,450,522`
   (`Criar_TB_SQL_AUX`, `Excluir_TB_SQL_AUX`, `Extrair_Linha`, `Extrair_Base_Final`). Fora do cluster:
   `TK_Functions.bas:587`, `Auxiliar.bas:85`, `Refresh_DP_Segmento.bas:18`, `Refresh_Sup_Linhas.bas:18,326,381`.
7. **Procedimentos chamados:** nenhum da lista indexada (usa apenas `ADODB.Connection` nativo e `MsgBox`).
8. **Parâmetros:** nenhum.
9. **Retorno:** Function sem valor atribuído ao próprio nome (ver nota transversal) — na prática usada como Sub.
10. **Variáveis/objetos relevantes:** `conn` (Global `ADODB.Connection`, declarada `Conexoes.bas:6`),
    `tProv`="Sqloledb", `tBase`="SNEPDB24V" (Data Source), catálogo primário "BPAM", usuário "AdminBPAM",
    senha **[INFORMAÇÃO SENSÍVEL OMITIDA]**; fallback catálogo "InfoGER", usuário "URELATIG"/"Report2BPAM",
    senha **[INFORMAÇÃO SENSÍVEL OMITIDA]**.
11. **Abas/intervalos/arquivos acessados:** nenhum — acessa exclusivamente o servidor SQL Server via rede.
12. **Pré-condições:** biblioteca de referência "Microsoft ActiveX Data Objects 2.8" habilitada
    (`Conexoes.bas:11-12`, comentário); rede/VPN corporativa com acesso a `SNEPDB24V`.
13. **Passos principais:**
    - Monta `conString` com provider/servidor/catálogo/usuário/senha (linha 26).
    - `conn.Open` com a connection string primária (linha 27).
    - Se falhar (rota de erro), monta connection string alternativa (catálogo InfoGER) e reabre (linhas 36-37).
    - Se ambas falharem, `MsgBox` de erro (linha 43).
14. **Pós-condições:** variável global `conn` aberta e pronta para uso por `ADODB.Command`/`Recordset`
    subsequentes (ou `MsgBox` de erro se ambas as tentativas falharem).
15. **Efeitos colaterais/tratamento de erro/mensagens/regra de negócio/riscos/evidência:**
    - **CRÍTICO (segurança):** credenciais de banco de dados em texto plano no código-fonte
      (linhas 26-27 e 36-37) — usuário e senha visíveis a qualquer pessoa com acesso ao VBA Project. A senha do
      usuário `AdminBPAM` é a mesma reutilizada na tentativa de fallback com usuário `Report2BPAM`/`URELATIG`
      (mesmo valor de texto em ambos os blocos) — indício de padrão de senha institucional reaproveitado.
    - `On Error GoTo tratar_erro` está **comentado** (linha 19: `'    On Error GoTo tratar_erro`) — o bloco de
      fallback (linhas 32 em diante) só é alcançado se algum mecanismo externo de tratamento de erro do projeto
      ainda assim direcionar a execução para lá; como a linha está desativada, não há confirmação de que o
      fallback realmente dispara em uma falha real de conexão — **[VALIDAR COM O NEGÓCIO / TESTE TÉCNICO]**.
    - Mensagem ao usuário (literal, linha 43): "Problema de conexão. Verifique sua conexão de rede e reinicie o
      arquivo."
    - Evidência: linhas 19, 26-27, 32, 36-37, 43.

### 2. FechaConexao

1. **Nome completo:** `FechaConexao`
2. **Módulo:** `Conexoes.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Conexoes.bas:47`
5. **Objetivo (negócio):** Encerrar a conexão SQL aberta por `AbreConexao`, liberando o recurso de rede/banco
   ao final de cada operação.
6. **Quem chama (evidência):** Dentro do cluster: `Extracao_SQL_Hubble.bas:412,444,515,567` (labels `FimdaMacro`
   de `Criar_TB_SQL_AUX`, `Excluir_TB_SQL_AUX`, `Extrair_Linha`, `Extrair_Base_Final`). Fora do cluster:
   `TK_Functions.bas:641`, `Auxiliar.bas:99,114`, `Refresh_DP_Segmento.bas:83`, `Refresh_Sup_Linhas.bas:307,370`.
7. **Procedimentos chamados:** nenhum da lista indexada.
8. **Parâmetros:** nenhum.
9. **Retorno:** Function sem valor atribuído (ver nota transversal).
10. **Variáveis/objetos relevantes:** `conn` (fechada e setada `Nothing`).
11. **Abas/intervalos/arquivos:** nenhum.
12. **Pré-condições:** `conn` deve ter sido aberta previamente (senão o erro ao fechar é capturado por
    `On Error GoTo ErroConexao`, linha 51).
13. **Passos principais:** `conn.Close` (linha 52); `Set conn = Nothing` (linha 53); em caso de erro, `MsgBox`.
14. **Pós-condições:** `conn` liberado; nenhuma conexão SQL permanece aberta.
15. **Efeitos colaterais/erro/mensagens/riscos:** mensagem ao usuário (literal, linha 57): "Problemas para fechar
    conexão com SQL. Verifique sua rede." Se `AbreConexao` tiver falhado antes (deixando `conn` em estado
    inconsistente), esta função pode mascarar a causa raiz com uma mensagem genérica de fechamento. Evidência:
    linhas 51-58.

---

## Módulo `Aux_Leitura_Nome_Arqs.bas`

### 3. Extrair_Diretorio_Arq_Base_Consolidada

1. **Nome completo:** `Extrair_Diretorio_Arq_Base_Consolidada`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:3`
5. **Objetivo (negócio):** Acionar a seleção manual (via caixa de diálogo do Windows) do arquivo-fonte "Base
   Consolidada" pelo usuário, gravando o caminho escolhido na aba de configuração de extração.
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Extrair_Diretorio_Arq_Base_Consolidada` encontrada
   no `grep` do diretório — não referenciado no código lido — provável botão/shape na aba "Extração" (Sheet8),
   cujo `OnAction` **[NÃO ACESSÍVEL]** a partir do texto VBA extraído.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 6.
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `Chave` = "Base Consolidada" (literal, usado como texto-âncora de busca
    por `Cells.Find` dentro de `GetArquivo`).
11. **Abas/intervalos/arquivos acessados:** `Sheet8` (Extração), indiretamente via `GetArquivo`.
12. **Pré-condições:** a aba Extração deve conter uma célula com o texto exato "Base Consolidada".
13. **Passos principais:** define `Chave`; chama `GetArquivo(Chave)`.
14. **Pós-condições:** caminho/nome do arquivo gravados nas células correspondentes de `Sheet8` (efeito de
    `GetArquivo`).
15. **Efeitos colaterais/riscos:** nenhum tratamento de erro próprio; depende inteiramente do comportamento de
    `GetArquivo`. Evidência: linhas 3-8.

### 4. Extrair_Diretorio_Fixed_Revenues

1. **Nome completo:** `Extrair_Diretorio_Fixed_Revenues`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:10`
5. **Objetivo (negócio):** Mesmo padrão do item 3, para o arquivo-fonte "Fixed Revenues".
6. **Quem chama (evidência):** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 13.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis relevantes:** `Chave` = "FIXED REVENUES".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula com texto "FIXED REVENUES" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetArquivo`.
14. **Pós-condições:** caminho/arquivo gravados em `Sheet8`.
15. **Riscos:** idem item 3. Evidência: linhas 10-15.

### 5. Extrair_Diretorio_Quick_Data

1. **Nome completo:** `Extrair_Diretorio_Quick_Data`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:17`
5. **Objetivo (negócio):** Mesmo padrão do item 3, para o arquivo-fonte "Quick Data".
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 20.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Chave` = "QUICK DATA".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula "QUICK DATA" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetArquivo`.
14. **Pós-condições:** caminho/arquivo gravados em `Sheet8`.
15. **Riscos:** idem item 3. Evidência: linhas 17-22.

### 6. Extrair_Diretorio_Arq_1009

1. **Nome completo:** `Extrair_Diretorio_Arq_1009`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:24`
5. **Objetivo (negócio):** Mesmo padrão do item 3, para o arquivo-fonte "1009".
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 27.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Chave` = "1009".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula "1009" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetArquivo`.
14. **Pós-condições:** caminho/arquivo gravados em `Sheet8`.
15. **Riscos:** texto-âncora curto ("1009") tem maior chance de casar com múltiplas células por engano
    (`Cells.Find` retorna a primeira ocorrência). Evidência: linhas 24-29.

### 7. Extrair_Diretorio_RGM

1. **Nome completo:** `Extrair_Diretorio_RGM`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:31`
5. **Objetivo (negócio):** Mesmo padrão do item 3, para o arquivo-fonte "(RGM)".
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 34.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Chave` = "(RGM)".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula "(RGM)" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetArquivo`.
14. **Pós-condições:** caminho/arquivo gravados em `Sheet8`.
15. **Riscos:** idem item 3. Evidência: linhas 31-36.

### 8. Extrair_Diretorio_Arq_Other_Ico

1. **Nome completo:** `Extrair_Diretorio_Arq_Other_Ico`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:38`
5. **Objetivo (negócio):** Acionar a seleção manual de uma **pasta** (não um arquivo único) contendo os
   múltiplos arquivos de "Other Income", já que essa fonte processa vários arquivos por operadora.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetPasta(Chave)` — linha 41 (diferente dos demais, que chamam `GetArquivo`).
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Chave` = "Other income".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula "Other income" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetPasta`.
14. **Pós-condições:** caminho da pasta gravado em `Sheet8`.
15. **Regra de negócio:** confirma que "Other Income" é a única fonte do cluster que opera sobre uma pasta
    inteira de arquivos, não um arquivo único (ver RN-030). Evidência: linhas 38-43.

### 9. Extrair_Diretorio_Arq_MockUP

1. **Nome completo:** `Extrair_Diretorio_Arq_MockUP`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:45`
5. **Objetivo (negócio):** Mesmo padrão do item 3, para o arquivo-fonte "MOCKUP".
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `GetArquivo(Chave)` — linha 48.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Chave` = "MOCKUP".
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula "MOCKUP" deve existir na aba Extração.
13. **Passos:** define `Chave`; chama `GetArquivo`.
14. **Pós-condições:** caminho/arquivo gravados em `Sheet8`.
15. **Riscos:** idem item 3. Evidência: linhas 45-50.

### 10. GetArquivo

1. **Nome completo:** `GetArquivo`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:52`
5. **Objetivo (negócio):** Rotina genérica que localiza a célula-âncora na aba Extração pelo texto passado em
   `Chave`, abre o diálogo nativo do Windows para o usuário escolher um arquivo, e grava separadamente diretório
   e nome do arquivo nas duas células adjacentes — é o mecanismo comum de apontamento manual de fonte de dados
   usado por 6 das 7 fontes externas do cluster.
6. **Quem chama (evidência):** Dentro do cluster: `Extrair_Diretorio_Arq_Base_Consolidada` (linha 6),
   `Extrair_Diretorio_Fixed_Revenues` (13), `Extrair_Diretorio_Quick_Data` (20), `Extrair_Diretorio_Arq_1009`
   (27), `Extrair_Diretorio_RGM` (34), `Extrair_Diretorio_Arq_MockUP` (48) — 6 chamadas. Fora do cluster:
   `Form_Importacao.frm:710` (`Call GetArquivo(Diret_e_Arq)`).
7. **Procedimentos chamados:** nenhum da lista indexada (usa `Application.GetOpenFilename` nativo).
8. **Parâmetros:** `Chave` (Variant implícito, passagem `ByRef` padrão VBA) — texto usado como critério de busca
   de célula-âncora; se igual a `"*"`, é reaproveitado para armazenar o caminho completo (branch nunca acionado
   pelos 6 chamadores atuais — ver risco).
9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `sArquivo` (caminho completo escolhido), `Cell_Diret_Arq` (endereço da
    célula-âncora +1 linha), `Range_Diret`/`Range_Arq` (endereços onde gravar diretório/nome do arquivo),
    `sEspecificação` = "All Files (*.*),*.*".
11. **Abas/intervalos/arquivos acessados:** `Sheet8` ("Extração"); diálogo do sistema operacional (não é
    arquivo do domínio de dados).
12. **Pré-condições:** célula com o texto de `Chave` deve existir na aba Extração — senão `Cells.Find` retorna
    `Nothing` e a linha seguinte (`.Offset(1,0).Address`) gera erro de execução não tratado.
13. **Passos principais:**
    - Localiza célula-âncora via `Cells.Find(What:=Chave, ...)` (linha 60).
    - Abre `Application.GetOpenFilename` sem filtro de extensão (linha 69).
    - Se cancelado, `MsgBox` crítico "Nenhum arquivo foi selecionado!" (linha 97).
    - Senão, separa diretório/nome pelo último "\" (linhas 79-85).
    - Grava em `Range_Diret`/`Range_Arq` e recalcula a linha (linhas 87-91).
14. **Pós-condições:** células de caminho/nome de arquivo atualizadas na aba Extração, usadas pelos módulos de
    extração correspondentes na próxima execução.
15. **Efeitos colaterais/erro/mensagens/riscos:** filtro "All Files (*.*)" não restringe o tipo de arquivo — o
    usuário pode selecionar um arquivo de formato incorreto sem aviso nesta etapa (o erro só aparece depois, no
    módulo de extração real, com mensagem pouco clara). `Cells.Find` sem checagem de `Is Nothing`. Branch morto
    (`If Chave = "*"`, linha 73) nunca é acionado pelos 6 chamadores atuais — código vestigial. Evidência:
    linhas 60-69, 73, 97.

### 11. GetPasta

1. **Nome completo:** `GetPasta`
2. **Módulo:** `Aux_Leitura_Nome_Arqs.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Aux_Leitura_Nome_Arqs.bas:103`
5. **Objetivo (negócio):** Variante de `GetArquivo` que pede ao usuário selecionar uma **pasta**, usada
   exclusivamente pela extração de "Other Income" (múltiplos arquivos de operadoras).
6. **Quem chama (evidência):** Dentro do cluster: `Extrair_Diretorio_Arq_Other_Ico` (linha 41). Fora do cluster:
   `Form_Exportacao.frm:91` (`Call GetPasta(Diret)`).
7. **Procedimentos chamados:** nenhum da lista indexada (`Application.FileDialog(msoFileDialogFolderPicker)`).
8. **Parâmetros:** `Chave` (Variant implícito, `ByRef`).
9. **Retorno:** N/A.
10. **Variáveis relevantes:** `MyFolder` (`FileDialog`), `Range_Diret`.
11. **Abas/arquivos:** `Sheet8`.
12. **Pré-condições:** célula-âncora com texto `Chave` deve existir.
13. **Passos principais:** localiza célula-âncora (linha 106); abre `FileDialog` de pasta (linha 112); se
    cancelado, `Exit Sub` **silencioso** (linhas 116-118, sem `MsgBox` — diferente de `GetArquivo`); senão grava
    a pasta selecionada (linha 125).
14. **Pós-condições:** célula de diretório atualizada em `Sheet8`.
15. **Riscos:** inconsistência de UX com `GetArquivo` — aqui o cancelamento é silencioso, lá é avisado por
    `MsgBox`; mesma ausência de checagem de `Is Nothing` no `Find`. Evidência: linhas 106-118, 120-129.

---

## Módulo `Extracao_SQL_Hubble.bas`

### 12. Extrair_Base_Hubble

1. **Nome completo:** `Extrair_Base_Hubble`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_SQL_Hubble.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para atualizar a Base com dados do Hubble — orquestra
   limpeza da fonte, extração via SQL dinâmico e pós-processamento (listas de KPI/Versão, cronômetro de
   duração).
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Extrair_Base_Hubble` encontrada — não referenciado
   no código lido — provável botão/shape na aba Extração. **[NÃO ACESSÍVEL]** o `OnAction` real.
7. **Procedimentos chamados:** `Verifica_Versao` (`Auxiliar.bas`), `Desligar_Tudo` (`Auxiliar.bas`),
   `Processo_Limpar_Base_Hubble` (linha 11), `Processo_Extrair_Base_Hubble` (linha 12),
   `Atualizar_Lista_KPI_Versao_Interna` (`Refresh_Sup_Linhas.bas`), `PopUp_Tempo_Processamento`
   (`Auxiliar.bas`), `Ativar_Tudo` (`Auxiliar.bas`). Linhas 14 e 16 (`Calcular_Comb_Meses_Intervalo`,
   `UPDATE_Tratar_IFRS16`) estão **comentadas** — não executadas.
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `Inicio` (Global, `TimeValue(Now)`, usada por `PopUp_Tempo_Processamento` para
    medir duração — linha 5).
11. **Abas/intervalos/arquivos acessados:** `Sheet8` (`.Select`, linha 15).
12. **Pré-condições:** aba Extração configurada com ao menos uma linha "Sim" para extração; conexão SQL
    disponível.
13. **Passos principais:** cronometra início; verifica versão; desliga eventos/telas; limpa base Hubble; extrai;
    atualiza listas de KPI/Versão; seleciona `Sheet8`; exibe popup de tempo; reativa a UI.
14. **Pós-condições:** base Hubble atualizada na aba Base; UI reativada.
15. **Efeitos colaterais/erro/regra de negócio/riscos:** linha 16 (`'Call UPDATE_Tratar_IFRS16`) confirma que o
    tratamento de IFRS16 **não é aplicado** à base Hubble, ao contrário da base 1009 (que chama
    `UPDATE_Tratar_IFRS16` ativamente — ver item 18) — decisão de negócio implícita não documentada em
    comentário — **[VALIDAR COM O NEGÓCIO]**. Linha 14 (`Calcular_Comb_Meses_Intervalo`) também desativada —
    código morto. Evidência: linhas 14, 16.

### 13. Limpar_Base_Hubble

1. **Nome completo:** `Limpar_Base_Hubble`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_SQL_Hubble.bas:23`
5. **Objetivo (negócio):** Ponto de entrada (botão) para limpar apenas a base Hubble sem reextrair, com
   confirmação visual ao usuário mostrando os critérios usados.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Processo_Limpar_Base_Hubble` (linha 26), `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído (ver nota transversal).
10. **Variáveis relevantes:** `Campo`, `Chave` (globais de módulo implícitas, Variant).
11. **Abas/arquivos:** nenhuma diretamente (delega a `Processo_Limpar_Base_Hubble`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga tudo; limpa; reativa; exibe `MsgBox` de sucesso mostrando `Campo`/`Chave`.
14. **Pós-condições:** registros da fonte "Base_Hubble" removidos da Base.
15. **Mensagem ao usuário (literal, linhas 29-31):** "Limpeza concluída com sucesso !!!! ... • Campo >>>
    FONTE ... • Chave >>> BASE_HUBBLE". Evidência: linhas 26-31.

### 14. Processo_Limpar_Base_Hubble

1. **Nome completo:** `Processo_Limpar_Base_Hubble`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_SQL_Hubble.bas:35`
5. **Objetivo (negócio):** Definir os critérios (Campo="Fonte", Chave="Base_Hubble") e delegar a exclusão de
   registros históricos da fonte Hubble ao utilitário genérico de limpeza.
6. **Quem chama (evidência):** `Extrair_Base_Hubble` (linha 11), `Limpar_Base_Hubble` (linha 26). Nenhuma
   ocorrência adicional no `grep`.
7. **Procedimentos chamados:** `Limpar_Base_Historica(Campo, Chave)` (`Auxiliar.bas:496`).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Hubble".
11. **Abas/arquivos:** delega a `Sheet3` (Base) via `Limpar_Base_Historica`.
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** seta as duas variáveis; chama `Limpar_Base_Historica`.
14. **Pós-condições:** linhas com Fonte="Base_Hubble" removidas da aba Base.
15. **Riscos:** nenhum tratamento de erro próprio. Evidência: linhas 37-39.

### 15. Processo_Extrair_Base_Hubble

1. **Nome completo:** `Processo_Extrair_Base_Hubble`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Function (retorno não usado) — **a rotina mais extensa e crítica do módulo** (≈320 linhas)
4. **Escopo:** `Public` — `Extracao_SQL_Hubble.bas:44`
5. **Objetivo (negócio):** Para cada linha de configuração marcada "Sim" na aba Extração, construir e executar
   dinamicamente uma query SQL de extração/agregação contra a base Hubble, aplicando os filtros de
   KPI/Versão/Empresa/IFRS/Organic/mês configurados pelo usuário, e gravar o resultado consolidado na aba Base.
6. **Quem chama (evidência):** `Extrair_Base_Hubble` (linha 12, dentro do cluster); `Auxiliar.bas:15`
   (`Extrair_Todas_as_Bases`, fora do cluster — chamada direta, sem passar pelo wrapper `Extrair_Base_Hubble`).
7. **Procedimentos chamados:** `Excluir_TB_SQL_AUX` (linha 66), `Criar_TB_SQL_AUX` (linha 67), `Extrair_Linha`
   (linha 284, dentro do loop `For Lin = Lin_Cabecalho_Capa To Lin_Final`), `Extrair_Base_Final` (linha 293),
   `Excluir_TB_SQL_AUX` (linha 341, segunda chamada — limpeza pós-uso), `Reclassificar_Combinacoes_Empresas` /
   `_IFRS_Contabil` / `_Proforma` / `Calcular_Comb_Meses_Intervalo_Linha` (linhas 356-359, todas em
   `Auxiliar.bas`), `Form_Linha_BD` (linha 346, `Aux_Formulas_Base.bas`).
8. **Parâmetros:** nenhum (opera sobre estado global de planilha).
9. **Retorno:** Function sem valor atribuído.
10. **Variáveis/objetos relevantes:** `Sh8`=`Sheet8` (Extração), `Sh3`=`Sheet3` (Base/destino), `Sh9`=`Sheet9`
    (tabela de chaves De-Para p/ filtros complementares), `Lin_Cabecalho_Capa`/`Lin_Cabecalho_Base` (linhas de
    cabeçalho localizadas via `fn.Match`), `Filtr_Lin`/`Filtr_Col`/`Classif_Extra`/`Classif_Organic`/
    `SomaMensal` (fragmentos de SQL construídos dinamicamente por concatenação de string).
11. **Abas/intervalos/arquivos acessados:** `Sheet8` (config, colunas a partir de "Extrair?"), `Sheet3`
    (destino), `Sheet9` (chaves de filtro Empresa/IFRS_Contábil/Organic), `Sheet24!D19` (nome da tabela SQL de
    origem, lido dentro de `Extrair_Linha`). Tabela SQL de origem: `[BPAM].[dbo].[<nome em Sheet24!D19>]` — não
    hardcoded no VBA, configurado em célula.
12. **Pré-condições:** aba Extração com pelo menos uma linha "Sim"; `Sheet9` com tabela de chaves preenchida;
    conexão SQL acessível.
13. **Passos principais:**
    - Localiza cabeçalhos e limites de linha/coluna nas abas Extração e Base (linhas 73-88).
    - Monta cláusulas de filtro complementares por tipo de coluna: Empresa/IFRS_Contábil (via `Sheet9`, lógica
      OR/AND por "=" e "<>", linhas 100-135), Organic (via `CASE WHEN` mapeando faixa de `REF_Organic`, linhas
      137-148), meses (soma condicional por mês, linhas 150-157).
    - Para cada linha marcada "Sim" (loop linhas 164-289): monta filtro de KPI/Versão (`Split(" > ")`), filtros
      de Empresa/Organic/IFRS/Proforma, filtro de Exercício (`YEAR([MES_REF])`), decide agregação "c/ CdC Fiel"
      vs "s/ CdC" conforme coluna "Extrair*CDC*"; chama `Extrair_Linha`.
    - Após o loop: chama `Extrair_Base_Final` (linha 293).
    - Redefine KPI/Versão na Base via fórmula `INDEX`/`FIND` que separa "KPI_DESTINO"/"KPI_VERSAO" (linhas
      299-336), convertendo fórmula em valor.
    - Remove a tabela auxiliar SQL; recalcula `Form_Linha_BD`; `ActiveWorkbook.RefreshAll`; reclassifica
      combinações (linhas 341-359).
14. **Pós-condições:** linhas da fonte "Base_Hubble" inseridas na aba Base; tabela auxiliar SQL removida do
    servidor.
15. **Efeitos colaterais/erro/regra de negócio/riscos:**
    - **RN relevante:** distinção "c/ CdC Fiel" vs "s/ CdC" controla se o Centro de Custo é agregado
      granularmente ou consolidado em `'-'` (linhas 276-281) — configurável por linha da matriz de extração
      (ver RN-005).
    - **RN relevante:** filtros de Visão/IFRS_9/IFRS_15 tratam valor vazio ou igual ao rótulo padrão como
      wildcard SQL "%" (linhas 260-266) — ver RN-006.
    - Nenhum tratamento de erro (`On Error`) envolvendo o loop principal — se `Extrair_Linha` falhar no meio do
      loop (rede instável), a função para sem *rollback* da tabela auxiliar já parcialmente inserida.
    - Uso de `Sh3.Select`/`ActiveSheet.ShowAllData`/`ActiveWorkbook.RefreshAll` (padrão Select/Activate).
    - Evidência: linhas 100-134 (filtros Empresa/IFRS/Organic), 162-287 (loop principal), 276-281 (CdC),
      299-336 (redefinição KPI/Versão).

### 16. Criar_TB_SQL_AUX

1. **Nome completo:** `Criar_TB_SQL_AUX`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_SQL_Hubble.bas:365`
5. **Objetivo (negócio):** Criar dinamicamente, no SQL Server, uma tabela temporária de *staging*
   (`TB_AUX_HUBBLE_QUICK_DATA_<usuário>`) com schema derivado das colunas da aba Base, para receber os
   resultados agregados de cada linha de extração antes de trazê-los ao Excel.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Hubble` (linha 67). Não referenciado fora do módulo.
7. **Procedimentos chamados:** `AbreConexao` (linha 367), `FechaConexao` (linha 412, label `FimdaMacro`).
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `Comando`/`Info` (`ADODB.Command`/`Recordset`), `Campos_Criacao` (string DDL
    montada coluna a coluna a partir de `Sh3`); tipos especiais: Empresa/IFRS_Contábil → `VARCHAR(2)`, Organic →
    `FLOAT`, demais → tipo declarado na linha 1 da aba Base.
11. **Abas/intervalos/arquivos acessados:** `Sheet3` (linha 1 = tipos de dado por coluna; linha "LIN_BASE" =
    nomes de coluna). Tabela criada: `[BPAM].[dbo].[TB_AUX_HUBBLE_QUICK_DATA_<Environ("UserName")>]`.
12. **Pré-condições:** conexão SQL disponível; usuário com permissão `CREATE TABLE` no schema `BPAM.dbo`.
13. **Passos principais:** monta lista de colunas+tipos a partir da aba Base (linhas 385-396); executa
    `CREATE TABLE` (linha 399).
14. **Pós-condições:** tabela de *staging* criada, vazia, pronta para `INSERT`s.
15. **Efeitos colaterais/erro/mensagens/riscos:** `On Error GoTo tratar_erro` **ausente/inativo** — o label
    `tratar_erro:` só é alcançável via `GoTo` explícito, nunca disparado automaticamente. Mensagem de erro
    (literal, linha 409): "Ocorreu um erro ao atualizar [Valor]. O erro é comum quando há oscilação da rede.
    Tente novamente e, se o problema persistir, por favor envie um e-mail para 'dfigsilva@timbrasil.com.br' com
    um PrintScreen dessa tela para ter o erro corrigido." — evidencia contato de suporte humano hardcoded. Nome
    de tabela ligado ao usuário do SO — risco de tabela "presa" se a exclusão (`Excluir_TB_SQL_AUX`) não rodar
    após falha. `Comando.CommandTimeout = 1000` (~16,6 min) sugere DDL historicamente lento/tabela grande.
    Evidência: linhas 399-409.

### 17. Excluir_TB_SQL_AUX

1. **Nome completo:** `Excluir_TB_SQL_AUX`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_SQL_Hubble.bas:417`
5. **Objetivo (negócio):** Remover a tabela de *staging* temporária criada por `Criar_TB_SQL_AUX`, tanto antes
   (limpeza preventiva de execução anterior malsucedida) quanto depois do processo de extração.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Hubble`, duas vezes (linhas 66 e 341 — início e fim do
   processo).
7. **Procedimentos chamados:** `AbreConexao` (419), `FechaConexao` (444).
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Comando`/`Info` ADODB.
11. **Tabela:** `IF EXISTS(...) DROP TABLE BPAM.DBO.TB_AUX_HUBBLE_QUICK_DATA_<usuário>` (linha 431).
12. **Pré-condições:** nenhuma (uso de `IF EXISTS` torna a operação idempotente).
13. **Passos principais:** monta e executa `DROP` condicional.
14. **Pós-condições:** tabela de *staging* garantidamente inexistente.
15. **Riscos:** mesma mensagem de erro genérica com contato de suporte (linha 441); sem `On Error` ativo.
    Evidência: linhas 431, 441.

### 18. Extrair_Linha

1. **Nome completo:** `Extrair_Linha`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub — **núcleo do SQL dinâmico do módulo**
4. **Escopo:** `Private` — `Extracao_SQL_Hubble.bas:448`
5. **Objetivo (negócio):** Montar e executar um único `INSERT INTO ... SELECT ... GROUP BY` contra a tabela de
   origem do Hubble, agregando valores mensais e anuais (FY) por combinação de dimensões configuradas, e
   inserir o resultado na tabela de *staging*.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Hubble` (linha 284), uma vez por linha "Sim" da matriz de
   extração.
7. **Procedimentos chamados:** `AbreConexao` (450), `FechaConexao` (515).
8. **Parâmetros (todos Variant implícito, `ByRef` padrão VBA, sem `ByVal`/tipagem declarada):** `Sh8`, `Sh3`
   (Worksheets), `Lin` (nº da linha da matriz, gravado como `Lin_Base` em cada registro), `Filtr_Lin` (cláusula
   WHERE), `Filtr_Col`/`Filtr_Col_Extra` (colunas de agrupamento), `Filtro_CC`/`Filtro_CDC` (regras de
   Classe/Centro de Custo), `Classif_Extra`/`Classif_Organic` (`CASE WHEN` adicionais),
   `Filtr_Col_Soma_Meses`, `SomaMensal` (expressões de soma condicional por mês).
9. **Retorno:** N/A.
10. **Variáveis relevantes:** `inicio_sql`/`final_sql`/`final_sql_2` (tratam caso especial de "DESCRICAO CC" com
    `LEFT JOIN` em `TB_HUBBLE_DBS_CC`).
11. **Abas/arquivos/tabelas acessados:** tabela de origem `[BPAM].[dbo].[<Sheet24!D19>]`; tabela de *staging*
    destino `TB_AUX_HUBBLE_QUICK_DATA_<usuário>`; `LEFT JOIN` opcional com `[BPAM].[dbo].[TB_HUBBLE_DBS_CC]`.
12. **Pré-condições:** tabela de *staging* já criada (por `Criar_TB_SQL_AUX`); parâmetros de filtro já montados
    pela chamadora.
13. **Passos principais:** monta `Group_CC`/`Group_CDC` (linhas 465-471); decide se inclui `LEFT JOIN` com
    tabela de descrição de CC (linhas 474-483); monta comando
    `INSERT...SELECT...FROM...WHERE...GROUP BY...ORDER BY` (linhas 485-498); executa via `Info.Open Comando`
    (linha 504).
14. **Pós-condições:** linhas agregadas inseridas na tabela de *staging*.
15. **Efeitos colaterais/erro/riscos:** **SQL montado 100% por concatenação de string, sem parametrização** —
    qualquer apóstrofo em nome de KPI/Versão/Empresa quebra a query ou permite injeção. `CommandTimeout=1000`.
    Mesma ausência de `On Error` ativo. Mensagem de erro genérica igual às demais (linha 512). Evidência:
    linhas 474-498 (montagem do SQL), 502 (timeout).

### 19. Extrair_Base_Final

1. **Nome completo:** `Extrair_Base_Final`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_SQL_Hubble.bas:520`
5. **Objetivo (negócio):** Ler o conteúdo consolidado da tabela de *staging* (já com todas as linhas inseridas
   por múltiplas chamadas de `Extrair_Linha`) e colar o resultado na aba Base via `CopyFromRecordset`.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Hubble` (linha 293), uma única vez após o loop de
   extração.
7. **Procedimentos chamados:** `AbreConexao` (522), `FechaConexao` (567).
8. **Parâmetros:** `Sh3` (Worksheet destino), `Lin_Cabecalho` (recebido mas não usado no corpo ativo),
   `Configuracao_CDC` (recebido mas não usado no corpo ativo — linha que o usaria, 554, está comentada).
9. **Retorno:** N/A.
10. **Variáveis:** `Lin_Destino` (calculada por `CountA` da coluna A + 4, linha 551).
11. **Abas/tabelas acessadas:** query `SELECT A.* FROM [BPAM].[dbo].[TB_AUX_HUBBLE_QUICK_DATA_<usuário>] A
    ORDER BY 1,2,3,4,5` (linhas 539-542); destino `Sheet3`.
12. **Pré-condições:** tabela de *staging* populada.
13. **Passos principais:** executa `SELECT`; se não vazio, `Sh3.Cells(Lin_Destino,1).CopyFromRecordset Info`
    (linha 555).
14. **Pós-condições:** dados agregados colados na aba Base a partir da primeira linha livre.
15. **Riscos:** parâmetros `Lin_Cabecalho` e `Configuracao_CDC` recebidos mas não efetivamente usados no código
    ativo (linha 554 comentada) — código morto/parâmetro vestigial, indício de refatoração incompleta.
    `CommandTimeout=1000`. Evidência: linhas 537, 551, 554 (comentada), 555.

### 20. Retirar_Duplicadas

1. **Nome completo:** `Retirar_Duplicadas`
2. **Módulo:** `Extracao_SQL_Hubble.bas`
3. **Tipo:** Sub — **rotina quebrada/incompleta**
4. **Escopo:** `Private` — `Extracao_SQL_Hubble.bas:575`
5. **Objetivo (negócio) aparente:** remover linhas duplicadas na Base com base na combinação Classe Custo +
   Centro Custo.
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Retirar_Duplicadas` em todo o projeto (o `grep`
   confirma apenas a linha de declaração) — não referenciado no código lido; provável resquício de
   desenvolvimento nunca finalizado nem ligado a um botão.
7. **Procedimentos chamados:** nenhum.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A.
10. **Variáveis:** `Col_CC`/`Col_CDC` (posições via `Match`), `Col_Check_Dupl_1`/`Col_Check_Dupl_2` (colunas
    auxiliares criadas após "Dezembro"), fórmulas `=RC&RC` (chave concatenada) e `=COUNTIFS(...)`.
11. **Abas:** `Sheet3` (Base).
12. **Pré-condições:** N/A (rotina aparentemente nunca executada em produção).
13. **Passos principais:** cria duas colunas auxiliares com fórmulas de chave/contagem (linhas 589-593); **o
    loop `For Lin = Ult_Lin To Lin_1 Step -1` está VAZIO** (linhas 595-599, sem nenhuma instrução dentro) — a
    lógica de exclusão de duplicatas nunca foi implementada.
14. **Pós-condições:** colunas auxiliares "Check_Dupl_Chave" criadas na Base, mas nenhuma linha é de fato
    removida.
15. **Risco:** código morto/incompleto que, se algum dia for acionado, cria colunas auxiliares sem completar a
    limpeza, deixando a planilha com colunas extras não removidas — e exibe uma `MsgBox` de "sucesso" enganosa
    (linha 601: "Limpeza concluída com sucesso!" mesmo sem nenhuma exclusão real ter ocorrido) —
    inconsistência entre mensagem e comportamento real. Evidência: linhas 589-601.

---

## Módulo `Extracao_Base_1009.bas`

### 21. Extrair_Base_1009

1. **Nome completo:** `Extrair_Base_1009`
2. **Módulo:** `Extracao_Base_1009.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_1009.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para importar a "Base 1009" (relatório contábil externo por
   Classe/Centro de Custo), aplicar tratamento IFRS16 e atualizar listas de KPI/Versão.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_1009` (linha 12),
   `Atualizar_Lista_KPI_Versao_Interna`, `UPDATE_Tratar_IFRS16(Chave)` (linha 16, `fx_IFRS16.bas`),
   `PopUp_Tempo_Processamento`, `Ativar_Tudo`. **Nota:** `Processo_Limpar_Base_1009` está **comentada** (linha
   10) — a limpeza prévia da fonte NÃO ocorre automaticamente nesta Sub.
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_1009" (linhas 4-5, usadas depois no `UPDATE_Tratar_IFRS16`).
11. **Abas/arquivos:** `Sheet8` (`.Select`, linha 15).
12. **Pré-condições:** aba Extração com caminho/arquivo "1009" configurado (via `GetArquivo`).
13. **Passos principais:** cronometra início; verifica versão; desliga UI; extrai; atualiza listas; aplica
    IFRS16; popup de tempo; reativa UI.
14. **Pós-condições:** base 1009 atualizada na aba Base, já com tratamento IFRS16 aplicado.
15. **Regra de negócio/riscos:** confirma que a base 1009, ao contrário da Hubble, **recebe tratamento IFRS16**
    (`UPDATE_Tratar_IFRS16`, linha 16 ativa) — ver RN-007. Como `Processo_Limpar_Base_1009` está desativada
    aqui, execuções repetidas sem uma limpeza manual prévia tendem a duplicar registros da fonte "Base_1009" —
    ver risco de duplicação (RN-002). Evidência: linhas 10, 16.

### 22. Processo_Limpar_Base_1009

1. **Nome completo:** `Processo_Limpar_Base_1009`
2. **Módulo:** `Extracao_Base_1009.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_1009.bas:23`
5. **Objetivo (negócio):** Remover da Base os registros históricos das fontes "Base_1009" e
   "Base_1009_IFRS16 Tratado" antes de nova extração, com confirmação visual ao usuário.
6. **Quem chama:** chamada **comentada** em `Extrair_Base_1009` (linha 10) — não ativa; não referenciado de
   outra forma no código lido — provável botão/shape dedicado à limpeza, separado do botão de extração.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica` (duas vezes: Chave="Base_1009" linha 28,
   depois Chave="Base_1009_IFRS16 Tratado" linha 38), `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte"; `Chave` reatribuída duas vezes (linhas 28 e 37). Duas chamadas adicionais
    para chaves "Base_1009 - AJUSTE RATEIO VBA - EXCLUSAO"/"- RATEIO" estão **comentadas** (linhas 31-35) —
    indício de uma funcionalidade de ajuste de rateio descontinuada.
11. **Abas:** `Sheet8` (`.Select`, linha 39).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa fonte "Base_1009"; limpa fonte "Base_1009_IFRS16 Tratado"; seleciona
    Sheet8; reativa UI; `MsgBox` de sucesso.
14. **Pós-condições:** registros de ambas as fontes removidos da Base.
15. **Riscos:** código morto relativo a "AJUSTE RATEIO VBA" (linhas 31-35) sugere feature de rateio de custos
    descontinuada sem remoção do vestígio — **[VALIDAR COM O NEGÓCIO]** se ainda é relevante para a
    reconstrução. Evidência: linhas 27-38.

### 23. Processo_Extrair_Base_1009

1. **Nome completo:** `Processo_Extrair_Base_1009`
2. **Módulo:** `Extracao_Base_1009.bas`
3. **Tipo:** Function (retorno não usado) — ≈420 linhas, a mais extensa do módulo
4. **Escopo:** `Public` — `Extracao_Base_1009.bas:50`
5. **Objetivo (negócio):** Abrir o arquivo externo "1009", normalizar sua estrutura mensal (colunas por mês),
   remover linhas zeradas/cabeçalho, copiar Centro Custo/Classe Custo/Exercício/meses para a Base, e enriquecer
   com todas as dimensões padrão (Organic, IFRS_Contábil, Empresa, Segmento, Diretoria etc.), com opção
   interativa de excluir linhas de "Other Income".
6. **Quem chama (evidência):** `Extrair_Base_1009` (linha 12, cluster); `Auxiliar.bas:18`
   (`Extrair_Todas_as_Bases`, fora do cluster).
7. **Procedimentos chamados:** `Transformar_Texto_Mes_Em_Valor` (linha 132, `Aux_Formulas_Base.bas`),
   `Apagar_Linhas_Zeradas` (214), `Copiar_Base_Origem` (4 chamadas: Centro Custo 233, Classe Custo 241,
   Exercício 249, meses em loop 263), `Form_Preenchimento_Generico` (7 chamadas: LIN_BASE 289, TIPO NÍVEL 2 296,
   FONTE 303, KPI 310, VERSÃO 317, VISAO 324, PROFORMA 346), `Form_Organic` (329), `Form_Ref_Organic` (334),
   `Form_IFRS_Contabil` (339), `Form_Opex_Driven` (351), `Form_Grupo_BD` (356), `Form_Empresa` (361),
   `Form_Segmentos` (366, com literal "1009"), `Form_Classe` (371), `Form_Diretoria_Gerencial_Com_Ref_Cruzada`
   (376), `Form_Zerar_Meses_Exceto_Um` (381), `Form_Calcular_FY` (386), `Gerar_Visao_Italia` (391),
   condicionalmente (se `Sheet8!I25 <> "Não"`): `Reclassificar_Combinacoes_Empresas` (397), `_IFRS_Contabil`
   (398), `_Proforma` (399), `Calcular_Comb_Meses_Intervalo_Linha` (400).
8. **Parâmetros:** nenhum.
9. **Retorno:** Function sem valor atribuído.
10. **Variáveis/objetos relevantes:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Aux_2`=`Sheet2`, `Sh_Destino`=
    `Sheet3`, `Sh_Destino_DM`=`Sheet31`; `Diret`/`Arq` (caminho/arquivo, lidos por offset de célula a partir de
    `Cells.Find(What:="1009")`); `Mes_1009_Inicial`/`Mes_1009_Final`.
11. **Abas/intervalos/arquivos acessados:** arquivo externo `.xlsx`/`.xls` cujo caminho vem de `Sheet8` (não
    hardcoded no VBA); `Sheet15`, `Sheet2`, `Sheet3`, `Sheet31`.
12. **Pré-condições:** célula com o texto "1009" deve existir em `Sheet8`; arquivo apontado deve existir e ter
    aba 1 com cabeçalho "Jan*" localizável.
13. **Passos principais:**
    - Localiza célula-âncora "1009" em `Sheet8`; lê `Diret`/`Arq`/meses inicial-final (linhas 67-81).
    - Abre o arquivo externo somente leitura (linha 90).
    - Corrige encoding quebrado: `Replace "Mar‡o"→"Março"`, `"Exerc¡cio"→"Exercício"` (linhas 97-98).
    - Localiza cabeçalho "Jan*" (linha 102); calcula última linha (linhas 110-119).
    - Remove linhas zeradas logo após o cabeçalho (linhas 183-187); ordena a planilha de origem (linhas
      194-208); remove linhas totalmente zeradas via `Apagar_Linhas_Zeradas` (linha 214).
    - Copia Centro Custo/Classe Custo/Exercício/meses para a Base (linhas 225-265).
    - Fecha o arquivo externo (linhas 270-272).
    - Preenche campos fixos (LIN_BASE, TIPO NÍVEL 2, FONTE, KPI, VERSÃO, VISAO) e roda a bateria de
      enriquecimento (linhas 285-391).
    - Pergunta via `MsgBox` (linha 406) se deve excluir linhas de "Other Income" da base 1009; se sim, ordena e
      remove essas linhas (linhas 417-461).
14. **Pós-condições:** linhas da fonte "Base_1009" inseridas na aba Base, com Visão="IFRS Brasil".
15. **Efeitos colaterais/erro/regra de negócio/riscos:**
    - **MsgBox bloqueante no meio do processo em lote** (linha 406: "Deseja apagar as informações de ''Other
      Income'' da base 1009 !?") — impede automação/agendamento desatendido.
    - Grandes blocos de código comentado (linhas 138-178, 405-463) — lógica antiga duplicada/morta misturada ao
      código ativo.
    - Uso extensivo de `Select`/`ActiveWorkbook`/`Windows(Arq).Activate`.
    - Ordenação via `Sort.SortFields` na planilha de origem inteira antes de processar — custoso.
    - Encoding corrompido tratado via `Replace` hardcoded — sintoma de charset incompatível entre sistemas, não
      correção robusta.
    - Evidência: linhas 97-98, 183-187, 406, 417-461.

---

## Módulo `Extracao_Base_Consolidad.bas`

### 24. Extrair_Base_Consolidada

1. **Nome completo:** `Extrair_Base_Consolidada`
2. **Módulo:** `Extracao_Base_Consolidad.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_Consolidad.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para importar a "Base Consolidada" (prévia contábil por
   Conta Contábil × Centro de Custo).
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_Consolidada` (linha
   13), `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`. **Nota:**
   `Processo_Limpar_Base_Consolidada` está **comentada** (linha 11) — não executada automaticamente.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Consolidada".
11. **Abas:** `Sheet8` (`.Select`, linha 16).
12. **Pré-condições:** aba Extração com caminho "Base Consolidada" configurado.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup de tempo;
    reativa UI.
14. **Pós-condições:** base Consolidada atualizada na aba Base.
15. **Riscos:** limpeza prévia desativada — mesma observação de duplicação potencial do item 21 (ver RN-002).
    Evidência: linha 11.

### 25. Processo_Limpar_Base_Consolidada

1. **Nome completo:** `Processo_Limpar_Base_Consolidada`
2. **Módulo:** `Extracao_Base_Consolidad.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_Consolidad.bas:24`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_Consolidada" antes de nova extração, com
   confirmação visual.
6. **Quem chama:** chamada comentada em `Extrair_Base_Consolidada` (linha 11) — não ativa; não referenciado de
   outra forma — provável botão/shape dedicado.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Consolidada".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** nenhum tratamento de erro próprio. Evidência: linhas 26-37.

### 26. Processo_Extrair_Base_Consolidada

1. **Nome completo:** `Processo_Extrair_Base_Consolidada`
2. **Módulo:** `Extracao_Base_Consolidad.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_Consolidad.bas:41`
5. **Objetivo (negócio):** Abrir o arquivo externo "Base Consolidada", remover linhas em branco/zeradas, mapear
   Conta Contábil→Abertura_2 via tabela De-Para, copiar Fonte/Ano/Conta/Centro de Custo/valor do mês de prévia
   para a Base, e oferecer exclusão interativa de linhas de receita já cobertas por outras fontes (evitar dupla
   contagem).
6. **Quem chama (evidência):** `Extrair_Base_Consolidada` (linha 13); `Auxiliar.bas:16`
   (`Extrair_Todas_as_Bases`).
7. **Procedimentos chamados:** `Copiar_Base_Preview` (5 chamadas: Fonte 121, Ano Previa→EXERC* 135, Conta
   Contábil→CLASSE CUSTO 142, Centro de Custo→CENTRO CUSTO 149, Valor Previa→mês 189; uma 6ª chamada, para
   "Previa"→"Versão", está **comentada**, linha 128), `Form_Zerar_Meses_Exceto_Um` (160), `Form_Preenchimento_Generico`
   (6 chamadas: LIN_BASE 198, TIPO NÍVEL 2 205, KPI 212, VERSÃO 219, VISAO 226, PROFORMA 248), `Form_Organic`
   (231), `Form_Ref_Organic` (236), `Form_IFRS_Contabil` (241), `Form_Opex_Driven` (253), `Form_Grupo_BD` (258),
   `Form_Empresa` (263), `Form_Segmentos` (268, literal "BASE CONSOLIDADA"), `Form_Classe` (273),
   `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (278), `Form_Calcular_FY` (283), `Gerar_Visao_Italia` (288),
   `Reclassificar_Combinacoes_Empresas` (298), `_IFRS_Contabil` (299), `_Proforma` (300),
   `Calcular_Comb_Meses_Intervalo_Linha` (301), `Excluir_Linhas_Base_Consolidada` (302).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Destino`=`Sheet3`; `Mes_Preview` (mês da prévia,
    3 primeiras letras); `Col_A2` (coluna auxiliar temporária criada/removida na planilha de origem, linhas
    164-183).
11. **Abas/intervalos/arquivos acessados:** arquivo externo "Base Consolidada" (caminho via `Sheet8`); `Sheet15`
    (De-Para Conta Contábil→Abertura_2); `Sheet3`.
12. **Pré-condições:** célula "BASE CONSOLIDADA" em `Sheet8`; arquivo apontado com colunas "Conta Contabil",
    "Centro*Custo", "Valor Previa", "Ano Previa", "Fonte".
13. **Passos principais:**
    - Abre arquivo externo (linha 76); remove linhas em branco em 3 colunas via `SpecialCells(xlCellTypeBlanks)`
      dentro de um único `On Error Resume Next` (linhas 86-97).
    - Remove linhas com Valor Prévia = 0, em loop decrescente linha-a-linha recalculando `Ult_Lin_Origem` a cada
      exclusão (linhas 102-110).
    - Copia Fonte/Ano/Conta/Centro de Custo via `Copiar_Base_Preview` (linhas 119-150).
    - Zera colunas FY→Dez (linha 156) e aplica `Form_Zerar_Meses_Exceto_Um` (linha 160).
    - Injeta fórmula `INDEX/MATCH` temporária na planilha de origem para mapear Conta Contábil→Abertura_2
      (linhas 164-170), depois apaga a coluna auxiliar (linha 183).
    - Copia valor do mês de prévia (linha 189).
    - Preenche campos fixos e roda enriquecimento padrão (linhas 196-289).
    - Fecha arquivo externo (linhas 292-294).
    - Reclassifica combinações e oferece exclusão de linhas de Product/Service Revenues (linhas 298-302).
14. **Pós-condições:** linhas da fonte "Base_Consolidada" na aba Base.
15. **Efeitos colaterais/erro/riscos:**
    - Loop de exclusão de linhas zeradas com `Lin = Lin - 2` (linha ~106) — decremento manual após deletar
      linha, propenso a pular/reprocessar linhas; recalcula `Ult_Lin_Origem` (`fn.CountA`) a cada exclusão —
      complexidade O(n²).
    - `On Error Resume Next` cobrindo 3 operações `SpecialCells(...).EntireRow.Delete` seguidas (linhas 86-97)
      sem checar erro entre elas — falhas silenciosas não diferenciadas.
    - Bloco de código comentado sobre inverter sinal de "Opex Driven" (linhas 174-179) — decisão de negócio
      implementada e depois desativada sem explicação.
    - Duas confirmações `MsgBox` sequenciais no final (via `Excluir_Linhas_Base_Consolidada`) — impede
      automação total.
    - Evidência: linhas 86-97, 102-110, 174-179.

### 27. Copiar_Base_Preview

1. **Nome completo:** `Copiar_Base_Preview`
2. **Módulo:** `Extracao_Base_Consolidad.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_Consolidad.bas:307`
5. **Objetivo (negócio):** Rotina genérica de cópia de uma coluna nomeada da planilha de origem para a coluna
   correspondente na Base, usada por todos os campos simples da extração da Base Consolidada.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Consolidada`, 5 chamadas ativas (linhas 121, 135, 142,
   149, 189) + 1 comentada (128). Sem chamadas externas ao módulo.
7. **Procedimentos chamados:** nenhum da lista indexada (usa `fn.Match`, `Copy`/`PasteSpecial` nativos).
8. **Parâmetros:** `Arq` (não utilizado no corpo — vestigial), `Sh_Origem`, `Col_Origem`, `Ult_Lin_Origem`,
   `Sh_Destino`, `Lin_Cabecalho`, `Lin_Destino`, `Col_Destino`, `Col_Destino_2` (saída, ByRef).
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `fn` (`WorksheetFunction`).
11. **Abas:** lê `Sh_Origem` (linha 1 = cabeçalho); escreve `Sh_Destino`.
12. **Pré-condições:** coluna `Col_Origem` deve existir na linha 1 da origem; `Col_Destino` deve existir na
    linha de cabeçalho da Base.
13. **Passos principais:** localiza colunas por nome (`fn.Match`); copia valores via `Copy`/`PasteSpecial
    xlPasteValues` (linhas 314-316).
14. **Pós-condições:** coluna de destino preenchida com os valores da origem.
15. **Riscos:** parâmetro `Arq` recebido mas não usado — vestígio de assinatura genérica compartilhada com
    outros módulos. Sem tratamento de erro se a coluna não existir (`fn.Match` retorna erro #N/A não tratado).
    Evidência: linha 307 (assinatura), 311-312.

### 28. Excluir_Linhas_Base_Consolidada

1. **Nome completo:** `Excluir_Linhas_Base_Consolidada`
2. **Módulo:** `Extracao_Base_Consolidad.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Base_Consolidad.bas:320`
5. **Objetivo (negócio):** Perguntar interativamente ao usuário se deve remover, da extração recém-realizada,
   linhas de "Product Revenues" e/ou "Net Service Revenues" — evita dupla contagem de receita já capturada por
   outras fontes (1009, RGM etc.).
6. **Quem chama (evidência):** `Processo_Extrair_Base_Consolidada` (linha 302).
7. **Procedimentos chamados:** `Processo_Exclusao_Linhas_Base` (`Auxiliar.bas:974`), até 2 vezes.
8. **Parâmetros:** `Lin_Destino`, `Ult_Lin_Destino`.
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Col_Abertura`="Abertura_3"; `Criterios` (string com valores separados por `;`).
11. **Abas:** `Sheet3` (Base), indiretamente via `Processo_Exclusao_Linhas_Base`.
12. **Pré-condições:** linhas recém-inseridas ainda identificáveis pelo intervalo `Lin_Destino`..`Ult_Lin_Destino`.
13. **Passos principais:** `MsgBox` "Deseja excluir as informações de ''RECEITA DE PRODUTO''..." (linha 322); se
    sim, exclui linhas "Product Revenues"; `MsgBox` "Deseja excluir ... 'RECEITA DE SERVIÇO' ..." (linha 328);
    se sim, exclui linhas "Net Service Revenues".
14. **Pós-condições:** linhas de receita potencialmente duplicada removidas (ou mantidas, conforme decisão do
    usuário).
15. **Regra de negócio/riscos:** decisão de exclusão de receita é **manual a cada execução**, não uma regra
    fixa — risco de inconsistência entre execuções feitas por analistas diferentes (uma vez inclui, outra vez
    exclui) sem trilha de auditoria de qual escolha foi feita. Evidência: linhas 322, 328.

---

## Módulo `Extracao_Base_MOCKUP_RGM.bas`

### 29. Extrair_Base_MOCKUP_RGM

1. **Nome completo:** `Extrair_Base_MOCKUP_RGM`
2. **Módulo:** `Extracao_Base_MOCKUP_RGM.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_MOCKUP_RGM.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para extrair a base de simulação/mockup do RGM (Receita e
   Gross Margin).
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_MOCKUP_RGM` (linha
   12), `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_MOCKUP_RGM".
11. **Abas:** `Sheet8` (`.Select`, linha 15).
12. **Pré-condições:** aba Extração com caminho "MOCKUP" configurado.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup; reativa UI.
14. **Pós-condições:** base MOCKUP_RGM atualizada na aba Base.
15. **Riscos:** **esta Sub NÃO chama `Processo_Limpar_Base_MOCKUP_RGM` em nenhum ponto** (nem comentado) —
    diferente do padrão observado em outros módulos (que ao menos têm a chamada comentada) — risco mais alto de
    duplicação acumulativa a cada execução sem limpeza manual prévia (ver RN-002). Além disso, esta fonte **não
    é chamada** pelo orquestrador `Extrair_Todas_as_Bases` (`Auxiliar.bas`) — só roda isoladamente (ver RN-019).
    Evidência: ausência de `Call Processo_Limpar_Base_MOCKUP_RGM` no corpo desta Sub; `Auxiliar.bas:15-21`.

### 30. Processo_Limpar_Base_MOCKUP_RGM

1. **Nome completo:** `Processo_Limpar_Base_MOCKUP_RGM`
2. **Módulo:** `Extracao_Base_MOCKUP_RGM.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_MOCKUP_RGM.bas:23`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_MOCKUP_RGM" antes de nova extração.
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Processo_Limpar_Base_MOCKUP_RGM` em todo o `grep` do
   diretório — não referenciado no código lido, nem mesmo comentado dentro do próprio módulo (diferente de
   todos os outros módulos irmãos) — provável botão/shape dedicado, isolado.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_MOCKUP_RGM".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** total desacoplamento do fluxo de extração — só executa se o usuário souber que o botão existe
    separadamente. Evidência: linhas 25-36.

### 31. Processo_Extrair_Base_MOCKUP_RGM

1. **Nome completo:** `Processo_Extrair_Base_MOCKUP_RGM`
2. **Módulo:** `Extracao_Base_MOCKUP_RGM.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_MOCKUP_RGM.bas:41`
5. **Objetivo (negócio):** Abrir o arquivo externo de mockup RGM, delegar a extração linha-a-chave a
   `Processo_Extracao_Sheet_Base`, aplicar escala de valores (×1.000.000) e enriquecer com as dimensões padrão.
6. **Quem chama (evidência):** `Extrair_Base_MOCKUP_RGM` (linha 12). **Não** chamada por
   `Auxiliar.bas:Extrair_Todas_as_Bases` (ver RN-019).
7. **Procedimentos chamados:** `Processo_Extracao_Sheet_Base` (própria privada do módulo — linha 102 ativa,
   para `Sheet22`; linha 110 **comentada**, seria para `Sheet14`), `Form_Preenchimento_Generico` (7 chamadas:
   TIPO NÍVEL 2 142, FONTE 149, EXERCICIO 156, KPI 163, VERSÃO 170, VISAO 177, PROFORMA 199), `Form_Organic`
   (182), `Form_Ref_Organic` (187), `Form_IFRS_Contabil` (192), `Form_Opex_Driven` (204), `Form_Grupo_BD` (209),
   `Form_Empresa` (214), `Form_Classe` (224), `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (229),
   `Form_Acertar_Escala` (235, `Multiplicador=1000000`), `Form_Calcular_FY` (240), `Gerar_Visao_Italia` (245),
   `Reclassificar_Combinacoes_Empresas` (249), `_IFRS_Contabil` (250), `_Proforma` (251),
   `Calcular_Comb_Meses_Intervalo_Linha` (252). **`Form_Segmentos` está comentada** (linha 219) — Segmento vem
   direto da tabela de chaves, não é recalculado (ver RN-027).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído; possui `Exit Function` antecipado
   (linha 128) se `Ult_Lin_Destino < Lin_Destino` (nenhuma linha extraída).
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Aux_2`=`Sheet2`, `Sh_Destino`=`Sheet3`; `Diret`/`Arq`/
    `Sh_Origem_Name` (lidos por offset a partir da célula "MOCKUP"); `ANO`, `Mes_Inicial`, `Mes_Final`, `KPI`,
    `Versao`.
11. **Abas/arquivos acessados:** arquivo externo de mockup (caminho via `Sheet8`); `Sheet22` (chaves De-Para);
    `Sheet3` (destino).
12. **Pré-condições:** célula "MOCKUP" em `Sheet8`; aba fonte com o nome esperado
    (`Sh_Origem_Name`) dentro do arquivo externo.
13. **Passos principais:** localiza âncora "MOCKUP" e lê parâmetros (linhas 57-86); abre arquivo externo (linha
    91); chama `Processo_Extracao_Sheet_Base` para a aba de chaves `Sheet22` (linhas 99-103); fecha arquivo
    (linhas 118-120); preenche campos fixos e roda enriquecimento (linhas 130-252).
14. **Pós-condições:** linhas da fonte "Base_MOCKUP_RGM" na aba Base, com valores multiplicados por 1.000.000.
15. **Riscos:** ausência de tratamento de erro na abertura do arquivo externo; `Exit Function` silencioso sem
    aviso ao usuário se nada for extraído (linha 128). Evidência: linhas 99-103, 128, 219, 235.

### 32. Processo_Extracao_Sheet_Base (variante `Extracao_Base_MOCKUP_RGM.bas`)

1. **Nome completo:** `Processo_Extracao_Sheet_Base` — **nome duplicado em 3 módulos** (`MOCKUP_RGM.bas`,
   `RGM.bas`, `Fixed_Revenues.bas`); cada um é uma implementação distinta, escopo `Private`, sem colisão em
   tempo de execução.
2. **Módulo:** `Extracao_Base_MOCKUP_RGM.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Base_MOCKUP_RGM.bas:258`
5. **Objetivo (negócio):** Para cada linha de chave marcada "INF"/"Fórmula" na aba de chaves (`Sh_Chaves`,
   `Sheet22`), validar que a estrutura da planilha fonte ainda corresponde ao esperado, calcular/copiar os
   valores mensais e gravar Classe Custo/Centro Custo/Abertura_1/Segmento/Linha RGM na Base.
6. **Quem chama (evidência):** `Processo_Extrair_Base_MOCKUP_RGM`, mesmo módulo (linha 102 ativa; linha 110
   comentada). Não referenciado fora do módulo (é `Private`).
7. **Procedimentos chamados:** `Ativar_Tudo` (`Auxiliar.bas`, linha 334, apenas no caminho de erro, antes do
   `End`). Nenhum outro procedimento indexado.
8. **Parâmetros:** `Arq`, `Sh_Chaves`, `Sh_Destino`, `Sh_Origem_Name`, `Linhas_excluidas`, `KPI`, `Versao`,
   `Lin_Destino`, `Mes_Ano_Inicial`, `Mes_Ano_Final`, `Col_Menu_Origem`, `Lin_Cabecalho_Origem` (todos Variant
   implícito, `ByRef`).
9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `Campos_Chaves(10000)` (array de tamanho fixo, sem checagem de limite),
    `Sheet_Existente` (flag de validação de existência da aba fonte), `Qtd_Chaves` (contador).
11. **Abas acessadas:** `Sh_Chaves` (`Sheet22`, chaves De-Para); planilha fonte dentro do arquivo externo
    (nome recebido em `Sh_Origem_Name`); `Sh_Destino` (`Sheet3`).
12. **Pré-condições:** aba com nome `Sh_Origem_Name` deve existir no arquivo externo (senão `MsgBox` crítico e
    `Exit Sub`, linhas 303-306); estrutura de linhas da aba fonte deve corresponder ao rótulo esperado em
    `Sh_Chaves`.
13. **Passos principais:**
    - Verifica existência da aba fonte (linhas 294-306).
    - Para cada linha de chave "INF" (linhas 311-342): compara rótulo esperado (`Sh_Chaves`) com o rótulo real
      na planilha fonte; se divergir, `MsgBox` crítico (linhas 319-323), seleciona a linha divergente em ambas
      as planilhas, chama `Ativar_Tudo` e executa **`End`** (linha 335) — encerra TODO o processo Excel.
    - Localiza colunas de início/fim de mês por data (linhas 346-353).
    - Para cada linha "INF"/"Fórmula" (loop linhas 358-417): recalcula fórmula se aplicável (linhas 372-383);
      se houver valor não-zero no intervalo, copia para a Base e grava Classe Custo/Centro
      Custo/Abertura_1/Segmento/Linha RGM (linhas 389-411).
14. **Pós-condições:** linhas válidas da aba de chaves copiadas para a Base; em caso de divergência estrutural,
    o processo Excel é encerrado abruptamente antes de completar.
15. **Efeitos colaterais/erro/riscos:**
    - **ALTO:** `End` abrupto (linha 335) mata todo o processo Excel sem fechar conexões nem restaurar
      `Application.ScreenUpdating`/`EnableEvents`, deixando a aplicação em estado inconsistente.
    - Loop célula-a-célula com `.Copy`/`.PasteSpecial` via clipboard por linha de chave (linhas 391-393) — lento
      em bases grandes.
    - Array fixo `Campos_Chaves(10000)` sem verificação de limite — estoura se houver mais de 10.000 chaves.
    - `Application.StatusBar` atualizado a cada linha do loop — *overhead* de UI.
    - Mensagem de erro (literal, linhas 319-323): "EXISTE DIFERENÇA ENTRE A ESTRUTURA INICIAL E A ATUAL!!!!! ...
      Sheet: [nome] ... Estrutura inicial: Linha [n] | [valor esperado] ... Estrutura atual: Linha [n] |
      [valor real]".
    - Evidência: linhas 303-306, 317-335, 389-393, 403-411.

---

## Módulo `Extracao_Base_Other_Inco.bas`

### 33. Extrair_Base_Other_Income

1. **Nome completo:** `Extrair_Base_Other_Income`
2. **Módulo:** `Extracao_Base_Other_Inco.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_Other_Inco.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para varrer uma pasta com múltiplos arquivos de "Other
   Income" (por operadora) e consolidá-los na Base.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_Other_Income` (linha
   13), `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`. **Nota:**
   `Processo_Limpar_Base_Other_Income` está **comentada** (linha 11).
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_OtherIncome".
11. **Abas:** `Sheet8` (`.Select`, linha 16).
12. **Pré-condições:** aba Extração com pasta "Other income" configurada.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup; reativa UI.
14. **Pós-condições:** base Other Income atualizada na aba Base.
15. **Riscos:** limpeza prévia desativada (ver RN-002). Evidência: linha 11.

### 34. Processo_Limpar_Base_Other_Income

1. **Nome completo:** `Processo_Limpar_Base_Other_Income`
2. **Módulo:** `Extracao_Base_Other_Inco.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_Other_Inco.bas:24`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_OtherIncome" antes de nova extração.
6. **Quem chama:** chamada comentada em `Extrair_Base_Other_Income` (linha 11) — não ativa; provável
   botão/shape dedicado.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_OtherIncome".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** nenhum tratamento de erro próprio. Evidência: linhas 26-37.

### 35. Processo_Extrair_Base_Other_Income

1. **Nome completo:** `Processo_Extrair_Base_Other_Income`
2. **Módulo:** `Extracao_Base_Other_Inco.bas`
3. **Tipo:** Function (retorno não usado) — **única rotina do cluster que itera uma pasta inteira de arquivos**
4. **Escopo:** `Public` — `Extracao_Base_Other_Inco.bas:42`
5. **Objetivo (negócio):** Percorrer todos os arquivos de uma pasta que casem com um critério de nome
   (wildcard), extrair de cada um os valores contábeis de "Contas de resultados", inferir o Centro de Custo a
   partir do nome do arquivo, e consolidar tudo na Base, com opção de excluir linhas de receita duplicada.
6. **Quem chama (evidência):** `Extrair_Base_Other_Income` (linha 13); `Auxiliar.bas:19`
   (`Extrair_Todas_as_Bases`).
7. **Procedimentos chamados:** `Transformar_Texto_Mes_Em_Valor` (linha 114, `Aux_Formulas_Base.bas`),
   `Apagar_Linhas_Zeradas` (119), `Copiar_Base_PreClosing` (164, mesmo módulo), `Copiar_Base_Origem` (178,
   dentro do loop de meses, `Aux_Formulas_Base.bas`), `Form_Preenchimento_Generico` (9 chamadas: Centro Custo
   217, LIN_BASE 245, TIPO NÍVEL 2 252, FONTE 259, EXERCICIO 273, KPI 280, VERSÃO 287, VISAO 294, PROFORMA 316),
   `Form_Organic` (299), `Form_Ref_Organic` (304), `Form_IFRS_Contabil` (309), `Form_Opex_Driven` (321),
   `Form_Grupo_BD` (326), `Form_Empresa` (331), `Form_Segmentos` (336, literal "Other Income"), `Form_Classe`
   (341), `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (346), `Form_Zerar_Meses_Exceto_Um` (351),
   `Form_Acertar_Sinal_Revenues` (356), `Form_Calcular_FY` (361), `Gerar_Visao_Italia` (366),
   `Reclassificar_Combinacoes_Empresas` (370), `_IFRS_Contabil` (371), `_Proforma` (372),
   `Calcular_Comb_Meses_Intervalo_Linha` (373), `Excluir_Linhas_Base_Other_Income` (374, mesmo módulo).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis relevantes:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Destino`=`Sheet3`; `Diret`/`Criterio`
    (padrão wildcard de nome de arquivo); `Arq_Extraidos` (string `;`-separada de arquivos já processados);
    `CDC` (Centro de Custo inferido por nome de arquivo, linhas 201-211).
11. **Abas/arquivos acessados:** pasta inteira (via `Dir(Diret & Criterio)`); cada arquivo tem aba 1 com
    cabeçalho "Contas de resultados"; `Sheet3` (destino).
12. **Pré-condições:** pasta configurada em `Sheet8`; cada arquivo deve ter a estrutura "Contas de resultados" e
    a célula "apurado" (para extrair o Exercício).
13. **Passos principais:**
    - Normaliza critério wildcard (linhas 78-81); itera `Dir()` (loop `While Arq <> ""`, linhas 84-232).
    - Para cada arquivo: abre; localiza cabeçalho "Contas de resultados" (linha 102); localiza última linha
      "Total" (linha 110); converte texto de mês em valor e remove linhas zeradas (linhas 114-120).
    - Extrai Classe Custo do início da string da célula, removendo linhas de total/cabeçalho/subtotal (`.` no
      primeiro token) — loop linha-a-linha com `.EntireRow.Delete` (linhas 127-153).
    - Copia Classe Custo tratada e meses para a Base (linhas 162-181).
    - Localiza "Exercício" via célula "apurado" (linhas 186-189).
    - Infere Centro de Custo pelo nome do arquivo — 8 padrões `InStr` fixos + *fallback* (linhas 201-211).
    - Fecha arquivo; registra como processado; avança para o próximo (linhas 221-230).
    - Após o loop: preenche campos fixos e roda enriquecimento padrão (linhas 241-374).
14. **Pós-condições:** linhas de todos os arquivos da pasta consolidadas na fonte "Base_OtherIncome".
15. **Efeitos colaterais/erro/regra de negócio/riscos:**
    - **RN-004:** mapeamento de Centro de Custo hardcoded por padrão de nome de arquivo (TBP→XPHP999,
      CTCEL→XPHR999, TPAR→XPHS100, FIBER_RS→XPFR000, FIBER_RR→XPRU999, FIBER_OG→XPFS020, METIS_CZ→XPMRJ000,
      INTELIG_OG→XPHI999, *fallback*→XPHR999) — qualquer novo arquivo/operadora exige alteração de código, não
      de configuração.
    - Parsing de Classe de Custo por checagem de dígito via 10 `InStr` sequenciais (linhas 132-136) —
      ineficiente comparado a regex/`Like`.
    - Loop com `.EntireRow.Delete` dentro do próprio loop de varredura (linhas 138-140) — risco de off-by-one.
    - `Arq_Extraidos` como string concatenada em vez de coleção/dicionário — estrutura de dados improvisada.
    - Sem tratamento de erro se um arquivo da pasta não tiver a estrutura esperada (`Cells.Find` sem checar
      `Is Nothing`).
    - Evidência: linhas 201-211, 132-153.

### 36. Copiar_Base_PreClosing

1. **Nome completo:** `Copiar_Base_PreClosing`
2. **Módulo:** `Extracao_Base_Other_Inco.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Base_Other_Inco.bas:379`
5. **Objetivo (negócio):** Rotina genérica de cópia de uma coluna nomeada (usada aqui para "Col_CC_Tratada" →
   "CLASSE CUSTO") da origem para a Base.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Other_Income` (linha 164). Não referenciada fora do
   módulo.
7. **Procedimentos chamados:** nenhum da lista indexada.
8. **Parâmetros:** `Arq` (não utilizado no corpo), `Sh_Origem`, `Lin_Cabecalho_Origem`, `Col_Origem`,
   `Ult_Lin_Origem`, `Sh_Destino`, `Lin_Cabecalho`, `Lin_Destino`, `Col_Destino`, `Col_Destino_2` (saída).
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `fn` (`WorksheetFunction`).
11. **Abas:** lê `Sh_Origem`; escreve `Sh_Destino`.
12. **Pré-condições:** coluna `Col_Origem` deve existir na linha de cabeçalho de origem.
13. **Passos principais:** localiza colunas por `fn.Match`; copia via `Copy`/`PasteSpecial xlPasteValues`
    (linhas 387-389).
14. **Pós-condições:** coluna de destino preenchida.
15. **Riscos:** parâmetro `Arq` vestigial (não usado); mesma ausência de tratamento de erro que
    `Copiar_Base_Preview`/`Copiar_Base_Origem`. Evidência: linha 379 (assinatura).

### 37. Excluir_Linhas_Base_Other_Income

1. **Nome completo:** `Excluir_Linhas_Base_Other_Income`
2. **Módulo:** `Extracao_Base_Other_Inco.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Base_Other_Inco.bas:393`
5. **Objetivo (negócio):** Remover automaticamente linhas de custo ("Market/Process/Labour/Volume Driven
   Costs") sempre, e perguntar interativamente se deve remover "Net Service Revenues" — evita dupla contagem de
   receita já capturada por outra fonte.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Other_Income` (linha 374).
7. **Procedimentos chamados:** `Processo_Exclusao_Linhas_Base` (`Auxiliar.bas:974`), 2 vezes (uma
   automática, linha 397; uma condicionada a `MsgBox`, linha 408).
8. **Parâmetros:** `Lin_Destino`, `Ult_Lin_Destino`.
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Col_Abertura`="Abertura_3"; `Criterios` = ";Market Driven Costs;Process Driven Costs;Labour
    Cost;Volume Driven Costs;-;" (exclusão automática, linha 396) e ";Net Service Revenues;" (exclusão
    condicionada, linha 407).
11. **Abas:** `Sheet3` (Base), indiretamente.
12. **Pré-condições:** linhas recém-inseridas identificáveis pelo intervalo `Lin_Destino`..`Ult_Lin_Destino`.
13. **Passos principais:** exclui automaticamente as 4 categorias de custo listadas (sem confirmação); pergunta
    via `MsgBox` (linha 405) se deve excluir "Net Service Revenues"; bloco para "Product Revenues" está
    **comentado** (linhas 399-403, nunca executado).
14. **Pós-condições:** linhas de custo já cobertas por outras fontes são sempre removidas; receita de serviço é
    removida conforme decisão manual do usuário.
15. **Regra de negócio/riscos:** diferente de `Excluir_Linhas_Base_Consolidada` (que pergunta para AMBAS as
    categorias), aqui a exclusão de custo é automática e apenas a de receita de serviço é interativa —
    inconsistência de padrão entre módulos irmãos que tratam do mesmo tipo de decisão de negócio (evitar dupla
    contagem). Evidência: linhas 396-408.

---

## Módulo `Extracao_Base_Quick_Data.bas`

### 38. Extrair_Base_Quick_Data

1. **Nome completo:** `Extrair_Base_Quick_Data`
2. **Módulo:** `Extracao_Base_Quick_Data.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_Quick_Data.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para importar um arquivo já pré-formatado no layout "Quick
   Data" — provavelmente extrato de outra instância do sistema ou arquivo já tratado por terceiros.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_Quick_Data` (linha
   13), `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`. **Nota:**
   `Processo_Limpar_Base_Quick_Data` está **comentada** (linha 11).
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Quick_Data".
11. **Abas:** `Sheet8` (`.Select`, linha 16).
12. **Pré-condições:** aba Extração com caminho "Quick Data" configurado.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup; reativa UI.
14. **Pós-condições:** base Quick Data atualizada na aba Base.
15. **Riscos:** limpeza prévia desativada (RN-002); **fonte não incluída no orquestrador**
    `Extrair_Todas_as_Bases` (RN-019) — só roda isoladamente. Evidência: linha 11; `Auxiliar.bas:15-21`.

### 39. Processo_Limpar_Base_Quick_Data

1. **Nome completo:** `Processo_Limpar_Base_Quick_Data`
2. **Módulo:** `Extracao_Base_Quick_Data.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_Quick_Data.bas:24`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_Quick_Data" antes de nova extração.
6. **Quem chama:** chamada comentada em `Extrair_Base_Quick_Data` (linha 11) — não ativa; provável botão/shape
   dedicado.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Quick_Data".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** nenhum tratamento de erro próprio. Evidência: linhas 26-37.

### 40. Processo_Extrair_Base_Quick_Data

1. **Nome completo:** `Processo_Extrair_Base_Quick_Data`
2. **Módulo:** `Extracao_Base_Quick_Data.bas`
3. **Tipo:** Function (retorno não usado) — **a rotina mais simples/direta do cluster**
4. **Escopo:** `Public` — `Extracao_Base_Quick_Data.bas:42`
5. **Objetivo (negócio):** Abrir o arquivo externo "Quick Data" (colunas já nomeadas de forma idêntica ao
   destino) e copiar diretamente coluna a coluna para a Base, com enriquecimento parcial (o arquivo fonte já
   traz Organic, Proforma, Abertura_1 e Segmento prontos, dispensando o recálculo).
6. **Quem chama (evidência):** `Extrair_Base_Quick_Data` (linha 13). **Não** chamada por
   `Extrair_Todas_as_Bases` (RN-019).
7. **Procedimentos chamados:** `Copiar_Base_Origem` (9 chamadas: EMPRESA 107, CENTRO CUSTO 116, CLASSE CUSTO
   124, EXERCICIO 132, TIPO NÍVEL 2 140, ORGANIC 148, PROFORMA 156, ABERTURA_1 164, SEGMENTO 172, + meses em
   loop 192), `Form_Preenchimento_Generico` (5 chamadas: LIN_BASE 218, FONTE 225, KPI 232, VERSÃO 239, VISAO
   246), `Form_Ref_Organic` (251), `Form_IFRS_Contabil` (256), `Form_Opex_Driven` (268), `Form_Grupo_BD` (273),
   `Form_Empresa` (278), `Form_Classe` (283), `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (288),
   `Form_Zerar_Meses_Exceto_Um` (293), `Form_Calcular_FY` (298), `Gerar_Visao_Italia` (303),
   `Reclassificar_Combinacoes_Empresas` (307), `_IFRS_Contabil` (308), `_Proforma` (309),
   `Calcular_Comb_Meses_Intervalo_Linha` (310). **Nota:** `Form_Organic` e `Form_Segmentos` **não são chamadas**
   (ORGANIC e SEGMENTO já vêm copiados diretamente da origem); `PROFORMA` fixo também está **comentado**
   (linhas 259-263) pelo mesmo motivo.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Aux_2`=`Sheet2`, `Sh_Destino`=`Sheet3`,
    `Sh_Destino_DM`=`Sheet31`; `Fonte`="Base_Quick_Data"; `Mes_Quick_Data_Inicial`/`_Final`.
11. **Abas/arquivos acessados:** arquivo externo "Quick Data" (caminho via `Sheet8`, cabeçalho "FONTE"
    localizado por `Cells.Find`); `Sheet3`, `Sheet31`.
12. **Pré-condições:** célula "Quick Data" em `Sheet8`; arquivo com colunas nomeadas de forma idêntica às
    esperadas (EMPRESA, CENTRO CUSTO, CLASSE CUSTO, EXERCICIO, TIPO NÍVEL 2, ORGANIC, PROFORMA, ABERTURA_1,
    SEGMENTO, meses).
13. **Passos principais:** localiza cabeçalho "FONTE" (linha 89); copia 9 colunas simples + 12 meses via
    `Copiar_Base_Origem` (linhas 105-194); fecha arquivo (linhas 199-201); preenche campos fixos remanescentes e
    roda enriquecimento parcial (linhas 214-310).
14. **Pós-condições:** linhas da fonte "Base_Quick_Data" na aba Base.
15. **Riscos:** é o módulo mais simples do cluster (cópia direta, sem lógica de fórmula complexa), mas por isso
    também o mais silencioso sobre erros — se uma coluna esperada não existir no arquivo fonte, `fn.Match`
    (dentro de `Copiar_Base_Origem`) retorna erro #N/A não tratado, interrompendo a macro sem mensagem clara.
    Evidência: linhas 105-194, 259-263.

---

## Módulo `Extracao_Base_RGM.bas`

### 41. Extrair_Base_RGM

1. **Nome completo:** `Extrair_Base_RGM`
2. **Módulo:** `Extracao_Base_RGM.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Base_RGM.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para extrair a base "real" de Receita e Gross Margin (RGM),
   a partir de duas fontes de chaves diferentes dentro do mesmo arquivo externo.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_RGM` (linha 12),
   `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_RGM".
11. **Abas:** `Sheet8` (`.Select`, linha 15).
12. **Pré-condições:** aba Extração com caminho "(RGM)" configurado.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup; reativa UI.
14. **Pós-condições:** base RGM atualizada na aba Base.
15. **Riscos:** **esta Sub NÃO chama `Processo_Limpar_Base_RGM`, nem mesmo comentado** — mesmo padrão de risco
    de duplicação já visto em MOCKUP_RGM (item 29). Evidência: ausência de `Call Processo_Limpar_Base_RGM` no
    corpo desta Sub.

### 42. Processo_Limpar_Base_RGM

1. **Nome completo:** `Processo_Limpar_Base_RGM`
2. **Módulo:** `Extracao_Base_RGM.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_RGM.bas:23`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_RGM" antes de nova extração.
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Processo_Limpar_Base_RGM` em todo o `grep` do
   diretório, nem comentada — não referenciado no código lido — provável botão/shape dedicado, totalmente
   isolado do fluxo de extração.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_RGM".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** total desacoplamento do fluxo de extração. Evidência: linhas 25-36.

### 43. Processo_Extrair_Base_RGM

1. **Nome completo:** `Processo_Extrair_Base_RGM`
2. **Módulo:** `Extracao_Base_RGM.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Base_RGM.bas:41`
5. **Objetivo (negócio):** Abrir o arquivo externo RGM, extrair dados de **duas** abas de chaves distintas
   (`Sheet7` e `Sheet14`, representando duas visões/fontes internas do mesmo arquivo), aplicar escala
   (×1.000.000) e enriquecer com as dimensões padrão.
6. **Quem chama (evidência):** `Extrair_Base_RGM` (linha 12); `Auxiliar.bas:20` (`Extrair_Todas_as_Bases`).
7. **Procedimentos chamados:** `Processo_Extracao_Sheet_Base` (própria privada — 2 chamadas: `Sheet7`/linha 100,
   `Sheet14`/linha 108), `Form_Preenchimento_Generico` (8 chamadas: LIN_BASE 133, TIPO NÍVEL 2 140, FONTE 147,
   EXERCICIO 154, KPI 161, VERSÃO 168, VISAO 175, PROFORMA 197), `Form_Organic` (180), `Form_Ref_Organic` (185),
   `Form_IFRS_Contabil` (190), `Form_Opex_Driven` (202), `Form_Grupo_BD` (207), `Form_Empresa` (212),
   `Form_Classe` (222), `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (227), `Form_Acertar_Escala` (233,
   `Multiplicador=1000000`), `Form_Calcular_FY` (238), `Gerar_Visao_Italia` (243),
   `Reclassificar_Combinacoes_Empresas` (247), `_IFRS_Contabil` (248), `_Proforma` (249),
   `Calcular_Comb_Meses_Intervalo_Linha` (250). **`Form_Segmentos` está comentada** (linha 217, RN-027).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído; possui `Exit Function` antecipado
   (linha 126) se `Ult_Lin_Destino < Lin_Destino`.
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Aux_2`=`Sheet2`, `Sh_Destino`=`Sheet3`;
    `Diret`/`Arq`; `ANO`, `Mes_Inicial`, `Mes_Final`, `KPI`, `Versao`.
11. **Abas/arquivos acessados:** arquivo externo RGM (caminho via `Sheet8`); abas fonte com sufixo dinâmico por
    KPI (`_Act`/`_Bdg`/`_Fcst##`); `Sheet7`, `Sheet14` (chaves De-Para); `Sheet3`.
12. **Pré-condições:** célula "(RGM)" em `Sheet8`; abas fonte com sufixo correspondente ao KPI configurado
    devem existir no arquivo externo.
13. **Passos principais:** localiza âncora "(RGM)" e lê parâmetros (linhas 57-84); abre arquivo externo (linha
    89); chama `Processo_Extracao_Sheet_Base` duas vezes (linhas 97-109); fecha arquivo (linhas 116-118);
    preenche campos fixos e roda enriquecimento (linhas 128-250).
14. **Pós-condições:** linhas da fonte "Base_RGM" na aba Base, com valores multiplicados por 1.000.000.
15. **Riscos:** ausência de tratamento de erro na abertura do arquivo; `Exit Function` silencioso.
    Evidência: linhas 97-109, 126.

### 44. Processo_Extracao_Sheet_Base (variante `Extracao_Base_RGM.bas`)

1. **Nome completo:** `Processo_Extracao_Sheet_Base` — implementação específica deste módulo (ver nota no item
   32 sobre nome duplicado em 3 módulos).
2. **Módulo:** `Extracao_Base_RGM.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Base_RGM.bas:256`
5. **Objetivo (negócio):** Idêntico em estrutura ao item 32 (MOCKUP_RGM), mas para o RGM real: valida
   estrutura, calcula/copia valores mensais por linha de chave e grava dimensões na Base, incluindo a descrição
   textual da Classe de Custo.
6. **Quem chama (evidência):** `Processo_Extrair_Base_RGM`, mesmo módulo (linhas 100 e 108, para `Sheet7` e
   `Sheet14` respectivamente). `Private` — não referenciável fora do módulo.
7. **Procedimentos chamados:** `PEGA_A_DESCRICAO_DA_CONTA` (`TK_Functions.bas:2`, linha 391 — enriquece a
   Classe de Custo com sua descrição textual), `Ativar_Tudo` (`Auxiliar.bas`, linha 331, apenas no caminho de
   erro, antes do `End`).
8. **Parâmetros:** `Arq`, `Sh_Chaves`, `Sh_Destino`, `Linhas_excluidas`, `KPI`, `Versao`, `Lin_Destino`,
   `Mes_Ano_Inicial`, `Mes_Ano_Final`, `Col_Menu_Origem`, `Lin_Cabecalho_Origem` (Variant implícito, `ByRef`).
9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `Campos_Chaves(10000)` (array fixo), `Sheet_Existente`, `Qtd_Chaves`.
11. **Abas acessadas:** `Sh_Chaves` (`Sheet7`/`Sheet14`); planilha fonte no arquivo externo (nome montado por
    `Sh_Chaves.Cells(1,Col_Menu) & Compl_Origem_Name`, onde `Compl_Origem_Name` depende do KPI — `_Act`/`_Bdg`/
    `_Fcst##`); `Sh_Destino` (`Sheet3`).
12. **Pré-condições:** aba fonte com o nome montado deve existir (senão `MsgBox` crítico e `Exit Sub`, linhas
    300-303); estrutura de linhas deve corresponder ao rótulo esperado.
13. **Passos principais:** monta nome da aba fonte por KPI (linhas 280-292); verifica existência (linhas
    293-306); valida rótulos linha a linha, com `MsgBox` crítico + `End` em caso de divergência (linhas
    308-334); localiza colunas de mês por data (linhas 341-352, com `Exit Sub` se não encontrar, linha 352);
    para cada linha "INF"/"Fórmula": recalcula fórmula se aplicável e copia valores + `PEGA_A_DESCRICAO_DA_CONTA`
    (linhas 357-412).
14. **Pós-condições:** linhas válidas copiadas para a Base, incluindo descrição da Classe de Custo; ou processo
    Excel encerrado abruptamente em caso de divergência estrutural.
15. **Efeitos colaterais/erro/riscos:** mesmo risco **ALTO** de `End` abrupto do item 32 (linha 332). Mesmo
    padrão de loop célula-a-célula com clipboard (linhas 386-388) e array fixo de 10.000 posições. Nome de aba
    fonte montado dinamicamente por concatenação — se o KPI não corresponder a nenhum dos 3 padrões
    (ACT/BUDGET/FCT##), `Compl_Origem_Name` fica vazio e a aba não é encontrada, mas o erro só aparece depois,
    com mensagem genérica. Evidência: linhas 280-306, 332, 352, 386-393.

---

## Módulo `Extracao_Fixed_Revenues.bas`

### 45. Extrair_Base_Fixed_Revenues

1. **Nome completo:** `Extrair_Base_Fixed_Revenues`
2. **Módulo:** `Extracao_Fixed_Revenues.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Fixed_Revenues.bas:3`
5. **Objetivo (negócio):** Ponto de entrada (botão) para extrair receitas de serviços fixos (Fixed Revenues) a
   partir de um arquivo externo com layout de cenários por mês.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Processo_Extrair_Base_Fixed_Rev` (linha 12),
   `Atualizar_Lista_KPI_Versao_Interna`, `PopUp_Tempo_Processamento`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Fixed_Revenues".
11. **Abas:** `Sheet8` (`.Select`, linha 16).
12. **Pré-condições:** aba Extração com caminho "FIXED REVENUES" configurado.
13. **Passos principais:** cronometra; verifica versão; desliga UI; extrai; atualiza listas; popup; reativa UI.
14. **Pós-condições:** base Fixed Revenues atualizada na aba Base.
15. **Riscos:** **esta Sub NÃO chama `Processo_Limpar_Base_Fixed_Rev`, nem mesmo comentado** — mesmo padrão de
    risco de duplicação já visto em RGM e MOCKUP_RGM. Evidência: ausência de `Call
    Processo_Limpar_Base_Fixed_Rev` no corpo desta Sub.

### 46. Processo_Limpar_Base_Fixed_Rev

1. **Nome completo:** `Processo_Limpar_Base_Fixed_Rev`
2. **Módulo:** `Extracao_Fixed_Revenues.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Fixed_Revenues.bas:24`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_Fixed_Revenues" antes de nova extração.
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Processo_Limpar_Base_Fixed_Rev` no `grep`, nem
   comentada — não referenciado no código lido — provável botão/shape dedicado, isolado.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Fixed_Revenues".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** total desacoplamento do fluxo de extração. Evidência: linhas 26-37.

### 47. Processo_Extrair_Base_Fixed_Rev

1. **Nome completo:** `Processo_Extrair_Base_Fixed_Rev`
2. **Módulo:** `Extracao_Fixed_Revenues.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Fixed_Revenues.bas:42`
5. **Objetivo (negócio):** Abrir o arquivo externo de Fixed Revenues, montar um array de 14 "cenários"
   (KPI+Ano+Mês) a partir da configuração na aba Extração, delegar a extração linha-a-chave a
   `Processo_Extracao_Sheet_Base` (2 abas de chaves), aplicar escala (×1.000.000) e enriquecer com dimensões
   padrão.
6. **Quem chama (evidência):** `Extrair_Base_Fixed_Revenues` (linha 12); `Auxiliar.bas:21`
   (`Extrair_Todas_as_Bases`).
7. **Procedimentos chamados:** `Processo_Extracao_Sheet_Base` (própria privada — 2 chamadas: `Sheet19`/linha
   104, `Sheet20`/linha 112), `Form_Preenchimento_Generico` (7 chamadas: TIPO NÍVEL 2 143, FONTE 150, EXERCICIO
   157, KPI 164, VERSÃO 171, VISAO 178, PROFORMA 200), `Form_Organic` (183), `Form_Ref_Organic` (188),
   `Form_IFRS_Contabil` (193), `Form_Opex_Driven` (205), `Form_Grupo_BD` (210), `Form_Empresa` (215),
   `Form_Segmentos` (220, literal "FIXED REVENUES"), `Form_Classe` (225),
   `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (230), `Form_Acertar_Escala` (236, `Multiplicador=1000000`),
   `Form_Calcular_FY` (241), `Gerar_Visao_Italia` (246), `Reclassificar_Combinacoes_Empresas` (250),
   `_IFRS_Contabil` (251), `_Proforma` (252), `Calcular_Comb_Meses_Intervalo_Linha` (253).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Aux_2`=`Sheet2`, `Sh_Destino`=`Sheet3`;
    `Lista_Meses_Cenario(14)` (array com KPI+data formatada `AAAAMM`+mês abreviado, linhas 82-88); `KPI`,
    `Versao`, `ANO`.
11. **Abas/arquivos acessados:** arquivo externo Fixed Revenues (caminho via `Sheet8`); `Sheet19`, `Sheet20`
    (chaves De-Para); `Sheet3`.
12. **Pré-condições:** célula "FIXED REVENUES" em `Sheet8`; abas fonte esperadas dentro do arquivo externo.
13. **Passos principais:** localiza âncora "FIXED REVENUES" e lê parâmetros incluindo os 12 cenários de mês
    (linhas 59-88); abre arquivo externo (linha 93); chama `Processo_Extracao_Sheet_Base` duas vezes (linhas
    101-113); fecha arquivo (linhas 120-122); preenche campos fixos e roda enriquecimento (linhas 132-253).
14. **Pós-condições:** linhas da fonte "Base_Fixed_Revenues" na aba Base, com valores multiplicados por
    1.000.000.
15. **Riscos:** ausência de tratamento de erro na abertura do arquivo. Evidência: linhas 82-88, 101-113.

### 48. Processo_Extracao_Sheet_Base (variante `Extracao_Fixed_Revenues.bas`)

1. **Nome completo:** `Processo_Extracao_Sheet_Base` — implementação específica deste módulo (ver nota no item
   32).
2. **Módulo:** `Extracao_Fixed_Revenues.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Fixed_Revenues.bas:259`
5. **Objetivo (negócio):** Para cada linha de chave, injetar uma fórmula de concatenação Ano+Mês na planilha
   fonte para localizar dinamicamente a coluna do cenário correspondente, calcular/copiar o valor, e gravar
   Classe Custo (com descrição)/Centro Custo/Abertura_1/Segmento na Base.
6. **Quem chama (evidência):** `Processo_Extrair_Base_Fixed_Rev`, mesmo módulo (linhas 104 e 112, para
   `Sheet19` e `Sheet20`). `Private` — não referenciável fora do módulo.
7. **Procedimentos chamados:** `PEGA_A_DESCRICAO_DA_CONTA` (`TK_Functions.bas:2`, linha 388), `Ativar_Tudo`
   (`Auxiliar.bas`, linha 325, apenas no caminho de erro, antes do `End`).
8. **Parâmetros:** `Arq`, `Sh_Chaves`, `Sh_Destino`, `Linhas_excluidas`, `KPI`, `Versao`, `Lin_Destino`,
   `Mes_Ano_Inicial`, `Mes_Ano_Final`, `Col_Menu_Origem`, `Lin_Cabecalho_Origem`, `Lista_Meses_Cenario` (Variant
   implícito, `ByRef`).
9. **Retorno:** N/A (Sub).
10. **Variáveis relevantes:** `Campos_Chaves(10000)` (array fixo), `Ult_Col` (última coluna + 10, usada para
    injetar a fórmula de chave, linha 298).
11. **Abas acessadas:** `Sh_Chaves` (`Sheet19`/`Sheet20`); planilha fonte no arquivo externo (nome = valor de
    `Sh_Chaves.Cells(1,Col_Menu)`, sem sufixo de KPI diferente das outras duas variantes); `Sh_Destino`
    (`Sheet3`).
12. **Pré-condições:** aba fonte deve existir com o nome esperado; estrutura de linhas deve corresponder ao
    rótulo esperado.
13. **Passos principais:** injeta fórmula `=R[1]C&TEXT(R[2]C,"AAAAMM")` na linha 1 da planilha fonte, criando
    chaves de coluna combinando cenário+data (linha 299); valida rótulos linha a linha, com `MsgBox` crítico
    (linhas 305-329, **sem `End`** neste módulo — diferente das outras 2 variantes, inconsistência de
    tratamento entre módulos irmãos); para cada linha "INF"/"Fórmula" e cada um dos 14 cenários do array
    `Lista_Meses_Cenario`: localiza a coluna pela chave gerada, recalcula fórmula se aplicável, copia o valor se
    não-zero (linhas 344-406).
14. **Pós-condições:** linhas válidas copiadas para a Base, incluindo descrição da Classe de Custo.
15. **Efeitos colaterais/erro/riscos:**
    - **Escreve fórmula diretamente na planilha de origem** (arquivo de terceiros, ainda que aberto
      `ReadOnly:=True` — a alteração só afeta a cópia em memória, mas mistura transformação com o arquivo
      externo) — linha 299.
    - **Único dos 3 módulos irmãos que NÃO executa `End`** no caminho de divergência estrutural — inconsistência
      de comportamento entre RGM/MOCKUP_RGM (que abortam todo o Excel) e este módulo (que apenas avisa e seleciona
      a célula).
    - Loop duplo aninhado (linha de chave × 14 cenários) com `fn.Match`/`fn.CountIfs` a cada iteração — custo
      computacional crescente.
    - Array fixo `Campos_Chaves(10000)`.
    - Evidência: linhas 299, 305-329, 350-381.

---

## Módulo `Extracao_Sheet_Ajustes.bas`

### 49. Extrair_Base_Ajustes

1. **Nome completo:** `Extrair_Base_Ajustes`
2. **Módulo:** `Extracao_Sheet_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Extracao_Sheet_Ajustes.bas:4`
5. **Objetivo (negócio):** Ponto de entrada (botão) para extrair lançamentos de ajuste manual, informados
   diretamente pelos analistas em uma aba interna — única fonte do cluster 100% interna ao workbook, sem I/O
   de arquivo/rede/SQL.
6. **Quem chama:** não referenciado no código lido — provável botão/shape.
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica` (chamada **diretamente** com
   Campo="Fonte"/Chave="Base_Ajustes", linhas 10-12 — diferente do padrão dos outros 10 módulos, que delegam a
   limpeza a uma `Processo_Limpar_Base_X` própria), `Processo_Extrair_Base_Ajustes` (linha 13),
   `Atualizar_Lista_KPI_Versao` (linha 14 — **variante SEM "_Interna"**, diferente dos outros 10 módulos do
   cluster, que chamam `Atualizar_Lista_KPI_Versao_Interna`), `PopUp_Tempo_Processamento`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** N/A (Sub).
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Ajustes".
11. **Abas:** `Sheet8` (`.Select`, linha 16); `Sheet3` (Base, via `Limpar_Base_Historica`).
12. **Pré-condições:** aba de Ajustes (Sheet13) com lançamentos pendentes.
13. **Passos principais:** cronometra; desliga UI; limpa a fonte "Base_Ajustes" (SEMPRE, ao contrário das outras
    fontes do cluster — aqui a limpeza é incondicional); extrai; atualiza lista de KPI/Versão (variante não
    "Interna"); popup; reativa UI.
14. **Pós-condições:** base Ajustes atualizada na aba Base — sem risco de duplicação por reexecução, pois a
    limpeza é sempre executada.
15. **Riscos/inconsistências:** **é o único módulo do cluster cuja limpeza é incondicional e direta** (não
    delega à sua própria `Processo_Limpar_Base_Ajustes`) — padrão de implementação divergente dos demais 10
    módulos, o que dificulta entender a intenção original sem ler o código linha a linha. A chamada de
    `Atualizar_Lista_KPI_Versao` em vez de `Atualizar_Lista_KPI_Versao_Interna` (linha 14) pode ser bug
    (funcionalidade diferente por engano) ou distinção proposital não documentada —
    **[VALIDAR COM O NEGÓCIO]**. Evidência: linhas 10-14.

### 50. Processo_Limpar_Base_Ajustes

1. **Nome completo:** `Processo_Limpar_Base_Ajustes`
2. **Módulo:** `Extracao_Sheet_Ajustes.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Sheet_Ajustes.bas:24`
5. **Objetivo (negócio):** Remover registros históricos da fonte "Base_Ajustes" antes de nova extração, via
   botão dedicado (redundante com a limpeza incondicional já feita por `Extrair_Base_Ajustes`).
6. **Quem chama (evidência):** nenhuma ocorrência de `Call Processo_Limpar_Base_Ajustes` no `grep` — não
   referenciado no código lido — provável botão/shape dedicado, isolado (e redundante com o item 49).
7. **Procedimentos chamados:** `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído.
10. **Variáveis:** `Campo`="Fonte", `Chave`="Base_Ajustes".
11. **Abas:** `Sheet8` (`.Select`).
12. **Pré-condições:** nenhuma crítica.
13. **Passos principais:** desliga UI; limpa; reativa; `MsgBox` de sucesso.
14. **Pós-condições:** registros da fonte removidos.
15. **Riscos:** funcionalmente redundante com a limpeza já incondicional de `Extrair_Base_Ajustes` — mantém
    dois caminhos de código para o mesmo efeito. Evidência: linhas 26-37.

### 51. Processo_Extrair_Base_Ajustes

1. **Nome completo:** `Processo_Extrair_Base_Ajustes`
2. **Módulo:** `Extracao_Sheet_Ajustes.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Extracao_Sheet_Ajustes.bas:41`
5. **Objetivo (negócio):** Copiar da aba interna de Ajustes (Sheet13) os lançamentos manuais de correção
   (Versão, Exercício, Centro/Classe de Custo, valores mensais, Organic, Abertura_1, Segmento) para a Base,
   tratando vazios como zero e separando "KPI > Versão" quando informado nesse formato combinado.
6. **Quem chama (evidência):** `Extrair_Base_Ajustes` (linha 13); `Auxiliar.bas:17`
   (`Extrair_Todas_as_Bases`).
7. **Procedimentos chamados:** `Copiar_Base_Ajustes` (8 chamadas: VERSÃO 98, EXERCICIO 105, CENTRO CUSTO 112,
   CLASSE CUSTO 119, meses em loop 140, ORGANIC 198, ABERTURA_1 (via "A1_DESTINO") 238, SEGMENTO (via
   "SEG_N2_DESTINO") 246), `Form_Preenchimento_Generico` (5 chamadas: LIN_BASE 169, TIPO NÍVEL 2 176, FONTE 183,
   VISAO 190, PROFORMA 216), `Form_Ref_Organic` (204), `Form_IFRS_Contabil` (209), `Form_Opex_Driven` (221),
   `Form_Grupo_BD` (226), `Form_Empresa` (231), `Form_Segmentos` (252, literal "Ajustes"), `Form_Classe` (257),
   `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (262), `Form_Calcular_FY` (267), `Gerar_Visao_Italia` (272),
   `Reclassificar_Combinacoes_Empresas` (276), `_IFRS_Contabil` (277), `_Proforma` (278),
   `Calcular_Comb_Meses_Intervalo_Linha` (279). **Nota:** `Form_Organic` (versão calculada) **não é chamada** —
   ORGANIC vem direto da aba de origem (ver RN-021).
8. **Parâmetros:** nenhum. 9. **Retorno:** Function sem valor atribuído; possui `Exit Function` antecipado
   (linha 90) se `Ult_Lin_Origem <= Lin_Cabecalho_Origem` (nenhum ajuste a extrair).
10. **Variáveis:** `sh`=`Sheet8`, `Sh_Aux`=`Sheet15`, `Sh_Destino`=`Sheet3`, `Sh_Origem`=`Sheet13`.
11. **Abas acessadas:** `Sheet13` ("Ajustes", 100% interna); `Sheet3` (Base).
12. **Pré-condições:** `Sheet13` deve conter uma célula com o texto "Versão" (cabeçalho) e ao menos uma linha de
    dado após esse cabeçalho.
13. **Passos principais:**
    - Localiza cabeçalho "Versão" em `Sheet13` (linhas 68-76); se não houver linhas de dado, `MsgBox`
      informativo e `Exit Function` (linhas 86-92).
    - Copia Versão/Exercício/Centro Custo/Classe Custo (linhas 96-121).
    - Copia os 12 meses em loop, substituindo vazio por "0" via `Selection.Replace` após cada cópia (linhas
      124-145).
    - Separa "KPI > Versão" quando aplicável (linhas 155-163).
    - Preenche campos fixos, copia ORGANIC/ABERTURA_1/SEGMENTO diretamente da origem (linhas 165-252) e roda o
      restante do enriquecimento (linhas 254-279).
14. **Pós-condições:** linhas da fonte "Base_Ajustes" na aba Base, com valores mensais nunca vazios (sempre "0"
    ou um número).
15. **Efeitos colaterais/erro/regra de negócio/riscos:**
    - **RN-022:** células de valor mensal vazias tornam-se "0" explícito — trata ausência de lançamento como
      zero, não como omissão.
    - `Selection.Replace What:="", Replacement:="0"` dentro de um loop de 12 iterações que faz `Sh_Destino.Select`
      a cada passagem (linhas 142-143) — usa `Select`/`Selection` (anti-padrão VBA clássico), lento e frágil.
    - `MsgBox` informativo se não houver ajustes (linha 88) — ainda exige clique em OK, impedindo execução 100%
      desatendida mesmo neste caso simples.
    - Sem validação de tipo/formato dos valores inseridos manualmente pelo usuário em `Sheet13` antes de gravá-los
      na Base — o `Replace` só troca vazio por "0", não valida se o restante é numérico.
    - Evidência: linhas 86-92, 142-143.

### 52. Copiar_Base_Ajustes

1. **Nome completo:** `Copiar_Base_Ajustes`
2. **Módulo:** `Extracao_Sheet_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Extracao_Sheet_Ajustes.bas:284`
5. **Objetivo (negócio):** Rotina genérica de cópia de uma coluna nomeada da aba de Ajustes (Sheet13) para a
   Base, ancorada dinamicamente pela célula "*Versão*".
6. **Quem chama (evidência):** `Processo_Extrair_Base_Ajustes`, 8 chamadas (linhas 98, 105, 112, 119, 140, 198,
   238, 246). Não referenciada fora do módulo.
7. **Procedimentos chamados:** nenhum da lista indexada (`Cells.Find`, `Copy`/`PasteSpecial` nativos).
8. **Parâmetros:** `Arq` (não utilizado no corpo), `Sh_Origem`, `Col_Origem`, `Ult_Lin_Origem`, `Sh_Destino`,
   `Lin_Cabecalho`, `Lin_Cabecalho_Origem`, `Lin_Destino`, `Col_Destino`, `Col_Destino_2` (saída).
9. **Retorno:** N/A (Sub).
10. **Variáveis:** `fn` (`WorksheetFunction`), `Lin_Inicial` (linha seguinte à célula "*Versão*").
11. **Abas:** lê `Sh_Origem` (`Sheet13`); escreve `Sh_Destino` (`Sheet3`).
12. **Pré-condições:** célula "*Versão*" deve existir em `Sh_Origem`; coluna `Col_Origem` deve existir no
    cabeçalho correspondente.
13. **Passos principais:** localiza `Lin_Inicial` via `Cells.Find(What:="*Versão*")` + 1 (linhas 289-291);
    localiza colunas por `fn.Match`; copia via `Copy`/`PasteSpecial xlPasteValues` (linhas 296-297).
14. **Pós-condições:** coluna de destino preenchida.
15. **Riscos:** parâmetro `Arq` vestigial (não usado); `Cells.Find` sem checagem de `Is Nothing`. Evidência:
    linha 284 (assinatura), 289-291.

---

## Módulo `Gerar_Base_Pre_Closing.bas`

### 53. Gerar_Base_Versao_Pre_Closing

1. **Nome completo:** `Gerar_Base_Versao_Pre_Closing`
2. **Módulo:** `Gerar_Base_Pre_Closing.bas`
3. **Tipo:** Function (retorno não usado) — **não busca dados externos; reclassifica dados já presentes na
   Base**
4. **Escopo:** `Public` — `Gerar_Base_Pre_Closing.bas:3`
5. **Objetivo (negócio):** Gerar uma nova "versão" de Pré-Closing dentro da própria Base, aplicando, para cada
   coluna de cenário configurada na aba "Preview", os filtros de Empresa/IFRS/Proforma/Organic/período
   definidos pelo usuário, copiando os dados filtrados para uma nova versão rotulada — simula o fechamento
   contábil antes dos números oficiais existirem.
6. **Quem chama:** não referenciado no código lido — provável botão/shape ancorado na célula "AI39" da aba
   "Preview" (a própria função seleciona essa célula ao iniciar, linha 21, sugerindo que o botão fica próximo
   dela). **[NÃO ACESSÍVEL]** o `OnAction` real.
7. **Procedimentos chamados:** `Verifica_Versao`, `Desligar_Tudo`, `Filtrar_Item_Base` (3 chamadas: linhas 132,
   149, 214, mesmo módulo), `Limpar_Base_Historica` (condicional a `MsgBox`, linha 293, `Auxiliar.bas`),
   `Form_Calcular_FY` (312, `Aux_Formulas_Base.bas`), `Calcular_Comb_Meses_Intervalo("Fonte","Geral")` (313,
   `Auxiliar.bas`), `PopUp_Tempo_Processamento` (322), `Ativar_Tudo` (323).
8. **Parâmetros:** nenhum (opera sobre a seleção ativa e a aba "Preview").
9. **Retorno:** Function sem valor atribuído; possui `Exit Function` (linha 325) e rótulo de erro
   `Sheet_Inexistente` (linha 327, acionado se a aba "Preview" não existir).
10. **Variáveis relevantes:** `Sh_Preview` (=`Sheets("Preview")`), `Sh_Base`=`Sheet3`, `Cel_Select` (célula
    "AI39"), `Col_Cenario`, `Lin_Cabecalho`/`Lin_Cabecalho_Base`, `Range_Base`/`Range_Cabecalho_Base`,
    `Arq_Temp`/`Sh_Temp` (pasta de trabalho temporária criada via `Workbooks.Add`).
11. **Abas/intervalos/arquivos acessados:** aba "Preview" (matriz de cenários, coluna AI usada como gatilho
    inicial); `Sheet3` (Base, fonte e destino final); pasta de trabalho temporária (criada e fechada dentro da
    própria execução, sem persistência em disco).
12. **Pré-condições:** aba "Preview" deve existir (linha 10, `On Error GoTo Sheet_Inexistente`); usuário deve
    confirmar via `MsgBox` (linha 12) antes de prosseguir; matriz de cenários em "Preview" deve estar
    preenchida com os filtros desejados.
13. **Passos principais:**
    - Confirmação inicial via `MsgBox` (linha 12); se "Não", `Exit Function`.
    - Localiza cabeçalhos e limites de linha/coluna em "Preview" e na Base (linhas 40-56).
    - Cria pasta de trabalho temporária e copia o cabeçalho da Base (linhas 88-110).
    - Para cada coluna de cenário (loop linhas 114-187): remove filtros da Base; aplica filtros de
      Linha/Coluna via `Filtrar_Item_Base`; copia linhas visíveis para a pasta temporária; zera meses fora do
      intervalo do cenário; marca `Fonte = "QD AUTOMATIC - SELEÇÃO MANUAL"`.
    - Repete o mesmo mecanismo para o bloco de "Ajustes" vinculados à versão de Pré-Closing, marcando
      `Fonte = "QD AUTOMATIC - INPUT AJUSTE"` (linhas 190-244).
    - Atualiza campos Exercício/KPI/Versão na pasta temporária (linhas 255-286).
    - Pergunta via `MsgBox` (linha 290) se deve limpar a base histórica da versão gerada; se sim, chama
      `Limpar_Base_Historica`.
    - Cola todos os dados da pasta temporária de volta na Base e fecha a pasta temporária (linhas 298-305).
    - Recalcula FY e combinações de meses (linhas 309-313); reativa UI (linhas 317-323).
14. **Pós-condições:** nova(s) linha(s) na Base com `Fonte` = "QD AUTOMATIC - SELEÇÃO MANUAL" ou "QD AUTOMATIC -
    INPUT AJUSTE", representando a versão de Pré-Closing simulada.
15. **Efeitos colaterais/erro/regra de negócio/riscos:**
    - **RN-029:** duas categorias de linha geradas, rastreáveis pelo valor do campo Fonte.
    - **RN-010/RN-011:** versão default "Divulgado" e tratamento binário S/N de Empresa/Proforma/IFRS_Contábil
      (ver `Filtrar_Item_Base`, item 54).
    - Uso extensivo de `AutoFilter` + `SpecialCells(xlCellTypeVisible)` + `Copy`/`Paste` para simular um
      `WHERE` — lento em bases grandes, recalculado a cada cenário dentro do loop.
    - Cria e fecha uma pasta de trabalho inteira por execução (`Workbooks.Add`) — *overhead* de I/O/memória.
    - `MsgBox` de confirmação embutida no meio do fluxo (linhas 12, 290) — impede execução desatendida.
    - Coordenadas hardcoded (`Lin_1=3`, `Col_1=3`, célula "AI39", `Col_Cenario = ActiveCell.Column - 2`) tornam
      o código extremamente sensível a qualquer reformatação da aba "Preview".
    - Reaproveita a variável `Lin` fora do escopo do loop de cenários ao processar Ajustes (linha 214, usa `Lin`
      de um contexto de loop já encerrado) — indício de uso incorreto/compartilhado de variável entre blocos —
      **[VALIDAR COM O NEGÓCIO / TESTE TÉCNICO]** se produz resultado correto.
    - Evidência: linhas 12, 45-46, 88-90, 114-187, 214, 290, 298-305.

### 54. Filtrar_Item_Base

1. **Nome completo:** `Filtrar_Item_Base`
2. **Módulo:** `Gerar_Base_Pre_Closing.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Gerar_Base_Pre_Closing.bas:333`
5. **Objetivo (negócio):** Decidir, campo a campo, como montar o critério de `AutoFilter` a aplicar na Base
   para um determinado cenário — trata casos especiais (Proforma, IFRS_Contábil, Organic por faixa min/máx,
   Versão default) de forma centralizada.
6. **Quem chama (evidência):** `Gerar_Base_Versao_Pre_Closing`, 3 chamadas (linhas 132, 149, 214), mesmo
   módulo. `Private` — não referenciável fora do módulo.
7. **Procedimentos chamados:** nenhum da lista indexada (`AutoFilter`, `fn.Match` nativos).
8. **Parâmetros:** `fn`, `sh`, `Sh_Base`, `Lin`, `Lin_Cabecalho`, `Range_Cabecalho_Base`, `Range_Base`,
   `Campo_Filtr`, `Valor_Filtr` (também usado como saída/entrada — recalculado dentro da própria Sub),
   `Col_Menu_Selecao` (todos Variant implícito, `ByRef`).
9. **Retorno:** N/A (Sub); efeito é aplicar `AutoFilter` diretamente em `Sh_Base`.
10. **Variáveis relevantes:** `Col_Aux_1`/`Col_Aux_2` (faixa min/máx de Organic), `Interv` (índice de coluna do
    filtro).
11. **Abas/intervalos acessados:** `Sh_Base` (`Sheet3`, via `AutoFilter`); aba de origem `sh` (Preview, para ler
    `Aux Organic_1`/`Aux Organic_2`).
12. **Pré-condições:** `Range_Cabecalho_Base` deve conter o nome do campo a ser filtrado (senão `fn.Match`
    retorna erro não tratado).
13. **Passos principais:**
    - Campos "Aux Empresa"/"Aux IFRS Contabil"/"Aux Proforma"/"Aux Organic_1"/"Aux Organic_2" e qualquer campo
      contendo "Período" são **ignorados** (linhas 335-340).
    - Campos Empresa/Proforma/IFRS_Contábil: se houver valor, o **nome do campo de filtro passa a ser o próprio
      valor selecionado**, e o filtro aplicado é "S" (linhas 341-345) — RN-011.
    - Campo Organic: usa `Aux Organic_1`/`_2` como faixa min/máx e aplica `AutoFilter` com dois critérios
      (`>=`/`<=`) e operador `xlAnd` (linhas 346-357).
    - Se `Valor_Filtr` ficou vazio após os tratamentos acima: Proforma→"All effects Proforma"="S";
      IFRS_Contábil→"All effects IFRS"="S"; Versão→"Divulgado" (linhas 361-371) — RN-010.
    - Aplica `AutoFilter` final se `Campo_Filtr`/`Valor_Filtr` não vazios (linhas 373-376).
14. **Pós-condições:** `Sh_Base` filtrada de acordo com o cenário, pronta para a cópia das linhas visíveis pela
    chamadora.
15. **Riscos:** lógica condicional densa e pouco comentada — qualquer novo tipo de campo especial exige alterar
    esta Sub diretamente (regra de negócio embutida em código, não configurável). Sem tratamento de erro se
    `fn.Match` não encontrar o campo (`Interv`). Evidência: linhas 335-376.

---


## 10.2 Cluster Core / Motor de Cálculo (74 procedimentos — 3 módulos)
## A. Catálogo de Procedimentos

### A.1 — Módulo `Aux_Formulas_Base.bas` (20 procedimentos)

---

#### 1. `Atualizar_Formula_Manualmente`
1. **Nome completo**: `Atualizar_Formula_Manualmente`
2. **Módulo**: Aux_Formulas_Base.bas (linha 4)
3. **Tipo**: Sub
4. **Escopo**: Private
5. **Objetivo**: Utilitário de manutenção manual para recalcular a Diretoria Gerencial (com referência cruzada) apenas do bloco de linhas cuja Fonte = "Base_1009", sem rodar a extração completa. Serve para correção pontual pelo desenvolvedor.
6. **Quem chama**: Nenhuma chamada encontrada no dump (`grep` só retorna a própria declaração). Por ser `Private` e não estar ligada a nenhum botão identificado, é executada manualmente pelo desenvolvedor via VBE. `[NÃO ACESSÍVEL]` confirmar se há atalho/botão.
7. **Procedimentos chamados**: `Form_Diretoria_Gerencial_Com_Ref_Cruzada` (linha 17).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A (Sub).
10. **Variáveis/objetos relevantes**: `Campo_Chave="Fonte"`, `Chave="Base_1009"`, `Sh_Destino=Sheet3`, `fn=WorksheetFunction`.
11. **Abas/intervalos acessados**: `Sheet3` ("Base") — coluna "Fonte" para localizar `Lin_Inicial`/`Lin_Final` do bloco "Base_1009" via `Match`/`CountIfs`.
12. **Pré-condições**: A Base deve conter linhas com Fonte = "Base_1009" e cabeçalho identificável por "LIN_BASE".
13. **Passos principais**:
    - Localiza linha de cabeçalho via `Match("LIN_BASE", ...)`.
    - Localiza coluna "Fonte" e o intervalo de linhas com valor "Base_1009".
    - Chama `Form_Diretoria_Gerencial_Com_Ref_Cruzada` só nesse intervalo.
14. **Pós-condições**: Coluna(s) de Diretoria N1-N3 Gerencial recalculadas apenas para o bloco Base_1009.
15. **Efeitos colaterais/riscos/evidência**: Nome sugere uso ad-hoc/depuração ("Manualmente"); não integrado ao fluxo automatizado — risco de ficar desatualizado se a lógica de `Form_Diretoria_Gerencial_Com_Ref_Cruzada` mudar de assinatura. Evidência: linhas 4-19.

---

#### 2. `Form_Calcular_FY`
1. **Nome completo**: `Form_Calcular_FY`
2. **Módulo**: Aux_Formulas_Base.bas (linha 22)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito — sem `Private`)
5. **Objetivo**: Calcula a coluna "FY" (total anual) de cada linha como soma de Jan a Dez, e converte o resultado em valor fixo (não fórmula viva). É usado por praticamente todo módulo de extração para fechar o total anual de cada linha recém-inserida na Base.
6. **Quem chama**: `Extracao_Base_1009.bas:386`, `Auxiliar.bas:193`, `Extracao_Sheet_Ajustes.bas:267`, `Extracao_Base_Other_Inco.bas:361`, `Extracao_Fixed_Revenues.bas:241`, `Extracao_Base_Quick_Data.bas:298`, `Extracao_Base_Consolidad.bas:283`, `Extracao_Base_MOCKUP_RGM.bas:240`, `Form_Importacao.frm:646`, `Extracao_Base_RGM.bas:238`, `Gerar_Base_Pre_Closing.bas:312` — chamado por praticamente **todos** os módulos de extração e pela rotina de manutenção `Atualizar_Base_Todos_Campos_Auxiliares` (via `Auxiliar.bas:193`).
7. **Procedimentos chamados**: nenhum (usa apenas `WorksheetFunction` e `Range.FormulaR1C1`).
8. **Parâmetros**: `Sh_Destino` (Worksheet, ByRef implícito), `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho` (todos sem tipo declarado — Variant).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Col_INFO` (coluna "FY"), `Col_INFO_1` ("Jan"), `Col_INFO_2` ("Dez").
11. **Abas/intervalos acessados**: planilha passada por parâmetro (na prática sempre `Sheet3`/"Base"), colunas "FY", "Jan"..."Dez" localizadas por `Match` no cabeçalho.
12. **Pré-condições**: cabeçalho da planilha destino deve conter as colunas "FY", "Jan" e "Dez" nomeadas exatamente assim na linha de cabeçalho.
13. **Passos principais**:
    - Localiza colunas FY/Jan/Dez via `Match`.
    - Aplica `FormulaR1C1 = "=SUM(RC[Jan]:RC[Dez])"` no range Lin_Inicial:Lin_Final.
    - `.Copy` + `.PasteSpecial xlPasteValues` (bake-in) para congelar o resultado.
14. **Pós-condições**: coluna FY preenchida com valor numérico fixo (não fórmula) para o intervalo processado.
15. **Efeitos colaterais/riscos/evidência**: Padrão fórmula→calcular→copiar→colar repetido; nenhum tratamento de erro; se "FY"/"Jan"/"Dez" não existirem no cabeçalho, `Match` retorna erro não tratado (`#N/A` propagado, sem `On Error`). Evidência: linhas 22-37.

---

#### 3. `Form_Zerar_Meses_Exceto_Um`
1. **Nome completo**: `Form_Zerar_Meses_Exceto_Um`
2. **Módulo**: Aux_Formulas_Base.bas (linha 39)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Zera os valores de todos os meses de um intervalo de linhas, exceto o mês informado — usado para bases que representam eventos pontuais em um único mês (ex.: um ajuste ou uma base "n.a." que só tem valor em um mês específico).
6. **Quem chama**: `Extracao_Base_1009.bas:381`, `Extracao_Base_Consolidad.bas:160` (com `"n.a."` fixo), `Extracao_Base_Quick_Data.bas:293`, `Extracao_Base_Other_Inco.bas:351`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`, `Mes_Excessao` (mês que deve permanecer intacto).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: loop `For x = 1 To 12` mapeando índice para nome de mês abreviado em português (Jan..Dez) via cadeia de `If`.
11. **Abas/intervalos acessados**: planilha destino, colunas dos 12 meses.
12. **Pré-condições**: cabeçalho deve conter as 12 abreviações de mês em português.
13. **Passos principais**:
    - Para cada um dos 12 meses, se o mês não é o `Mes_Excessao`, zera (`= 0`) o intervalo de linhas naquela coluna.
14. **Pós-condições**: apenas a coluna do mês de exceção mantém valor original; demais ficam 0.
15. **Efeitos colaterais/riscos/evidência**: cadeia de 12 `If` sequenciais (não `Select Case`) — estilisticamente redundante mas funcional; nenhum tratamento de erro se `Mes_Excessao` não corresponder a nenhum mês válido (nesse caso todos os 12 meses são zerados silenciosamente). Evidência: linhas 39-65.

---

#### 4. `Form_Diretoria_Gerencial_Com_Ref_Cruzada`
1. **Nome completo**: `Form_Diretoria_Gerencial_Com_Ref_Cruzada`
2. **Módulo**: Aux_Formulas_Base.bas (linha 68)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Resolve a hierarquia de Diretoria Gerencial (Níveis N1, N2, N3) de cada linha da Base, tentando em cascata quatro fontes de verdade, da mais específica para a mais genérica. É a regra de negócio central para atribuir "dono" gerencial a cada linha de custo/receita.
6. **Quem chama**: `Aux_Formulas_Base.bas:17` (via `Atualizar_Formula_Manualmente`), `Extracao_Base_RGM.bas:227`, `Auxiliar.bas:183` (via `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_Base_Quick_Data.bas:288`, `Extracao_Base_Consolidad.bas:278`, `Extracao_Base_Other_Inco.bas:346`, `Extracao_Base_1009.bas:376`, `Extracao_Base_MOCKUP_RGM.bas:229`, `Extracao_Fixed_Revenues.bas:230`, `Extracao_Sheet_Ajustes.bas:262` — chamada por **todos** os módulos de extração.
7. **Procedimentos chamados**: nenhum diretamente (monta e aplica fórmula `FormulaR1C1`).
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15` ("Sup_Linhas"), `Sh_Ref_1=Sheet16`, `Sh_Ref_2=Sheet17`, `Chave_Resp_Grupo_BD` (concatenação de "Grupo BD 3..8"), fórmulas `Form_N1_CDC/Form_N2_CDC/Form_N3_CDC` (INDEX/MATCH por Centro de Custo em Sup_Linhas).
11. **Abas/intervalos acessados**: `Sheet3` (destino, colunas "Diretoria N1/N2/N3 Gerencial", "CLASSE CUSTO", "CENTRO CUSTO", "Grupo BD 3..8"), `Sheet15` ("Diretoria N1/N2/N3", "CENTROdeCUSTO"), `Sheet16` ("Chave_BD", "Classe Custo", "Diretoria N1..N3 Gerencial"), `Sheet17` ("Chave", "Diretoria_N1..N3_Gerencial_Destino").
12. **Pré-condições**: colunas "Grupo BD 3" e "Grupo BD 8" podem não existir em todas as bases (há checagem `CountIf`); cabeçalhos das 4 tabelas auxiliares devem estar consistentes.
13. **Passos principais**:
    - Monta 3 fórmulas INDEX/MATCH (`Form_N1_CDC`, `Form_N2_CDC`, `Form_N3_CDC`) que resolvem Diretoria N1-N3 a partir do Centro de Custo em Sup_Linhas.
    - Monta chave concatenada "Grupo BD 3&...&8" (pulando Grupo BD 3/8 se não existirem no cabeçalho da Base).
    - Para cada nível (N1, N2, N3): fórmula em cascata `IFERROR(INDEX/MATCH em Sh_Ref_2 por Classe Custo+N1+N2+N3, IFERROR(INDEX/MATCH em Sh_Ref_1 por Chave_Resp_Grupo_BD, IFERROR(INDEX/MATCH em Sh_Ref_1 por Classe Custo, INDEX/MATCH em Sup_Linhas por Centro de Custo)))`.
    - Aplica, calcula, `.Copy`+`PasteSpecial xlPasteValues` (bake-in).
14. **Pós-condições**: colunas "Diretoria N1/N2/N3 Gerencial" preenchidas com valor fixo resultante da cascata de fallback.
15. **Efeitos colaterais/riscos/evidência**: fórmula com até 4 níveis de `IFERROR` aninhado é cara de calcular em massa; existe bloco de código idêntico comentado logo acima (linhas 125-132) do bloco ativo (linhas 134-141) — cópia morta mantida como "backup inline", risco de editar a versão errada. Evidência: linhas 68-149, duplicação em 125-141.

---

#### 5. `Form_Classe`
1. **Nome completo**: `Form_Classe`
2. **Módulo**: Aux_Formulas_Base.bas (linha 152)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Classifica a coluna "CLASSE" de cada linha (categoria de custo), com uma regra especial de exceção para custo de mão-de-obra.
6. **Quem chama**: `Extracao_Base_Consolidad.bas:273`, `Extracao_Base_Quick_Data.bas:283`, `Extracao_Base_1009.bas:371`, `Extracao_Base_Other_Inco.bas:341`, `Auxiliar.bas:182` (via `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_Base_MOCKUP_RGM.bas:224`, `Extracao_Sheet_Ajustes.bas:257`, `Extracao_Fixed_Revenues.bas:225`, `Extracao_Base_RGM.bas:222`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15`; fórmula `=IF(Abertura_3="Labour Cost","w/o Fiber",INDEX(Sup_Linhas!CLASSE,MATCH(Centro Custo,Sup_Linhas!CENTROdeCUSTO,0)))`.
11. **Abas/intervalos acessados**: `Sheet3` (colunas "CLASSE", "CENTRO CUSTO", "Abertura_3"), `Sheet15` (colunas "CLASSE", "CENTROdeCUSTO").
12. **Pré-condições**: coluna "Abertura_3" deve já estar calculada (dependência implícita de ordem de execução com `Form_Opex_Driven`/`Form_Linha_BD`).
13. **Passos principais**:
    - Monta fórmula IF/INDEX/MATCH com exceção "Labour Cost"→"w/o Fiber".
    - Aplica, calcula, converte em valor.
14. **Pós-condições**: coluna "CLASSE" preenchida com valor fixo.
15. **Efeitos colaterais/riscos/evidência**: regra de exceção "Labour Cost"→"w/o Fiber" hardcoded sem comentário explicativo do porquê de negócio; se `Abertura_3` ainda não tiver sido calculada quando esta Sub roda, resultado fica incorreto (dependência de ordem não documentada em código, só na sequência de chamadas dos módulos de extração). Evidência: linhas 152-171.

---

#### 6. `Form_Empresa`
1. **Nome completo**: `Form_Empresa`
2. **Módulo**: Aux_Formulas_Base.bas (linha 174)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Resolve a coluna "EMPRESA" de cada linha via lookup direto no Centro de Custo em Sup_Linhas.
6. **Quem chama**: `Extracao_Base_Other_Inco.bas:331`, `Auxiliar.bas:180`, `Extracao_Base_MOCKUP_RGM.bas:214`, `Extracao_Base_1009.bas:361`, `Extracao_Base_Consolidad.bas:263`, `fx_IFRS16.bas:124` (pós-tratamento IFRS16, com `Sheet3, 6, qtdeLinhasBase+5, 5` fixo), `Extracao_Sheet_Ajustes.bas:231`, `Extracao_Fixed_Revenues.bas:215`, `Extracao_Base_Quick_Data.bas:278`, `Extracao_Base_RGM.bas:212`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15`; fórmula `=INDEX(Sup_Linhas!'EMPRESA 2',MATCH(Centro Custo,Sup_Linhas!CENTROdeCUSTO,0))`.
11. **Abas/intervalos acessados**: `Sheet3` ("EMPRESA", "CENTRO CUSTO"), `Sheet15` ("EMPRESA 2", "CENTROdeCUSTO").
12. **Pré-condições**: cada Centro de Custo da Base deve existir em Sup_Linhas (senão gera `#N/A` sem tratamento).
13. **Passos principais**: monta fórmula INDEX/MATCH, aplica, calcula, converte em valor.
14. **Pós-condições**: coluna "EMPRESA" preenchida com valor fixo.
15. **Efeitos colaterais/riscos/evidência**: sem `IFERROR` — diferente de `Form_Classe`/`Form_IFRS_Contabil`, aqui um Centro de Custo não encontrado propaga `#N/A` diretamente para a Base sem fallback. Evidência: linhas 174-192; uso pós-IFRS16 com linha/coluna hardcoded em `fx_IFRS16.bas:124`.

---

#### 7. `Form_Segmentos`
1. **Nome completo**: `Form_Segmentos`
2. **Módulo**: Aux_Formulas_Base.bas (linha 195)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: **Procedimento mais complexo do cluster** — calcula Abertura_1 e Segmento de cada linha, e implementa o **motor de rateio de custos indiretos**: desdobra linhas marcadas como "Rateio" em múltiplas linhas por combinação Empresa×Abertura_1 elegível, aplicando percentuais fixos ou variáveis, e gera uma linha de estorno/exclusão compensatória.
6. **Quem chama**: `Extracao_Sheet_Ajustes.bas:252` (`"Ajustes"`), `Extracao_Base_Other_Inco.bas:336` (`"Other Income"`), `Extracao_Fixed_Revenues.bas:220` (`"FIXED REVENUES"`), `Auxiliar.bas:181` (`"ALL BASES"`, via `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_Base_1009.bas:366` (`"1009"`), `Extracao_Base_Consolidad.bas:268` (`"BASE CONSOLIDADA"`). Chamadas **comentadas** (não ativas) em `Extracao_SQL_Hubble.bas:348` e `Extracao_Base_MOCKUP_RGM.bas:219`/`Extracao_Base_RGM.bas:217` — ou seja, a base Hubble, RGM e MOCKUP_RGM **não passam** por `Form_Segmentos` no fluxo atual (chamada desativada por comentário).
7. **Procedimentos chamados**: nenhum diretamente, mas usa `ActiveWorkbook.RefreshAll` (atualiza todas conexões/pivots) e manipula `SlicerCaches`.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`, `Sh_Origem` (string identificando a fonte, usada apenas para pular linhas de "Base_Ajustes" já preenchidas).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux_1=Sheet21` ("DP_Segmento"), `Sh_Aux_2=Sheet15` ("Sup_Linhas"), `Sh_Rateio=Sheet26`, `Lista_Empresa_A1` (lista de combinações Empresa>Abertura_1 elegíveis a rateio), `Formula_Padrao` (fórmula de rateio fixo/variável).
11. **Abas/intervalos acessados**: `Sheet3` (colunas Abertura_1, Segmento, Fonte, Linha_BD, Empresa, KPI, Versão, Exercício, FY, Jan-Dez), `Sheet21` (De-Para Segmento/Abertura_1 via VLOOKUP), `Sheet15` (fallback via chave `AA_Chave_Segmento`), `Sheet26` (colunas "Rateio Fixo", "Rateio Variavel", "Index A1", "Index_Lin_BD"), `SlicerCaches`: `Slicer_EMPRESA`, `Slicer_ABERTURA_1`, `Slicer_ABERTURA_2`, `Slicer_LINHA_BD`, `Slicer_IFRS_CONTABIL`.
12. **Pré-condições**: Empresa e Centro/Classe de Custo já calculados; Sheet26 (Rateio) atualizada e calculada (`Sh_Rateio.Calculate` chamado 2x); Slicers existentes no workbook com esses nomes exatos.
13. **Passos principais**:
    - Calcula fórmula de Abertura_1 (`VLOOKUP` em DP_Segmento, fallback INDEX/MATCH em Sup_Linhas, fallback fixo "Staff" para FIBER/INTELIG, senão "Total"), célula a célula, pulando linhas de "Base_Ajustes" já preenchidas.
    - Mesmo padrão para "Segmento" (fallback "Others").
    - Ajusta filtros de 5 Slicers para restringir o universo de rateio (Empresa=INTELIG ONGOING; exclui TIM FIXO/VOIP; exclui Revenues em Abertura_2 — comentado/desativado; Linha_BD restrita a 159-240 ou 265; exclui IFRS 9/15/16).
    - `ActiveWorkbook.RefreshAll`.
    - Recalcula `Sh_Rateio` 2x.
    - Monta `Lista_Empresa_A1` (combinações Empresa>Abertura_1 elegíveis, a partir de Sheet26 colunas "Index A1" e "Rateio Fixo").
    - Para cada linha com Abertura_1="Rateio": para cada combinação elegível da mesma Empresa, duplica a linha, seta o novo Abertura_1, recalcula Segmento, aplica fórmula de rateio (fixo via VLOOKUP simples, ou variável via VLOOKUP por KPI+Versão+Data+Empresa+A1+Linha_BD) mês a mês (com o valor original formatado como literal na fórmula).
    - Gera 1 linha de "EXCLUSAO" que estorna (multiplica por -1) o valor original, condicionada a pelo menos uma das linhas de rateio geradas ser diferente de zero, ou a existência de regra de rateio/revenue aplicável.
14. **Pós-condições**: linhas de rateio expandidas na Base (Fonte com sufixo "AJUSTE RATEIO VBA - RATEIO"/"- EXCLUSAO"); `Lin_Final` (parâmetro `ByRef` implícito) atualizado para refletir novo total de linhas; estado dos Slicers do workbook alterado (efeito colateral visível ao usuário).
15. **Efeitos colaterais/riscos/evidência**: (a) altera o estado de 5 Slicers do workbook, afetando a experiência do usuário mesmo fora desta rotina; (b) `ActiveWorkbook.RefreshAll` no meio da rotina é caro e de escopo maior que o necessário; (c) tratamento de erro por linha: se `VarType(vValor)=vbError` na coluna Abertura_1, emite `MsgBox` interativo "Atenção, erro encontrado na linha..." e pula a linha — **inviável em execução desatendida/batch**; chamada comentada a `fn_ListAllErrors` (linha 466) sugere que o log de erros estruturado foi implementado depois mas não reconectado aqui; (d) mensagens ao usuário: `MsgBox` de erro por linha (linha 467); (e) existe versão anterior mais simples (`Form_Segmentos_OLD`) sem a lógica de rateio, mantida como código morto no mesmo módulo. Evidência: linhas 195-552 (a Sub mais longa do módulo).

---

#### 8. `Form_Segmentos_OLD`
1. **Nome completo**: `Form_Segmentos_OLD`
2. **Módulo**: Aux_Formulas_Base.bas (linha 555)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Versão anterior/simplificada de `Form_Segmentos`, sem a lógica de rateio — calcula apenas Segmento e Abertura_1 via fórmula única aplicada ao range inteiro (sem desdobramento de linhas).
6. **Quem chama**: **Nenhuma chamada encontrada** — código morto mantido no módulo.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15`.
11. **Abas/intervalos acessados**: `Sheet3` ("Segmento", "Abertura_1", "CLASSE CUSTO", "CENTRO CUSTO", "EMPRESA"), `Sheet15` (chave `AA_Chave_Segmento`).
12. **Pré-condições**: N/A (não executado).
13. **Passos principais**: monta fórmula INDEX/MATCH com fallback fixo, aplica a todo o range de uma vez, calcula, converte em valor — para Segmento e depois para Abertura_1.
14. **Pós-condições**: N/A (código morto).
15. **Efeitos colaterais/riscos/evidência**: **Código morto confirmado por grep** — nenhuma chamada ativa em todo o dump. Risco de manutenção: um desenvolvedor pode confundir esta versão com a atual (`Form_Segmentos`) por semelhança de nome/estrutura e reativá-la por engano, perdendo a lógica de rateio. Recomenda-se remoção ou arquivamento explícito na reescrita. Evidência: linhas 555-622.

---

#### 9. `Form_Acertar_Sinal_Revenues`
1. **Nome completo**: `Form_Acertar_Sinal_Revenues`
2. **Módulo**: Aux_Formulas_Base.bas (linha 624)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Padroniza o sinal contábil das colunas de meses — inverte (multiplica por -1) os valores de linhas cuja Abertura_2 = "REVENUES", garantindo convenção de sinal consistente entre receita e despesa na Base.
6. **Quem chama**: `Extracao_Base_Other_Inco.bas:356` — **único chamador encontrado** (apenas o fluxo "Other Income" usa esta correção de sinal).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Destino`, `Ult_Lin_Destino`, `Lin_Cabecalho`. Nota: reatribui `Sh_Destino = Sheet3` internamente (linha 627), ignorando o parâmetro recebido — **inconsistência de assinatura** (o parâmetro é lido mas nunca usado após essa atribuição).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Col_Acerto` (coluna auxiliar temporária, "Dez"+1), loop pelos 12 meses.
11. **Abas/intervalos acessados**: `Sheet3` (fixo, via reatribuição) — colunas "ABERTURA_2", "Jan".."Dez", e coluna auxiliar temporária.
12. **Pré-condições**: existir uma coluna livre imediatamente após "Dez" para uso como área de trabalho temporária.
13. **Passos principais**: para cada mês, calcula em coluna auxiliar `=IF(Abertura_2="REVENUES", mês*-1, mês)`, copia como valor de volta para a coluna do mês original, depois limpa a coluna auxiliar.
14. **Pós-condições**: sinal dos valores de Revenues invertido; coluna auxiliar limpa ao final.
15. **Efeitos colaterais/riscos/evidência**: reatribuição de `Sh_Destino=Sheet3` internamente torna o parâmetro `Sh_Destino` enganoso (quem lê a assinatura pode achar que a Sub é genérica, mas está amarrada à planilha Base); usa uma coluna extra do grid como área de trabalho — se essa coluna estiver ocupada por outro dado, há risco de sobrescrita silenciosa. Evidência: linhas 624-649.

---

#### 10. `Form_Grupo_BD`
1. **Nome completo**: `Form_Grupo_BD`
2. **Módulo**: Aux_Formulas_Base.bas (linha 652)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Preenche as colunas "Grupo BD 2" a "Grupo BD 8" (hierarquia de agrupamento de plano de contas) via lookup por Classe Custo em uma tabela auxiliar de contas (Sheet2), com uma regra de exceção hardcoded para uma combinação específica de Classe Custo + prefixo de Centro de Custo.
6. **Quem chama**: `Extracao_Base_1009.bas:356`, `Extracao_Base_RGM.bas:207`, `Extracao_Base_MOCKUP_RGM.bas:209`, `Extracao_Base_Other_Inco.bas:326`, `Extracao_Base_Quick_Data.bas:273`, `Extracao_Base_Consolidad.bas:258`, `Extracao_Sheet_Ajustes.bas:226`, `Extracao_Fixed_Revenues.bas:210`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux_2=Sheet2`; loop `For x = 2 To 8` (Grupo BD 2..8), com checagem de existência de "Grupo BD 3"/"Grupo BD 8" no cabeçalho antes de processar (nem toda Base tem esses níveis).
11. **Abas/intervalos acessados**: `Sheet3` ("Grupo BD 2..8", "CLASSE CUSTO", "CENTRO CUSTO"), `Sheet2` ("Conta", "Descrição 2..8", "Descrição 9").
12. **Pré-condições**: coluna "Descrição 9" em Sheet2 deve conter o código "382" para a exceção especial funcionar.
13. **Passos principais**:
    - Para cada nível 2 a 8: monta fórmula `IFERROR(INDEX(Sheet2!Descrição_x, IF(Classe Custo="N203073156" AND Centro Custo começa com "NT", MATCH(382, Sheet2!Descrição_9), MATCH(Classe Custo, Sheet2!Conta))), "-")`.
    - Aplica, calcula, converte em valor.
14. **Pós-condições**: colunas "Grupo BD 2..8" preenchidas (ou "-" se não encontrado).
15. **Efeitos colaterais/riscos/evidência**: **regra de exceção hardcoded** (Classe Custo = "N203073156" + prefixo Centro de Custo "NT" → força busca pelo código 382) sem nenhum comentário explicando a razão de negócio — candidato a `[VALIDAR COM O NEGÓCIO]` na documentação funcional; mesma regra replicada de forma idêntica em `Form_Opex_Driven`/`Form_Linha_BD`, ou seja, a exceção está copiada 3 vezes no código-fonte (risco de divergência se corrigida em um lugar e esquecida nos outros). Evidência: linhas 652-690, mesma regra em Form_Opex_Driven (713-715) e Form_Linha_BD (743-745).

---

#### 11. `Form_Opex_Driven`
1. **Nome completo**: `Form_Opex_Driven`
2. **Módulo**: Aux_Formulas_Base.bas (linha 693)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Preenche as colunas "Abertura_2" a "Abertura_8" (árvore de classificação de P&L orientada a custo operacional) via lookup por Classe Custo em Sup_Linhas, com a mesma exceção hardcoded do código 382, e finaliza chamando `Form_Linha_BD`.
6. **Quem chama**: `Extracao_Base_MOCKUP_RGM.bas:204`, `Extracao_Base_Consolidad.bas:253`, `Extracao_Base_1009.bas:351`, `Auxiliar.bas:187` (via `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_Fixed_Revenues.bas:205`, `Extracao_Base_RGM.bas:202`, `Extracao_Base_Quick_Data.bas:268`, `Extracao_Base_Other_Inco.bas:321`, `Extracao_Sheet_Ajustes.bas:221`. Chamada **comentada** (desativada) em `Extracao_SQL_Hubble.bas:347` — a base Hubble não passa por esta rotina no fluxo atual (aparentemente resolve Abertura_2-8 por outro caminho, não coberto por este cluster).
7. **Procedimentos chamados**: `Form_Linha_BD` (linha 722, chamado ao final).
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15`; loop `For x = 2 To 8` (campos "Abertura_2".."Abertura_8"; nota: o `If x = 9` na linha 707 é código morto, pois o loop nunca chega a 9).
11. **Abas/intervalos acessados**: `Sheet3` ("Abertura_2..8", "CLASSE CUSTO", "CENTRO CUSTO"), `Sheet15` ("Abertura_2..8", "CLASSE CUSTO", "Linha_BD").
12. **Pré-condições**: mesma exceção 382 dependente de estrutura de Sup_Linhas.
13. **Passos principais**: mesma lógica de `Form_Grupo_BD`, mas usando Sup_Linhas em vez de Sheet2; ao final, chama `Form_Linha_BD`.
14. **Pós-condições**: colunas Abertura_2-8 e Linha_BD preenchidas.
15. **Efeitos colaterais/riscos/evidência**: linha 707 (`If x = 9 Then Campo = "Linha_BD"`) é inalcançável — o loop vai só até 8 — código morto residual dentro de uma Sub ativa (indício de refatoração incompleta: a lógica de Linha_BD foi extraída para uma Sub própria, mas o trecho antigo não foi limpo). Evidência: linhas 693-724.

---

#### 12. `Form_Linha_BD`
1. **Nome completo**: `Form_Linha_BD`
2. **Módulo**: Aux_Formulas_Base.bas (linha 727)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Preenche a coluna "Linha_BD" (código de linha do P&L) via lookup por Classe Custo em Sup_Linhas, com a mesma regra de exceção 382.
6. **Quem chama**: `Extracao_SQL_Hubble.bas:346` (chamada direta, fora do padrão `Form_Opex_Driven`), `Aux_Formulas_Base.bas:722` (interno, ao final de `Form_Opex_Driven`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet15`.
11. **Abas/intervalos acessados**: `Sheet3` ("Linha_BD", "CLASSE CUSTO", "CENTRO CUSTO"), `Sheet15` ("Linha_BD", "CLASSE CUSTO").
12. **Pré-condições**: idem `Form_Grupo_BD`.
13. **Passos principais**: monta fórmula INDEX/MATCH com a exceção 382, aplica, calcula, converte em valor.
14. **Pós-condições**: coluna "Linha_BD" preenchida.
15. **Efeitos colaterais/riscos/evidência**: terceira cópia da mesma regra de exceção 382 (ver risco documentado em `Form_Grupo_BD`); chamada de forma independente pela base Hubble (fora do padrão `Form_Opex_Driven`→`Form_Linha_BD` usado pelas demais bases) — indica que o pipeline de extração Hubble tem uma sequência de chamadas diferente das demais bases, risco de divergência de comportamento entre fontes. Evidência: linhas 727-751.

---

#### 13. `Form_IFRS_Contabil`
1. **Nome completo**: `Form_IFRS_Contabil`
2. **Módulo**: Aux_Formulas_Base.bas (linha 753)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Classifica cada linha quanto à norma contábil IFRS aplicável (IFRS 9/15/16 ou "w/o IFRS") via VLOOKUP em tabela de contas.
6. **Quem chama**: `Extracao_Base_1009.bas:339`, `Extracao_Base_MOCKUP_RGM.bas:192`, `Extracao_Base_Consolidad.bas:241`, `Extracao_Base_Other_Inco.bas:309`, `Extracao_Base_Quick_Data.bas:256`, `Extracao_Base_RGM.bas:190`, `Extracao_Sheet_Ajustes.bas:209`, `Extracao_Fixed_Revenues.bas:193`, `fx_IFRS16.bas:123` (pós-tratamento IFRS16, `Sheet3, 6, qtdeLinhasBase+5, 5` fixo).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh_Base=Sheet2`; fórmula `=IFERROR(VLOOKUP(Classe Custo, Sheet2!'Contas IFRS':+2, 3, 0), "w/o IFRS")`.
11. **Abas/intervalos acessados**: `Sheet3` ("IFRS_CONTABIL", "CLASSE CUSTO"), `Sheet2` ("Contas IFRS" e as 2 colunas seguintes).
12. **Pré-condições**: tabela "Contas IFRS" em Sheet2 deve ter 3 colunas contíguas (chave, intermediária, valor de retorno).
13. **Passos principais**: monta fórmula VLOOKUP com fallback "w/o IFRS", aplica, calcula, converte em valor.
14. **Pós-condições**: coluna "IFRS_CONTABIL" preenchida.
15. **Efeitos colaterais/riscos/evidência**: dependência posicional (`VLOOKUP` com índice de coluna fixo "3") — se a estrutura de Sheet2 mudar (inserir/remover coluna entre "Contas IFRS" e a coluna de retorno), a fórmula quebra silenciosamente (retorna outro dado, não necessariamente erro). Evidência: linhas 753-769.

---

#### 14. `Form_Organic`
1. **Nome completo**: `Form_Organic`
2. **Módulo**: Aux_Formulas_Base.bas (linha 771)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Marca todas as linhas do intervalo como "Ajusted" na coluna "ORGANIC" — atualmente uma constante fixa, não uma regra condicional (código anterior condicional está comentado).
6. **Quem chama**: `Extracao_Base_Consolidad.bas:231`, `Extracao_Base_Other_Inco.bas:299`, `Extracao_Base_1009.bas:329`, `Extracao_Base_MOCKUP_RGM.bas:182`, `Extracao_Fixed_Revenues.bas:183`, `Extracao_Base_RGM.bas:180`. **Não chamada** em `Extracao_Sheet_Ajustes.bas:195` (chamada comentada — base de Ajustes não recebe classificação Organic automática).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Formula = "=""Ajusted"""` (constante).
11. **Abas/intervalos acessados**: `Sheet3` (coluna "ORGANIC").
12. **Pré-condições**: nenhuma.
13. **Passos principais**: aplica constante "Ajusted" a toda a coluna do intervalo, calcula, converte em valor.
14. **Pós-condições**: coluna ORGANIC = "Ajusted" para todas as linhas processadas.
15. **Efeitos colaterais/riscos/evidência**: código comentado na linha 777 mostra que havia uma regra condicional anterior (`IF(Classe Custo="N904015189","Reported","Ajusted")`) que foi substituída por uma constante fixa — **mudança de regra de negócio sem explicação**, candidato a `[VALIDAR COM O NEGÓCIO]` (por que a condição foi removida? é proposital ou uma simplificação temporária esquecida?). Evidência: linhas 771-786.

---

#### 15. `Form_Ref_Organic`
1. **Nome completo**: `Form_Ref_Organic`
2. **Módulo**: Aux_Formulas_Base.bas (linha 788)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Atribui um peso numérico de 1 a 4 à coluna "REF_ORGANIC" a partir do valor de "ORGANIC" — provavelmente usado para ordenação/priorização em relatórios que distinguem visões "orgânicas" de "reportadas".
6. **Quem chama**: `Extracao_Base_1009.bas:334`, `Extracao_Fixed_Revenues.bas:188`, `Extracao_Base_Other_Inco.bas:304`, `Form_Importacao.frm:642`, `Extracao_Base_MOCKUP_RGM.bas:187`, `Extracao_Base_RGM.bas:185`, `Auxiliar.bas:188` (via `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_Base_Consolidad.bas:236`, `Extracao_Sheet_Ajustes.bas:204` (diferente de `Form_Organic`, esta é chamada mesmo para a base de Ajustes), `Extracao_Base_Quick_Data.bas:251`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: fórmula `=IF(ORGANIC="Reported",1,IF(ORGANIC="Recurrent",3,IF(ORGANIC="Normalized",4,2)))` — ou seja, "Ajusted" (e qualquer outro valor) cai no `ELSE` = 2.
11. **Abas/intervalos acessados**: `Sheet3` ("REF_ORGANIC", "ORGANIC").
12. **Pré-condições**: coluna "ORGANIC" já calculada (dependência de `Form_Organic` rodar antes).
13. **Passos principais**: aplica fórmula condicional, calcula, converte em valor.
14. **Pós-condições**: coluna REF_ORGANIC com peso numérico 1-4.
15. **Efeitos colaterais/riscos/evidência**: como `Form_Organic` hoje só produz "Ajusted" (peso 2 pelo `ELSE`), os valores 1/3/4 ("Reported"/"Recurrent"/"Normalized") parecem hoje inatingíveis pelo fluxo automático — possivelmente atingíveis apenas por dados inseridos manualmente (ex. Base_Ajustes) ou por versões anteriores do sistema; candidato a `[VALIDAR COM O NEGÓCIO]`. Evidência: linhas 788-802.

---

#### 16. `Form_Preenchimento_Generico`
1. **Nome completo**: `Form_Preenchimento_Generico`
2. **Módulo**: Aux_Formulas_Base.bas (linha 805)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Utilitário genérico que preenche uma coluna nomeada com um valor constante fixo em um intervalo de linhas — usado extensivamente pelos módulos de extração para setar metadados fixos (ex.: KPI, Versão, Fonte) em cada linha recém-extraída.
6. **Quem chama**: dezenas de chamadas em `Extracao_Base_Quick_Data.bas`, `Extracao_Sheet_Ajustes.bas`, `Extracao_Base_1009.bas`, `Extracao_Base_MOCKUP_RGM.bas`, `Extracao_Base_Consolidad.bas`, `Extracao_Base_Other_Inco.bas`, `Extracao_Base_RGM.bas`, `Extracao_Fixed_Revenues.bas` — é o utilitário de preenchimento de metadados **mais reutilizado do sistema** (>45 ocorrências de `Call`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Inicial`, `Lin_Final`, `Lin_Cabecalho`, `Campo` (nome da coluna), `Valor` (constante a atribuir).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma além dos parâmetros.
11. **Abas/intervalos acessados**: planilha destino (na prática sempre `Sheet3`), coluna localizada por `Match(Campo, cabeçalho)`.
12. **Pré-condições**: `Campo` deve existir no cabeçalho.
13. **Passos principais**: localiza coluna via `Match`, atribui `Valor` diretamente ao range (sem fórmula intermediária, atribuição direta de valor).
14. **Pós-condições**: coluna preenchida com valor constante.
15. **Efeitos colaterais/riscos/evidência**: Sub simples e de baixo risco individual, mas por ser chamada dezenas de vezes em sequência (uma vez por campo, por módulo de extração) dentro do mesmo range de linhas, representa uma oportunidade óbvia de otimização (poderia setar múltiplos campos em uma única passada) — não corrigir aqui, mas relevante para a reescrita em Python (vetorizar). Evidência: linhas 805-811.

---

#### 17. `Copiar_Base_Origem`
1. **Nome completo**: `Copiar_Base_Origem`
2. **Módulo**: Aux_Formulas_Base.bas (linha 814)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Copia uma coluna de uma planilha de origem para uma coluna nomeada na Base (por valor, não fórmula), com tratamento especial: se a coluna copiada for "CLASSE CUSTO" ou "CENTRO CUSTO", replica também os dados para o dashboard auxiliar "Painel_DM" e aciona as fórmulas de validação (`set_formula_CC`/`set_formula_CDC`/`set_formula_CDC_Parte_2`).
6. **Quem chama**: `Extracao_Base_1009.bas` (4x, linhas 233-263), `Extracao_Base_Other_Inco.bas:178`, `Extracao_Base_Quick_Data.bas` (8x, linhas 107-192) — usada pelos módulos de extração que copiam colunas de arquivos-fonte externos linha a linha por campo.
7. **Procedimentos chamados**: `set_formula_CC` (TK_Functions.bas, se Col_Destino="CLASSE CUSTO"), `set_formula_CDC` e `set_formula_CDC_Parte_2` (TK_Functions.bas, se Col_Destino="CENTRO CUSTO").
8. **Parâmetros**: `Arq`, `Sh_Origem`, `Lin_Cabecalho_Origem`, `Col_Origem`, `Ult_Lin_Origem`, `Sh_Destino`, `Lin_Cabecalho`, `Lin_Destino`, `Col_Destino`, `Col_Destino_2` (ByRef, recebe a coluna resolvida), `Sh_Destino_DM` (Optional, painel de validação).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma além dos parâmetros.
11. **Abas/intervalos acessados**: planilha de origem (arquivo externo aberto), `Sheet3` (destino), `Sh_Destino_DM` ("Painel_DM", condicional).
12. **Pré-condições**: `Col_Origem` deve existir no cabeçalho da planilha de origem na linha `Lin_Cabecalho_Origem`; se envolver CC/CDC, `Sh_Destino_DM` deve ser fornecido.
13. **Passos principais**:
    - Localiza colunas de origem/destino via `Match`.
    - Copia o intervalo de valores da origem para o destino (`.Copy` + `PasteSpecial xlPasteValues`).
    - Se Col_Destino="CLASSE CUSTO": também copia para `Sh_Destino_DM` (coluna 1) e chama `set_formula_CC`.
    - Se Col_Destino="CENTRO CUSTO": também copia para `Sh_Destino_DM` (coluna 4) e chama `set_formula_CDC` + `set_formula_CDC_Parte_2`.
14. **Pós-condições**: coluna na Base preenchida com valores copiados; opcionalmente, Painel_DM atualizado com flags de validação de CC/CDC.
15. **Efeitos colaterais/riscos/evidência**: acoplamento a nomes de coluna hardcoded como strings literais ("CLASSE CUSTO", "CENTRO CUSTO") comparados por igualdade exata — qualquer mudança de nomenclatura de coluna quebra silenciosamente o disparo do Painel_DM (sem erro, apenas deixa de popular o dashboard). Evidência: linhas 814-843.

---

#### 18. `Transformar_Texto_Mes_Em_Valor`
1. **Nome completo**: `Transformar_Texto_Mes_Em_Valor`
2. **Módulo**: Aux_Formulas_Base.bas (linha 846)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Normaliza colunas de meses de uma planilha de origem, convertendo textos como "-" em 0 e forçando tipo numérico, antes de copiar para a Base — limpeza de dado de entrada tipicamente vindo de exportações de outros sistemas com formatação de texto.
6. **Quem chama**: `Extracao_Base_1009.bas:132`, `Extracao_Base_Other_Inco.bas:114`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Origem`, `Lin_Cabecalho_Origem`, `Ult_Lin_Origem`, `Col_Inicial`, `Meses` (ByRef, reatribuído internamente para lista fixa), `Mes_Extracao_Inicial`, `Mes_Extracao_Final`, `Mes_Inicial` (ByRef, saída), `Mes_Final` (ByRef, saída), `Col_Mes_Origem_Inicial` (ByRef, saída), `Col_Mes_Origem_Final` (ByRef, saída).
9. **Retorno**: N/A (Sub, mas com múltiplos parâmetros de saída via ByRef).
10. **Variáveis/objetos relevantes**: `Col_Acerto` (coluna de trabalho temporária, 3 colunas após o fim do bloco de origem).
11. **Abas/intervalos acessados**: planilha de origem — colunas de mês localizadas por `Match(Mes & "*", cabeçalho)`.
12. **Pré-condições**: cabeçalho de origem deve conter os nomes dos meses como prefixo de coluna (usa wildcard `"*"`).
13. **Passos principais**: para cada um dos 12 meses, calcula em coluna auxiliar `=IF(TRIM(valor)="-",0,valor*1)`, copia como valor de volta, limpa a auxiliar; ao final resolve `Col_Mes_Origem_Inicial/Final` a partir dos meses de extração informados.
14. **Pós-condições**: colunas de mês da origem normalizadas para numérico; parâmetros de saída (`Mes_Inicial`, `Mes_Final`, `Col_Mes_Origem_Inicial/Final`) preenchidos para uso pelo chamador.
15. **Efeitos colaterais/riscos/evidência**: parâmetro `Meses` é sobrescrito internamente (linha 851) ignorando o que o chamador passou — nome de parâmetro enganoso (parece configurável, mas é sempre fixo); uso de coluna de trabalho fora da área de dados nomeada, risco de colisão com outras colunas. Evidência: linhas 846-875.

---

#### 19. `Apagar_Linhas_Zeradas`
1. **Nome completo**: `Apagar_Linhas_Zeradas`
2. **Módulo**: Aux_Formulas_Base.bas (linha 878)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Remove da planilha de origem as linhas cujos 12 meses são todos zero ou "-", reduzindo o volume de dados irrelevantes antes de incorporar à Base.
6. **Quem chama**: `Extracao_Base_1009.bas:214`, `Extracao_Base_Other_Inco.bas:119`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Origem`, `Lin_Cabecalho_Origem`, `Ult_Lin_Origem` (ByRef, atualizado se linhas forem removidas), `Col_Inicial`, `Col_Mes_Origem_Inicial`, `Col_Mes_Origem_Final`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Col_Acerto` (coluna auxiliar de marcação "DELETAR"/"-").
11. **Abas/intervalos acessados**: planilha de origem — colunas de mês e coluna auxiliar.
12. **Pré-condições**: `Col_Mes_Origem_Inicial/Final` já resolvidos (tipicamente vindos de `Transformar_Texto_Mes_Em_Valor`).
13. **Passos principais**: monta fórmula `=IF(AND(mês=0 OR mês="-", ...todos os 12 meses...), "DELETAR", "-")`; se houver ao menos 1 "DELETAR", ordena a planilha por essa coluna e deleta o bloco contíguo de linhas marcadas.
14. **Pós-condições**: linhas totalmente zeradas removidas; `Ult_Lin_Origem` atualizado.
15. **Efeitos colaterais/riscos/evidência**: reordena a planilha de origem (efeito colateral que pode afetar referências externas/manuais àquela planilha); usa o mesmo padrão de "ordenar + localizar bloco + `EntireRow.Delete`" visto em `Auxiliar.bas` — mesmo risco de custo O(n) por delete em massa. Evidência: linhas 878-922.

---

#### 20. `Form_Acertar_Escala`
1. **Nome completo**: `Form_Acertar_Escala`
2. **Módulo**: Aux_Formulas_Base.bas (linha 924)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Aplica um multiplicador de escala (ex.: milhares → unidades) a todas as colunas de meses de um intervalo de linhas.
6. **Quem chama**: `Extracao_Fixed_Revenues.bas:236`, `Extracao_Base_RGM.bas:233`, `Extracao_Base_MOCKUP_RGM.bas:235`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Sh_Destino`, `Lin_Cabecalho`, `Lin_Destino`, `Ult_Lin_Destino`, `Multiplicador`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: loop pelos 12 meses (mesma lista `";Jan;Fev;...;Dez"` usada em outras Subs do módulo).
11. **Abas/intervalos acessados**: planilha destino, colunas Jan-Dez.
12. **Pré-condições**: `Multiplicador` deve ser numérico válido.
13. **Passos principais**: para cada mês, coloca `Multiplicador` em uma célula auxiliar, copia, `PasteSpecial Operation:=xlMultiply` no intervalo de linhas, limpa a célula auxiliar.
14. **Pós-condições**: valores de todos os meses do intervalo multiplicados pelo fator de escala.
15. **Efeitos colaterais/riscos/evidência**: usa `Sh_Destino.Cells(2, Col_Acerto)` como célula de trabalho temporária (linha fixa 2) — se a linha 2 contiver dado relevante em alguma dessas bases, há risco de sobrescrita momentânea (ainda que revertida via `.ClearContents` ao final, uma falha no meio da execução deixaria a célula 2 poluída). Evidência: linhas 924-949.

---

### A.2 — Módulo `Auxiliar.bas` (28 procedimentos)

---

#### 21. `Extrair_Todas_as_Bases`
1. **Nome completo**: `Extrair_Todas_as_Bases`
2. **Módulo**: Auxiliar.bas (linha 4)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: **Macro-mestre de extração completa** — o botão "Extrair Tudo" do sistema. Orquestra a checagem de versão, limpeza de histórico "Geral", extração das 7 fontes de dados (Hubble, Consolidada, Ajustes, 1009, Other Income, RGM, Fixed Rev), atualização da lista de KPI/Versão e recálculo de combinações de meses.
6. **Quem chama**: **Nenhuma chamada `Call` encontrada no dump**. `[INFERÊNCIA]` — dado o nome e o papel de "botão principal", é presumivelmente acionada por um Shape/botão na aba "Extracao" (Sheet8) via `Shape.OnAction`, não capturável em texto VBA. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Verifica_Versao` (linha 8 — nota: escrito `cal Verifica_Versao`, **typo de `Call`**, mas funciona pois `Call` é opcional em VBA), `Desligar_Tudo`, `Limpar_Base_Historica`, `Processo_Extrair_Base_Hubble`, `Processo_Extrair_Base_Consolidada`, `Processo_Extrair_Base_Ajustes`, `Processo_Extrair_Base_1009`, `Processo_Extrair_Base_Other_Income`, `Processo_Extrair_Base_RGM`, `Processo_Extrair_Base_Fixed_Rev`, `Atualizar_Lista_KPI_Versao_Interna`, `Calcular_Comb_Meses_Intervalo`, `PopUp_Tempo_Processamento`, `Ativar_Tudo` (todos fora deste cluster, exceto os citados que pertencem a Auxiliar.bas).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Inicio`/`Termino` (globais, cronometragem), `Campo`/`Valor` = "Fonte"/"Geral".
11. **Abas/intervalos acessados**: `Sheet3` (indiretamente, via as rotinas chamadas), `Sheet8` (`.Select` ao final).
12. **Pré-condições**: conexão de rede/SQL Server disponível (para `Verifica_Versao` e as extrações SQL); arquivos-fonte externos acessíveis (para as extrações baseadas em arquivo).
13. **Passos principais**:
    - Marca `Inicio = Now`.
    - `Verifica_Versao` (bloqueia se desatualizado).
    - `Desligar_Tudo` (ScreenUpdating/Calculation/Events off).
    - Limpa histórico "Geral" (todas as fontes).
    - Chama sequencialmente as 7 rotinas de extração por fonte.
    - Atualiza lista de KPI/Versão interna.
    - Recalcula combinações de meses para toda a base ("Geral").
    - Seleciona Sheet8, mostra popup de tempo, religa otimizações.
14. **Pós-condições**: Base totalmente reconstruída a partir de todas as fontes; popup de conclusão exibido ao usuário com tempo decorrido.
15. **Efeitos colaterais/riscos/evidência**: **typo `cal` em vez de `Call`** na linha 8 — funciona por acidente (VBA aceita `NomeDaSub argumentos` sem `Call`), mas é um sinal de falta de revisão de código; execução sequencial e monolítica de 7 extrações + reclassificações — qualquer falha no meio interrompe o processo com o Excel possivelmente deixado em estado "desligado" (ScreenUpdating=False etc.) se o erro não for tratado, exigindo reinício manual do Excel; nenhuma barreira de erro (`On Error`) ao redor das chamadas em cadeia. Evidência: linhas 4-29.

---

#### 22. `Limpar_Todas_as_Bases`
1. **Nome completo**: `Limpar_Todas_as_Bases`
2. **Módulo**: Auxiliar.bas (linha 31)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Botão de "reset" — apaga todo o histórico da Base (`Tipo="Geral"`) e força atualização de todas as conexões/pivots do workbook, deixando o sistema pronto para nova extração do zero.
6. **Quem chama**: **Nenhuma chamada `Call` encontrada.** `[INFERÊNCIA]` botão na aba Extracao. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Desligar_Tudo`, `Limpar_Base_Historica`, `Ativar_Tudo`.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Campo_Pesq="Fonte"`, `Tipo_Limpeza="Geral"`.
11. **Abas/intervalos acessados**: `Sheet3` (limpeza completa), `Sheet8` (`.Select`).
12. **Pré-condições**: nenhuma especial.
13. **Passos principais**: desliga otimizações; limpa toda a Base; seleciona Sheet8; ativa workbook; `RefreshAll`; religa otimizações; `MsgBox` de sucesso.
14. **Pós-condições**: Base vazia (apenas cabeçalho); todas as conexões/pivots atualizadas (potencialmente contra dados agora vazios).
15. **Efeitos colaterais/riscos/evidência**: **ação destrutiva e irreversível** sem confirmação `MsgBox vbYesNo` prévia (diferente de outras rotinas destrutivas do sistema, como `RemoveStyles` ou `UPDATE_Combinacoes_Empresas`, que pedem confirmação) — um clique acidental apaga toda a base histórica sem chance de cancelar. Mensagem final: `"Limpeza concluída com sucesso !!!"` (linha 45). Evidência: linhas 31-48.

---

#### 23. `Exportar`
1. **Nome completo**: `Exportar`
2. **Módulo**: Auxiliar.bas (linha 50)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Abre o formulário `Form_Exportacao` (fora deste cluster) para o usuário configurar e disparar exportação de Front ou de Base.
6. **Quem chama**: Nenhuma chamada `Call` direta encontrada; `Form_Exportacao.frm` contém apenas referências textuais a "Exportar" em comentários/captions de UI (linhas 77, 104, 185, 509), não uma chamada de procedimento. `[INFERÊNCIA]` botão na aba Extracao. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Form_Exportacao.Show` (UserForm, fora do cluster).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: nenhuma diretamente (delega ao formulário).
12. **Pré-condições**: nenhuma.
13. **Passos principais**: `Form_Exportacao.Show`.
14. **Pós-condições**: formulário modal exibido; lógica de exportação real vive em `Form_Exportacao.frm` (fora deste cluster).
15. **Efeitos colaterais/riscos/evidência**: Sub trivial de uma linha; risco baixo isoladamente. Evidência: linhas 50-54.

---

#### 24. `Importar_Fronts`
1. **Nome completo**: `Importar_Fronts`
2. **Módulo**: Auxiliar.bas (linha 56)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Valida a versão do sistema e abre o formulário `Form_Importacao` para o usuário importar arquivos "Front" (layouts de entrada de dados).
6. **Quem chama**: **Nenhuma chamada `Call` encontrada.** `[INFERÊNCIA]` botão na aba Extracao. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Verifica_Versao`, `Form_Importacao.Show` (UserForm, fora do cluster).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: nenhuma diretamente.
12. **Pré-condições**: conexão SQL disponível para `Verifica_Versao`.
13. **Passos principais**: `Verifica_Versao`; `Form_Importacao.Show`.
14. **Pós-condições**: formulário modal exibido (lógica real em `Form_Importacao.frm`, fora do cluster).
15. **Efeitos colaterais/riscos/evidência**: se `Verifica_Versao` bloquear (versão desatualizada), a Sub encerra via `End` dentro de `Verifica_Versao` — comportamento abrupto herdado. Evidência: linhas 56-61.

---

#### 25. `Calcular_Range_Selecionado`
1. **Nome completo**: `Calcular_Range_Selecionado`
2. **Módulo**: Auxiliar.bas (linha 63)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Utilitário de conveniência que força o recálculo apenas da seleção atual do usuário (`Selection.Calculate`) — provavelmente um atalho/botão de "recalcular isto" para quando o workbook está em modo de cálculo manual.
6. **Quem chama**: **Nenhuma chamada encontrada.** `[INFERÊNCIA]` atalho de teclado ou botão de menu de contexto. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Selection` (objeto implícito do Excel).
11. **Abas/intervalos acessados**: qualquer range que esteja selecionado no momento da chamada.
12. **Pré-condições**: usuário deve ter uma seleção de células ativa.
13. **Passos principais**: `Selection.Calculate`.
14. **Pós-condições**: fórmulas na seleção recalculadas.
15. **Efeitos colaterais/riscos/evidência**: Sub trivial, risco mínimo. Evidência: linhas 63-67.

---

#### 26. `Ativar_Tudo`
1. **Nome completo**: `Ativar_Tudo`
2. **Módulo**: Auxiliar.bas (linha 69)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Contraparte de `Desligar_Tudo` — religa `ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, limpa a `StatusBar`, e força 2 recálculos completos de `Sheet15` antes de devolver o controle ao usuário. É chamada ao final de praticamente toda rotina pesada do sistema.
6. **Quem chama**: >35 ocorrências em todo o dump — `Auxiliar.bas` (linhas 27, 43, 196), `Extracao_Base_MOCKUP_RGM.bas`, `BackupCodigo_MainResults.bas`, `Extracao_Base_Consolidad.bas`, `Extracao_Base_1009.bas`, `Form_Importacao.frm`, `Limpeza_Base_Ajustes.bas`, `Front_Processos.bas`, `Form_Exportacao.frm`, `Gerar_Base_Pre_Closing.bas`, `Extracao_SQL_Hubble.bas`, `Extracao_Sheet_Ajustes.bas`, `Refresh_De_X_Para.bas`, `Extracao_Fixed_Revenues.bas`, `Extracao_Base_RGM.bas`, `Extracao_Base_Other_Inco.bas`, `Extracao_Base_Quick_Data.bas`, `Refresh_Drop_Comb.bas`, `TK_Functions.bas:572` (dentro de `UPDATE_aplicar_CDC_por_Referencia`) — é, junto com `Desligar_Tudo`, o par de rotinas mais amplamente reutilizado de todo o sistema.
7. **Procedimentos chamados**: nenhum (usa apenas propriedades `Application.*` e `Sheet15.Calculate`).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: loop `For x = 1 To 2` (`Sheet15.Calculate` duas vezes).
11. **Abas/intervalos acessados**: `Sheet15` ("Sup_Linhas") — recalculada explicitamente 2x.
12. **Pré-condições**: nenhuma.
13. **Passos principais**: recalcula Sheet15 duas vezes; liga `ScreenUpdating`, `DisplayAlerts`, `EnableEvents`; limpa `StatusBar`.
14. **Pós-condições**: Excel volta ao modo interativo normal.
15. **Efeitos colaterais/riscos/evidência**: **não religa `Application.Calculation` para automático** — apenas `Desligar_Tudo` seta `xlCalculationManual`, mas `Ativar_Tudo` não faz o `xlCalculationAutomatic` de volta; isso significa que, a menos que outra rotina explicitamente restaure o cálculo automático, o workbook pode permanecer em cálculo manual após uma extração — risco relevante de UX (usuário edita a planilha depois e nada recalcula sozinho). Por que recalcular `Sheet15` especificamente 2 vezes (e não outras abas, nem 1 ou 3 vezes) não está documentado — `[VALIDAR COM O NEGÓCIO]`. Evidência: linhas 69-80.

---

#### 27. `Verifica_Versao`
1. **Nome completo**: `Verifica_Versao`
2. **Módulo**: Auxiliar.bas (linha 82)
3. **Tipo**: Function (usada como Sub, sem uso do valor de retorno)
4. **Escopo**: Public
5. **Objetivo**: **Gate de controle de versão corporativo** — consulta o SQL Server para confirmar que a versão local do Quick Data (hardcoded "3.0" no código) corresponde à versão oficial cadastrada; se não corresponder, bloqueia todo o uso do sistema.
6. **Quem chama**: `Extracao_Base_Other_Inco.bas:10`, `Extracao_Base_MOCKUP_RGM.bas:10`, `Extracao_Base_Consolidad.bas:10`, `Extracao_Base_1009.bas:9`, `Gerar_Base_Pre_Closing.bas:31`, `Extracao_SQL_Hubble.bas:7`, `Auxiliar.bas:8` (com typo `cal`, dentro de `Extrair_Todas_as_Bases`), `Auxiliar.bas:58` (dentro de `Importar_Fronts`), `Auxiliar.bas:152` (dentro de `Atualizar_Base_Todos_Campos_Auxiliares`), `Auxiliar.bas:569` (dentro de `Refresh_Base_Aux`), `Extracao_Fixed_Revenues.bas:10`, `Extracao_Base_RGM.bas:10`, `Extracao_Base_Quick_Data.bas:10` — chamada como "guarda de entrada" por praticamente todos os fluxos principais do sistema.
7. **Procedimentos chamados**: `AbreConexao`, `FechaConexao` (Conexoes.bas, fora do cluster).
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem atribuição de valor de retorno em nenhum caminho do código (nome da função nunca é atribuído) — tecnicamente retorna `Empty`/`Nothing`; na prática usada apenas pelo efeito colateral (bloquear ou não a execução).
10. **Variáveis/objetos relevantes**: `Info` (ADODB.Recordset), `Comando` (ADODB.Command), `conn` (global de `Conexoes.bas`).
11. **Abas/intervalos acessados**: nenhuma planilha — consulta SQL Server, tabela `[BPAM].[dbo].[TB_HUBBLE_VERSAO_FERRAMENTAS]`, filtro `FERRAMENTA = 'QUICK DATA'`.
12. **Pré-condições**: conectividade de rede com o servidor SQL Server BPAM.
13. **Passos principais**:
    - Abre conexão (`AbreConexao`).
    - Executa `SELECT VERSAO FROM TB_HUBBLE_VERSAO_FERRAMENTAS WHERE FERRAMENTA='QUICK DATA'`.
    - Compara `Info!Versao` com o literal `"3.0"` hardcoded no código.
    - Se diferente: `MsgBox` crítico orientando contato com administradores, fecha conexão, `End` (encerra toda a execução VBA abruptamente).
    - Se igual: fecha recordset e conexão normalmente.
    - Bloco `tratar_erro`: em caso de exceção, `MsgBox` crítico com instrução de contato por e-mail (`dfigsilva@timbrasil.com.br`) incluindo usuário do Windows (`Environ("UserName")`).
14. **Pós-condições**: execução prossegue normalmente (versão OK) ou é abortada (`End`) com toda a pilha de chamadas VBA interrompida.
15. **Efeitos colaterais/riscos/evidência**: **`End` é uma interrupção abrupta de todo o processo VBA** — não fecha graciosamente recursos abertos por chamadores anteriores na pilha (ex.: se outra rotina já tinha um Workbook aberto ou uma transação em andamento, `End` não executa nenhum código de limpeza posterior); a versão de referência "3.0" está **hardcoded no código-fonte**, exigindo recompilação/redistribuição do arquivo toda vez que a versão mudar — mecanismo de controle de versão frágil comparado a, por exemplo, ler a versão local de uma célula nomeada; mensagem de erro de rede expõe e-mail de contato pessoal (`dfigsilva@timbrasil.com.br`) hardcoded — indica proprietário/mantenedor original do sistema. Evidência: linhas 82-117.

---

#### 28. `Reexibir_Sheets`
1. **Nome completo**: `Reexibir_Sheets`
2. **Módulo**: Auxiliar.bas (linha 119)
3. **Tipo**: Sub
4. **Escopo**: Private
5. **Objetivo**: Torna visíveis todas as planilhas do workbook (utilitário de depuração/manutenção, já que muitas abas de suporte ficam ocultas por padrão para o usuário final).
6. **Quem chama**: **Nenhuma chamada encontrada** — Private e sem chamador no dump. Executada manualmente via VBE. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: loop `For x = 1 To ThisWorkbook.Worksheets.Count`.
11. **Abas/intervalos acessados**: todas as planilhas do workbook (`.Visible = True` para cada uma).
12. **Pré-condições**: nenhuma.
13. **Passos principais**: itera todas as planilhas e força `Visible = True`.
14. **Pós-condições**: nenhuma planilha permanece oculta.
15. **Efeitos colaterais/riscos/evidência**: código morto/utilitário de depuração sem chamador ativo — candidato a remoção ou, alternativamente, indício de que deveria estar acessível ao usuário mas não está ligada a nenhum botão. Evidência: linhas 119-125.

---

#### 29. `Desligar_Tudo`
1. **Nome completo**: `Desligar_Tudo`
2. **Módulo**: Auxiliar.bas (linha 128)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Desliga as otimizações de performance do Excel (`ScreenUpdating`, `DisplayAlerts`, `EnableEvents`, `Calculation=Manual`) antes de qualquer rotina pesada — par de `Ativar_Tudo`.
6. **Quem chama**: >30 ocorrências em todo o dump, mesmo padrão de reutilização de `Ativar_Tudo` — `Auxiliar.bas` (linhas 10, 33, 155), praticamente todos os módulos `Extracao_Base_*.bas`, `Form_Exportacao.frm`, `Form_Importacao.frm`, `Front_Processos.bas`, `Limpeza_Base_Ajustes.bas`, `Gerar_Base_Pre_Closing.bas`, `Refresh_De_X_Para.bas`, `Refresh_Drop_Comb.bas`, `BackupCodigo_MainResults.bas`, `TK_Functions.bas:532` (dentro de `UPDATE_aplicar_CDC_por_Referencia`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno atribuído (mesmo padrão de `Verifica_Versao` — usada apenas pelo efeito colateral).
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: nenhuma.
12. **Pré-condições**: nenhuma.
13. **Passos principais**: seta `ScreenUpdating=False`, `DisplayAlerts=False`, `EnableEvents=False`, `Calculation=xlCalculationManual`.
14. **Pós-condições**: Excel em modo "silencioso"/não-interativo, cálculo manual.
15. **Efeitos colaterais/riscos/evidência**: se a rotina que chama `Desligar_Tudo` falhar com erro não tratado antes de chamar `Ativar_Tudo`, o Excel fica "travado" nesse estado (sem atualização de tela, sem alertas, sem eventos, cálculo manual) até o usuário fechar e reabrir o arquivo — padrão de risco sistêmico presente em todo o sistema, não só neste cluster. Evidência: linhas 128-135.

---

#### 30. `PopUp_Tempo_Processamento`
1. **Nome completo**: `PopUp_Tempo_Processamento`
2. **Módulo**: Auxiliar.bas (linha 138)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Exibe ao usuário um `MsgBox` de conclusão com o tempo total de processamento (usando os globais `Inicio`/`Termino`), ao final de rotinas longas.
6. **Quem chama**: `Form_Importacao.frm` (2x), `Gerar_Base_Pre_Closing.bas:322`, `Auxiliar.bas:26` (dentro de `Extrair_Todas_as_Bases`), `Auxiliar.bas:197` (dentro de `Atualizar_Base_Todos_Campos_Auxiliares`), `Extracao_SQL_Hubble.bas:17`, `Front_Processos.bas:87`, `Extracao_Base_Consolidad.bas:18`, `Extracao_Base_1009.bas:17`, `Extracao_Sheet_Ajustes.bas:18`, `Extracao_Base_MOCKUP_RGM.bas:17`, `Extracao_Base_Quick_Data.bas:18`, `Extracao_Base_Other_Inco.bas:18`, `Extracao_Base_RGM.bas:17`, `Extracao_Fixed_Revenues.bas:18`, `TK_Functions.bas:571` (dentro de `UPDATE_aplicar_CDC_por_Referencia`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum (usa globais `Inicio`/`Termino` declarados em `Auxiliar.bas:2`).
9. **Retorno**: Function sem valor de retorno atribuído.
10. **Variáveis/objetos relevantes**: `Inicio`, `Termino` (globais).
11. **Abas/intervalos acessados**: nenhuma.
12. **Pré-condições**: `Inicio` deve ter sido setado (`TimeValue(Now)`) pelo chamador antes de iniciar o processamento; caso contrário o tempo calculado é incorreto/sem sentido.
13. **Passos principais**: `On Error Resume Next`; marca `Termino=Now`; `MsgBox` com tempo formatado `hh:mm:ss` e horários de início/fim.
14. **Pós-condições**: popup exibido ao usuário.
15. **Efeitos colaterais/riscos/evidência**: `On Error Resume Next` silencioso — qualquer erro na formatação do tempo é engolido sem log; dependência de estado global (`Inicio`) compartilhado entre múltiplas rotinas que podem rodar em sequência ou aninhadas, criando risco de o tempo reportado estar incorreto se duas rotinas cronometradas se sobrepuserem. Evidência: linhas 138-147.

---

#### 31. `Atualizar_Base_Todos_Campos_Auxiliares`
1. **Nome completo**: `Atualizar_Base_Todos_Campos_Auxiliares`
2. **Módulo**: Auxiliar.bas (linha 150)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: **Botão de manutenção em massa** — recalcula, para toda a Base já extraída, os campos Empresa/Diretorias Gerenciais/Segmento/Classe (mediante confirmação do usuário, pois exige reextração da base Hubble depois) e sempre recalcula Opex_Driven, Ref_Organic, as 3 reclassificações de combinação e o FY.
6. **Quem chama**: **Nenhuma chamada encontrada.** `[INFERÊNCIA]` botão na aba Extracao. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Verifica_Versao`, `Desligar_Tudo`, `Limpar_Base_Historica`, `Form_Empresa`, `Form_Segmentos` (com `"ALL BASES"`), `Form_Classe`, `Form_Diretoria_Gerencial_Com_Ref_Cruzada`, `Form_Opex_Driven`, `Form_Ref_Organic`, `Reclassificar_Combinacoes_Empresas`, `Reclassificar_Combinacoes_IFRS_Contabil`, `Reclassificar_Combinacoes_Proforma`, `Calcular_Comb_Meses_Intervalo`, `Form_Calcular_FY`, `Ativar_Tudo`, `PopUp_Tempo_Processamento`.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Lin_Cabecalho`, `Lin_Inicial`, `Lin_Final` (localizados dinamicamente via "LIN_BASE"); `Campo`/`Chave` = "Fonte"/"Geral" (ou "Base_Hubble" no caminho condicional).
11. **Abas/intervalos acessados**: `Sheet3` (toda a Base).
12. **Pré-condições**: Base já extraída (Lin_Final > Lin_Inicial).
13. **Passos principais**:
    - Localiza extensão da Base.
    - Pergunta ao usuário (`MsgBox vbYesNo`) se deseja recalcular Empresa/Diretorias/Segmento (com aviso de que será necessário reextrair a base Hubble depois).
    - Se sim: limpa histórico Hubble, recalcula `Form_Empresa`, `Form_Segmentos`, `Form_Classe`, `Form_Diretoria_Gerencial_Com_Ref_Cruzada`.
    - Sempre executa: `Form_Opex_Driven`, `Form_Ref_Organic`, as 3 reclassificações, `Calcular_Comb_Meses_Intervalo`, `Form_Calcular_FY`.
    - `Ativar_Tudo` + `PopUp_Tempo_Processamento`.
14. **Pós-condições**: campos auxiliares recalculados; se o usuário confirmou o passo condicional, a base Hubble fica marcada como necessitando reextração (aviso apenas textual — não há trava técnica que impeça o uso do sistema até a reextração).
15. **Efeitos colaterais/riscos/evidência**: aviso ao usuário é apenas informativo (`"SERÁ NECESSÁRIO EXTRAIR NOVAMENTE A BASE HUBBLE!"`, linha 172) — não há nenhuma trava técnica ou flag de estado que force essa reextração; se o usuário ignorar o aviso, a Base fica com dados Hubble desatualizados/inconsistentes silenciosamente. Evidência: linhas 150-200.

---

#### 32. `Tit_Msg`
1. **Nome completo**: `Tit_Msg`
2. **Módulo**: Auxiliar.bas (linha 202)
3. **Tipo**: Function
4. **Escopo**: Public
5. **Objetivo**: Retorna a string de título padrão usada em todos os `MsgBox` do sistema (`" -->>  ( QD )  Quick Data Generation  "`), garantindo identidade visual consistente das mensagens.
6. **Quem chama**: dezenas de ocorrências como argumento de `MsgBox` em `Aux_Leitura_Nome_Arqs.bas`, `Auxiliar.bas` (múltiplas), `Gerar_Base_Pre_Closing.bas`, `Form_Importacao.frm`, `Front_Processos.bas`, `Form_Exportacao.frm` — usada como parâmetro `Title` de `MsgBox` em praticamente toda mensagem ao usuário do sistema.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: String — título padrão para caixas de diálogo.
10. **Variáveis/objetos relevantes**: constante literal.
11. **Abas/intervalos acessados**: nenhuma.
12. **Pré-condições**: nenhuma.
13. **Passos principais**: retorna string fixa.
14. **Pós-condições**: N/A.
15. **Efeitos colaterais/riscos/evidência**: nenhum risco; utilitário de UI puro. Evidência: linhas 202-206.

---

#### 33. `Calcular_Comb_Meses`
1. **Nome completo**: `Calcular_Comb_Meses`
2. **Módulo**: Auxiliar.bas (linha 210)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Versão "sem parâmetros de filtro" do cálculo de combinações de meses — recalcula a fórmula de soma de intervalo de meses (via `INDIRECT`) para **toda** a Base de uma vez.
6. **Quem chama**: **Nenhuma chamada ativa encontrada** (grep por `Calcular_Comb_Meses()` só retorna a própria declaração) — **código morto**, aparentemente superado por `Calcular_Comb_Meses_Intervalo` (que aceita filtro por Campo/Chave) e por `Calcular_Comb_Meses_Intervalo_Linha` (que aceita intervalo de linhas explícito).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum (usa apenas variáveis locais).
9. **Retorno**: Function sem valor de retorno atribuído.
10. **Variáveis/objetos relevantes**: `Sh3=Sheet3`, `Col_1` (primeira coluna de "combinação de meses", "Dez"+4).
11. **Abas/intervalos acessados**: `Sheet3`, colunas após "Dez"+4 até o fim do cabeçalho.
12. **Pré-condições**: N/A (não executado).
13. **Passos principais**: localiza intervalo de linhas/colunas de combinação; aplica fórmula `SUM(INDIRECT(...))` a todo o bloco; converte em valor.
14. **Pós-condições**: N/A (código morto).
15. **Efeitos colaterais/riscos/evidência**: **código morto confirmado por grep** — mantido no módulo sem uso; mesmo padrão de fórmula `INDIRECT` (volátil) presente nas versões ativas (`Calcular_Comb_Meses_Intervalo`/`_Linha`), então o risco de performance documentado nessas duas se aplicaria aqui também caso reativada por engano. Evidência: linhas 210-242.

---

#### 34. `Gerar_Visao_Italia`
1. **Nome completo**: `Gerar_Visao_Italia`
2. **Módulo**: Auxiliar.bas (linha 245)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Gera uma **visão contábil alternativa "IFRS Itália"**, duplicando linhas da Base que casam com regras de uma tabela de mapeamento (Sheet23), aplicando multiplicador opcional e remapeando os campos Abertura_2 a Abertura_8 conforme regras "DE→PARA".
6. **Quem chama**: `Extracao_Base_MOCKUP_RGM.bas:245`, `Extracao_Base_1009.bas:391` (chamada comentada em `:282`, ou seja, desativada em um ponto do fluxo 1009 mas ativa em outro), `Extracao_Base_Other_Inco.bas:366`, `Extracao_Base_RGM.bas:243`, `Extracao_Fixed_Revenues.bas:246`, `Extracao_Sheet_Ajustes.bas:272`, `Extracao_Base_Quick_Data.bas:303` (comentada em `:211`, ativa em `:303`), `Extracao_Base_Consolidad.bas:288` — chamada por quase todos os módulos de extração ao final do pipeline de cada fonte.
7. **Procedimentos chamados**: nenhum diretamente (usa `EntireRow.Copy`/`PasteSpecial`).
8. **Parâmetros**: `Sh_Destino`, `Lin_Destino`, `Ult_Lin_Destino` (ByRef, atualizado com o novo total de linhas), `Lin_Cabecalho`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído (efeito via `ByRef` em `Ult_Lin_Destino`).
10. **Variáveis/objetos relevantes**: `Sh_IT=Sheet23`; `Qtd_Lin_IT` (contador de linhas geradas); `Chave_Pesq` (concatenação de critérios Abertura_2..8).
11. **Abas/intervalos acessados**: `Sheet3` (leitura de Classe Custo e Aberturas, escrita das novas linhas), `Sheet23` (colunas "Chave", "Classe Custo", "TIPO", "DE_A2..DE_A8", "*Multip*", "PARA_A2..PARA_A8").
12. **Pré-condições**: Sheet23 deve estar preenchida com as regras de mapeamento Itália.
13. **Passos principais**:
    - **Laço triplo**: para cada linha de destino (Lin_Verif), para cada linha de regra em Sheet23 (Lin_IT), para cada coluna de critério "DE_A2..DE_A8" (Col).
    - Monta chave de critério a partir das colunas preenchidas na regra.
    - Se Classe Custo bate exatamente OU todos os critérios de Abertura batem: marca `Gerar_Lin_IT=True`.
    - Se marcado: copia a linha inteira (`EntireRow.Copy`/`PasteSpecial xlPasteFormulasAndNumberFormats`) para o final da Base, aplica multiplicador (se != 1) via `PasteSpecial Operation:=xlMultiply`, sobrescreve Abertura_2..8 conforme "PARA_A2..PARA_A8", marca "Visao"="IFRS Itália".
    - Atualiza `Ult_Lin_Destino` com o total de linhas geradas.
14. **Pós-condições**: novas linhas "IFRS Itália" anexadas ao final da Base.
15. **Efeitos colaterais/riscos/evidência**: **laço triplo (linhas destino × linhas de regra × colunas de critério) com `EntireRow.Copy`/`PasteSpecial` no caminho mais interno** — para uma Base de dezenas de milhares de linhas e uma tabela de regras não-trivial, este é o candidato mais forte a gargalo de performance do módulo (potencial O(n×m)); `Application.StatusBar` mostra progresso percentual (linha 299-300), confirmando que a própria equipe já reconhece que esta rotina é lenta o suficiente para justificar feedback visual; chamada comentada/ativa de forma inconsistente entre módulos de extração similares (1009 e Quick_Data têm um `Call` comentado em um ponto e outro ativo mais à frente no mesmo arquivo) sugere refatoração incompleta. Evidência: linhas 245-375, feedback de progresso 299-300.

---

#### 35. `Calcular_Comb_Meses_Intervalo`
1. **Nome completo**: `Calcular_Comb_Meses_Intervalo`
2. **Módulo**: Auxiliar.bas (linha 378)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Recalcula a fórmula de "combinação de meses" (soma de intervalos tipo "Jan:Mar" via `INDIRECT`) para um subconjunto da Base filtrado por um campo/chave (ex.: Fonte="Base_Hubble", ou "Geral" para tudo, ou o caso especial "Base_Consolidada"). Antes de localizar o bloco, **ordena toda a Base** por 5 chaves.
6. **Quem chama**: `Gerar_Base_Pre_Closing.bas:313` (`"Fonte","Geral"`), `Form_Importacao.frm:640`, `Auxiliar.bas:23` (dentro de `Extrair_Todas_as_Bases`), `Auxiliar.bas:192` (dentro de `Atualizar_Base_Todos_Campos_Auxiliares`). Chamadas **comentadas** (desativadas) em `Extracao_Base_Consolidad.bas:15`, `Extracao_Base_1009.bas:14`, `Extracao_Base_RGM.bas:14`, `Extracao_Sheet_Ajustes.bas:15`, `Extracao_Base_Quick_Data.bas:15`, `Extracao_Base_Other_Inco.bas:15`, `Extracao_Base_MOCKUP_RGM.bas:14`, `Extracao_Fixed_Revenues.bas:14`, `Extracao_SQL_Hubble.bas:14` — ou seja, **todos os módulos de extração individuais têm a chamada comentada**; apenas os fluxos "globais" (`Extrair_Todas_as_Bases`, `Atualizar_Base_Todos_Campos_Auxiliares`, `Gerar_Base_Pre_Closing`, importação de Front) a chamam de fato. Isso sugere que, por fonte individual, o recálculo de combinação de meses é feito via `Calcular_Comb_Meses_Intervalo_Linha` (mais barato, sem reordenar a Base inteira).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Campo`, `Chave`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Order_1..5` (colunas de ordenação: Campo informado, "KPI", "VERSÃO", "EXERCICIO", "TIPO NÍVEL 2").
11. **Abas/intervalos acessados**: `Sheet3` — toda a extensão da Base é reordenada.
12. **Pré-condições**: colunas "KPI", "VERSÃO", "EXERCICIO", "TIPO NÍVEL 2" e o `Campo` informado devem existir no cabeçalho.
13. **Passos principais**:
    - Localiza extensão da Base.
    - **Ordena toda a Base** por 5 chaves (Campo, KPI, Versão, Exercício, Tipo Nível 2).
    - Resolve `Lin_Inicial`/`Lin_Final` conforme `Chave`: "Geral" = toda a base; "Base_Consolidada" = trata o caso especial de múltiplos prefixos "RPD*"/"Preview *" como um bloco único; senão, bloco exato da Chave via `CountIfs`/`Match`.
    - Aplica fórmula `SUM(INDIRECT(...))` ao bloco de colunas de combinação, calcula, converte em valor.
14. **Pós-condições**: Base reordenada (efeito colateral permanente na ordem das linhas); colunas de combinação de meses recalculadas para o bloco filtrado.
15. **Efeitos colaterais/riscos/evidência**: **reordena toda a Base como efeito colateral** mesmo quando o filtro pede apenas um subconjunto pequeno — custo desproporcional ao objetivo; regra especial "Base_Consolidada" trata dois prefixos diferentes ("RPD*" e "Preview *") como um único bloco contíguo, assumindo que ficam adjacentes após a ordenação — hipótese frágil se a ordenação não garantir essa adjacência em todos os casos (`[VALIDAR COM O NEGÓCIO]`); uso de `INDIRECT` (fórmula volátil). Evidência: linhas 378-463, caso especial 428-439.

---

#### 36. `Calcular_Comb_Meses_Intervalo_Linha`
1. **Nome completo**: `Calcular_Comb_Meses_Intervalo_Linha`
2. **Módulo**: Auxiliar.bas (linha 466)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Variante mais barata de `Calcular_Comb_Meses_Intervalo` — recalcula a combinação de meses para um intervalo de linhas **explícito** (sem reordenar a Base, sem localizar por chave), usada pelos módulos de extração individuais logo após inserir suas linhas na Base.
6. **Quem chama**: `Extracao_SQL_Hubble.bas:359`, `Extracao_Sheet_Ajustes.bas:279`, `Extracao_Base_1009.bas:400`, `Extracao_Fixed_Revenues.bas:253`, `Extracao_Base_RGM.bas:250`, `Extracao_Base_Quick_Data.bas:310`, `Extracao_Base_Other_Inco.bas:373`, `Extracao_Base_Consolidad.bas:301`, `Extracao_Base_MOCKUP_RGM.bas:252` — chamada por **todos** os módulos de extração individual (o par exato dos que tinham a chamada de `Calcular_Comb_Meses_Intervalo` comentada).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Lin_Inicial`, `Lin_Final`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh3=Sheet3`; mesma fórmula `INDIRECT` das demais.
11. **Abas/intervalos acessados**: `Sheet3`, apenas o bloco de colunas de combinação, para as linhas informadas.
12. **Pré-condições**: `Lin_Inicial`/`Lin_Final` devem já corresponder ao bloco recém-inserido pelo chamador.
13. **Passos principais**: localiza colunas de combinação; aplica fórmula, atribui direto o `.Value` do próprio range sobre si mesmo (linha 488) — variante do bake-in que evita `.Copy`/`PasteSpecial` (usa atribuição direta `Range = Range.Value`).
14. **Pós-condições**: colunas de combinação de meses do bloco recalculadas.
15. **Efeitos colaterais/riscos/evidência**: por não reordenar nem escanear a Base inteira, é significativamente mais barata que `Calcular_Comb_Meses_Intervalo` — bom padrão de design quando o intervalo já é conhecido; ainda assim usa `INDIRECT` (volátil) para todo o bloco. Evidência: linhas 466-493.

---

#### 37. `Limpar_Base_Historica`
1. **Nome completo**: `Limpar_Base_Historica`
2. **Módulo**: Auxiliar.bas (linha 496)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Remove da Base o bloco de linhas correspondente a uma Fonte específica (ou tudo, se "Geral") antes de uma nova extração — garante idempotência (não duplicar dados) ao reprocessar uma fonte.
6. **Quem chama**: `Extracao_Base_MOCKUP_RGM.bas:29`, `Extracao_Base_1009.bas:29,38` (chamadas comentadas em 32,35), `Auxiliar.bas:14` (Extrair_Todas_as_Bases, "Geral"), `Auxiliar.bas:37` (Limpar_Todas_as_Bases), `Auxiliar.bas:176` (Atualizar_Base_Todos_Campos_Auxiliares, "Base_Hubble"), `Extracao_Base_Consolidad.bas:30`, `fx_IFRS16.bas:105` (Chave="DELETAR", pós-tratamento IFRS16), `Extracao_SQL_Hubble.bas:39`, `Extracao_Base_RGM.bas:29`, `Extracao_Fixed_Revenues.bas:30`, `Extracao_Sheet_Ajustes.bas:12,30`, `Extracao_Base_Quick_Data.bas:30`, `Gerar_Base_Pre_Closing.bas:293`, `Extracao_Base_Other_Inco.bas:30` — usada por **todos** os módulos de extração, mais o fluxo IFRS16.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Campo_Pesq`, `Tipo`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Lin_Inicial`, `Lin_Final` (localizados por varredura linear após ordenação).
11. **Abas/intervalos acessados**: `Sheet3` — reordenada por 4 chaves (colunas B-E) quando `Tipo <> "Geral"`.
12. **Pré-condições**: coluna `Campo_Pesq` deve existir no cabeçalho.
13. **Passos principais**:
    - `ActiveSheet.ShowAllData` (remove filtros ativos, `On Error Resume Next`).
    - Se `Tipo="Geral"`: deleta da linha seguinte ao cabeçalho até 1000 linhas além do fim atual (folga de segurança).
    - Senão: ordena a Base pelas colunas B, C, D, E; localiza o bloco contíguo de linhas cujo `Campo_Pesq` bate com `Tipo` — caso especial "BASE_CONSOLIDADA" trata prefixos "RPD"/"Preview" como um único bloco (mesma lógica de `Calcular_Comb_Meses_Intervalo`).
    - Deleta o bloco encontrado via `EntireRow.Delete`.
14. **Pós-condições**: linhas da fonte especificada removidas da Base (ou toda a Base, se "Geral").
15. **Efeitos colaterais/riscos/evidência**: reordena a Base como efeito colateral para localizar o bloco a deletar (mesmo padrão de custo de `Calcular_Comb_Meses_Intervalo`); a folga de "1000 linhas além do fim atual" no caso "Geral" (linha 519) é um número mágico sem explicação — risco de não cobrir tudo se houver mais de 1000 linhas de sobra, ou de ser desnecessariamente conservador; `On Error Resume Next` ao redor de `ShowAllData` mascara silenciosamente qualquer erro de filtro. Evidência: linhas 496-564.

---

#### 38. `Refresh_Base_Aux`
1. **Nome completo**: `Refresh_Base_Aux`
2. **Módulo**: Auxiliar.bas (linha 567)
3. **Tipo**: Sub
4. **Escopo**: Public
5. **Objetivo**: Botão que atualiza em sequência todas as bases auxiliares de suporte (Segmento, De-Para/Ref. Cruzadas, Suporte de Linhas, Drop de Combinações, Linhas Válidas) — manutenção periódica das tabelas de-para que alimentam o motor de classificação.
6. **Quem chama**: **Nenhuma chamada encontrada.** `[INFERÊNCIA]` botão na aba Extracao. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Verifica_Versao`, `Refresh_Base_Segmento`, `Refresh_Base_De_Para_Ref_Cruzadas`, `Refresh_Base_Suporte_Linhas`, `Refresh_Drop_Comb_Hubble`, `Extrair_Valid_Lin` (todos fora deste cluster).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma além das chamadas.
11. **Abas/intervalos acessados**: indiretamente, todas as tabelas de-para (Sup_Linhas, referências cruzadas, DropComb).
12. **Pré-condições**: conectividade SQL (via `Verifica_Versao`).
13. **Passos principais**: `Verifica_Versao`; chama as 5 rotinas de refresh em sequência; seleciona Sheet8; `MsgBox` de sucesso.
14. **Pós-condições**: tabelas auxiliares atualizadas; `MsgBox "Bases auxiliares extraídas com sucesso!"`.
15. **Efeitos colaterais/riscos/evidência**: nenhuma chamada `Desligar_Tudo`/`Ativar_Tudo` ao redor desta Sub especificamente — presume-se que cada rotina de refresh interna gerencia seu próprio estado de otimização (não verificável neste cluster, pois essas rotinas estão em outros módulos). Evidência: linhas 567-580.

---

#### 39. `Carregar_Sheets_Suporte`
1. **Nome completo**: `Carregar_Sheets_Suporte`
2. **Módulo**: Auxiliar.bas (linha 582)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Monta a lista de nomes de planilhas de suporte (concatenada em `;`) tanto na variável global `Sheets_Suporte` quanto em uma coluna da planilha Sheet11 — provavelmente para popular um combo/lista de seleção em formulários de importação/exportação.
6. **Quem chama**: `Form_Exportacao.frm:516`, `Form_Importacao.frm:85`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído (efeito via global `Sheets_Suporte` e escrita em Sheet11).
10. **Variáveis/objetos relevantes**: `Sheets_Suporte` (global, string `;`-delimitada com 20 planilhas nomeadas explicitamente: Sheet10, 11, 14-26, 2, 3, 7, 8, 9).
11. **Abas/intervalos acessados**: `Sheet11` (coluna "Sheets_Suporte", limpa linhas 2-100 e repovoa).
12. **Pré-condições**: nenhuma.
13. **Passos principais**: monta string fixa com os 20 nomes de planilha de suporte; limpa a coluna correspondente em Sheet11; para cada planilha do workbook cujo nome aparece na lista, adiciona uma linha em Sheet11.
14. **Pós-condições**: `Sheets_Suporte` (global) e coluna em Sheet11 preenchidos.
15. **Efeitos colaterais/riscos/evidência**: lista de planilhas de suporte é **hardcoded por codenome** (Sheet10, Sheet11, etc.) — se uma nova planilha de suporte for adicionada ao sistema, alguém precisa lembrar de atualizar esta lista manualmente; limite de 100 linhas na limpeza (`sh.Range(..., sh.Cells(100, Col))`, linha 594) é um número mágico que assume no máximo ~98 planilhas de suporte. Evidência: linhas 582-603.

---

#### 40. `Reclassificar_Combinacoes_Empresas`
1. **Nome completo**: `Reclassificar_Combinacoes_Empresas`
2. **Módulo**: Auxiliar.bas (linha 606)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Constrói dinamicamente (a partir de uma matriz de critérios em Sheet9/"DropComb") uma cadeia de fórmulas `IF(AND(...),"S","N")` que classifica cada linha da Base quanto a pertencer a cada "combinação de Empresas" pré-definida — usado por segmentações/tabelas dinâmicas dos relatórios.
6. **Quem chama**: `Extracao_Base_1009.bas:397`, `Extracao_Base_RGM.bas:247`, `Auxiliar.bas:189` (Atualizar_Base_Todos_Campos_Auxiliares), `Extracao_SQL_Hubble.bas:356`, `Extracao_Base_Quick_Data.bas:307`, `Extracao_Base_MOCKUP_RGM.bas:249`, `Extracao_Base_Other_Inco.bas:370`, `Extracao_Sheet_Ajustes.bas:276`, `Extracao_Base_Consolidad.bas:298`, `Extracao_Fixed_Revenues.bas:250`, `Form_Importacao.frm:643` — chamada por todos os módulos de extração e pela importação de Front. **Nota**: existe uma versão paralela `Reclassificar_Combinacoes_Empresas_TK` em `TK_Functions.bas`, chamada por um caminho de invocação diferente (ver item 74).
7. **Procedimentos chamados**: `Criar_Formula_Filtros` (chamado repetidamente dentro dos laços de montagem de fórmula).
8. **Parâmetros**: `Lin_Inicial`, `Ult_Linh`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9` ("DropComb"); `Formula_Empresa_PARTE_1`/`_PARTE_2` (a matriz de colunas-combinação é dividida ao meio e processada em 2 fórmulas separadas); `vCOL_Metade` (coluna que separa as duas metades).
11. **Abas/intervalos acessados**: `Sheet3` (colunas "EMPRESA" e o bloco de colunas nomeadas "EMPRESA" repetidas — cada uma representa uma combinação), `Sheet9` (colunas prefixadas "Empresa*", linhas de critério a partir da linha 3).
12. **Pré-condições**: `Sheet9` deve conter a matriz de critérios já estruturada (linha 2 = nomes de empresa, linha 3+ = combinações com marcadores por empresa).
13. **Passos principais**:
    - Localiza a combinação "do meio" (`qtde_Combinacoes_Empresas / 2`) para dividir o processamento em 2 partes.
    - **Parte 1**: para cada combinação até a metade, monta cláusula `IF(AND(Fonte=Criterio, filtros de empresa),"S", ...)` aninhada, com critério final `ELSE` comparando o cabeçalho diretamente à coluna Empresa.
    - **Parte 2**: mesmo processo para a segunda metade das combinações.
    - Aplica `Formula_Empresa_PARTE_1` ao bloco de colunas até `vCOL_Metade`, `PARTE_2` ao restante.
    - `sh.Calculate`; converte todo o bloco em valor via atribuição direta (`Range = Range.Value`, não `.Copy`/`PasteSpecial`).
14. **Pós-condições**: bloco de colunas "EMPRESA" (uma por combinação) preenchido com "S"/"N" fixo.
15. **Efeitos colaterais/riscos/evidência**: a divisão da matriz em 2 partes (código comentado nas linhas 765-769 mostra uma versão anterior sem divisão) parece ter sido introduzida para performance (fórmulas muito longas/aninhadas podem estourar limites do Excel ou ficar lentas) mas **não há comentário explicando a razão** — `[VALIDAR COM O NEGÓCIO]`; `Criar_Formula_Filtros` é chamado dentro de um laço `Do Until`, potencialmente centenas de vezes por execução, cada chamada concatenando strings (custo O(n²) de concatenação de string em VBA para strings muito longas). Evidência: linhas 606-778, divisão em 2 partes 625-763.

---

#### 41. `Reclassificar_Combinacoes_IFRS_Contabil`
1. **Nome completo**: `Reclassificar_Combinacoes_IFRS_Contabil`
2. **Módulo**: Auxiliar.bas (linha 781)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Mesmo padrão de `Reclassificar_Combinacoes_Empresas`, mas para combinações de "IFRS_Contabil" — sem a divisão em 2 partes (fórmula única).
6. **Quem chama**: `Auxiliar.bas:190` (Atualizar_Base_Todos_Campos_Auxiliares), `Extracao_Fixed_Revenues.bas:251`, `Extracao_Base_RGM.bas:248`, `Extracao_Base_Quick_Data.bas:308`, `Extracao_Base_Other_Inco.bas:371`, `Extracao_Base_MOCKUP_RGM.bas:250`, `Extracao_Base_1009.bas:398`, `Extracao_Base_Consolidad.bas:299`, `Extracao_SQL_Hubble.bas:357`, `Form_Importacao.frm:644`, `Extracao_Sheet_Ajustes.bas:277` — mesmo padrão de chamador que a versão Empresas. Existe versão paralela `_TK` em `TK_Functions.bas` (item 72), porém chamada apenas por `fx_IFRS16.bas` (não por `UPDATE_Combinacoes_*` do próprio TK_Functions — ver observação no item 72).
7. **Procedimentos chamados**: `Criar_Formula_Filtros`.
8. **Parâmetros**: `Lin_Inicial`, `Ult_Linh`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9`; ignora linhas de critério que sejam separadores visuais (`"------------"` ou `"============"`, linha 806).
11. **Abas/intervalos acessados**: `Sheet3` ("IFRS_Contabil"), `Sheet9` (colunas prefixadas "IFRS_Contabil*").
12. **Pré-condições**: idem item 40.
13. **Passos principais**: monta fórmula única (sem divisão em 2 partes) com a mesma lógica de `Criar_Formula_Filtros`; aplica ao bloco de colunas; converte em valor.
14. **Pós-condições**: bloco de colunas "IFRS_Contabil" preenchido com "S"/"N".
15. **Efeitos colaterais/riscos/evidência**: diferente de `Reclassificar_Combinacoes_Empresas`, aqui **não há divisão em 2 partes** — se a matriz de combinações IFRS crescer muito, pode reintroduzir o mesmo problema que motivou a divisão na versão Empresas, sem que ninguém tenha replicado a correção aqui (`[VALIDAR COM O NEGÓCIO]` se já foi necessário). Evidência: linhas 781-857.

---

#### 42. `Reclassificar_Combinacoes_Proforma`
1. **Nome completo**: `Reclassificar_Combinacoes_Proforma`
2. **Módulo**: Auxiliar.bas (linha 860)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Mesmo padrão das duas anteriores, para combinações de "Proforma".
6. **Quem chama**: `Extracao_Sheet_Ajustes.bas:278`, `Extracao_Base_Other_Inco.bas:372`, `Extracao_Base_1009.bas:399`, `Extracao_Base_Consolidad.bas:300`, `Auxiliar.bas:191` (Atualizar_Base_Todos_Campos_Auxiliares), `Extracao_Fixed_Revenues.bas:252`, `Extracao_Base_MOCKUP_RGM.bas:251`, `Form_Importacao.frm:645`, `Extracao_SQL_Hubble.bas:358`, `Extracao_Base_Quick_Data.bas:309`, `Extracao_Base_RGM.bas:249`.
7. **Procedimentos chamados**: `Criar_Formula_Filtros`.
8. **Parâmetros**: `Lin_Inicial`, `Ult_Linh`.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9`; mesma exclusão de linhas separadoras visuais.
11. **Abas/intervalos acessados**: `Sheet3` ("Proforma"), `Sheet9` (colunas prefixadas "Proforma*").
12. **Pré-condições**: idem item 40.
13. **Passos principais**: idêntico a `Reclassificar_Combinacoes_IFRS_Contabil`, trocando o campo-alvo para "Proforma".
14. **Pós-condições**: bloco de colunas "Proforma" preenchido com "S"/"N".
15. **Efeitos colaterais/riscos/evidência**: as três rotinas (`_Empresas`, `_IFRS_Contabil`, `_Proforma`) compartilham ~90% da mesma estrutura de código copiada 3 vezes (loop de leitura de critério, chamada a `Criar_Formula_Filtros`, montagem do `ELSE`, fechamento de parênteses, aplicação e bake-in) — forte candidato a refatoração/parametrização única na reescrita, hoje é lógica triplicada com risco de correções aplicadas em uma cópia e esquecidas nas outras duas. Evidência: linhas 860-936.

---

#### 43. `Criar_Formula_Filtros`
1. **Nome completo**: `Criar_Formula_Filtros`
2. **Módulo**: Auxiliar.bas (linha 939)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: **Motor de tradução de regras** — converte uma linha de critérios da matriz DropComb (string `;`-delimitada de valores de filtro, com suporte a exclusão via "<>") em um fragmento de fórmula Excel `IF(AND(...),"S", ...)`, reutilizado pelas 3 rotinas de reclassificação (e pelas 3 versões `_TK`).
6. **Quem chama**: `Auxiliar.bas` (8x, dentro de `Reclassificar_Combinacoes_Empresas/IFRS_Contabil/Proforma`), `TK_Functions.bas` (6x, dentro das versões `_TK` das mesmas 3 rotinas) — **compartilhado entre as duas famílias de reclassificação** (a única peça de lógica de negócio genuinamente reaproveitada entre `Auxiliar.bas` e `TK_Functions.bas`, ao invés de duplicada).
7. **Procedimentos chamados**: nenhum (usa apenas `WorksheetFunction.Match` para localizar a linha "LIN_BASE").
8. **Parâmetros**: `Formula` (ByRef, acumulador — a string cresce a cada chamada), `Criterio`, `Filtro`, `Col_Verif`.
9. **Retorno**: N/A (efeito via `ByRef Formula`).
10. **Variáveis/objetos relevantes**: `str_Array` (split do `Filtro` por vírgula); `Complemento_1` (critérios de exclusão "<>", unidos em AND); `Complemento_2` (critérios de igualdade, unidos em OR).
11. **Abas/intervalos acessados**: `Sheet3` apenas para localizar a linha "LIN_BASE" (usado na fórmula gerada, não como leitura de dado).
12. **Pré-condições**: `Filtro` deve estar no formato esperado (itens separados por vírgula, prefixo de operador antes do nome).
13. **Passos principais**:
    - Faz `Split(Filtro, ",")`.
    - Para cada item: se contém "<>", vai para `Complemento_1` (lógica AND); senão, vai para `Complemento_2` (lógica OR, com aspas ao redor se não for uma expressão "TRIM").
    - Monta fragmento `IF(AND(R[LIN_BASE]C=Criterio [,Complemento_1][,OR(Complemento_2)]),"S",` e concatena ao acumulador `Formula`.
14. **Pós-condições**: `Formula` (ByRef) cresce com mais um nível de `IF` aninhado.
15. **Efeitos colaterais/riscos/evidência**: concatenação de string em `Formula` (parâmetro ByRef acumulador) dentro de um laço externo, chamado potencialmente centenas de vezes — para uma matriz de critérios grande, a string de fórmula final pode se aproximar do limite de 8.192 caracteres de uma fórmula do Excel, risco não tratado no código (nenhuma validação de tamanho). Evidência: linhas 939-971.

---

#### 44. `Processo_Exclusao_Linhas_Base`
1. **Nome completo**: `Processo_Exclusao_Linhas_Base`
2. **Módulo**: Auxiliar.bas (linha 974)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Exclui blocos de linhas da Base cujo campo de Abertura contenha (via `InStr`, correspondência parcial) um dos critérios informados — usado por módulos de extração para remover linhas indesejadas após a inserção inicial (ex.: categorias específicas de Other Income ou Base Consolidada que não devem compor a Base final).
6. **Quem chama**: `Extracao_Base_Other_Inco.bas:397,408` (comentada em 402), `Extracao_Base_Consolidad.bas:325,331`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Lin_Destino`, `Ult_Lin_Destino` (ByRef, decrementado a cada bloco deletado), `Col_Abertura`, `Criterio`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Lin_Inicial`/`Lin_Final` (localizados via varredura linha a linha, não via `Match`).
11. **Abas/intervalos acessados**: `Sheet3`, coluna nomeada por `Col_Abertura`.
12. **Pré-condições**: `Criterio` deve ser uma string contendo os valores-alvo (comparação por `InStr`, então valores parciais/substring também casam).
13. **Passos principais**: laço `Do` linha a linha; identifica início/fim de blocos contíguos cujo valor da coluna está contido em `Criterio`; deleta cada bloco assim que identificado (`EntireRow.Delete`), ajusta `Lin`/`Ult_Lin_Destino`, continua a varredura a partir do ponto de exclusão; `Application.StatusBar` com percentual de progresso.
14. **Pós-condições**: linhas cujo campo de Abertura corresponde ao critério removidas; `Ult_Lin_Destino` atualizado.
15. **Efeitos colaterais/riscos/evidência**: correspondência por **substring (`InStr`)**, não igualdade exata — risco de falso positivo (ex. critério "Other" excluiria também "Other Income" e qualquer outro valor que contenha essa substring); `EntireRow.Delete` dentro de laço percorrendo a Base inteira — mesmo risco de custo O(n²) documentado no relatório agregado anterior. Evidência: linhas 974-1007.

---

#### 45. `RemoveStyles`
1. **Nome completo**: `RemoveStyles`
2. **Módulo**: Auxiliar.bas (linha 1010)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Rotina de manutenção/higienização — remove todos os estilos de célula não-nativos do workbook, mediante confirmação do usuário, contra o "inchaço" de estilos que é causa comum de lentidão/corrupção em arquivos Excel legados de longa vida.
6. **Quem chama**: **Nenhuma chamada encontrada** — provavelmente executada manualmente via VBE/Immediate Window por um administrador, não ligada a nenhum botão identificável no dump. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Total_Estilos` (contagem antes da limpeza).
11. **Abas/intervalos acessados**: nenhuma planilha — opera sobre `ActiveWorkbook.Styles` (coleção de estilos do workbook inteiro).
12. **Pré-condições**: nenhuma.
13. **Passos principais**: conta estilos; `MsgBox vbYesNo` de confirmação com a contagem; se confirmado, itera `Styles` de trás para frente deletando os não-`BuiltIn`; `MsgBox` final com contagem antes/depois.
14. **Pós-condições**: apenas estilos nativos do Excel permanecem no workbook.
15. **Efeitos colaterais/riscos/evidência**: a própria existência desta rotina é evidência de que o arquivo `.xlsb` **historicamente sofre acúmulo patológico de estilos** — sintoma comum de arquivos Excel copiados/colados entre workbooks por anos; ação com confirmação prévia (`vbYesNo`), diferente de `Limpar_Todas_as_Bases`, que é destrutiva sem confirmação — inconsistência de padrão de segurança entre rotinas destrutivas do sistema. Evidência: linhas 1010-1033.

---

#### 46. `Excluir_Item_Especifico`
1. **Nome completo**: `Excluir_Item_Especifico`
2. **Módulo**: Auxiliar.bas (linha 1036)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Utilitário de suporte à UI de exportação — ordena uma planilha temporária por um campo-chave e remove o bloco de linhas correspondente a um item específico selecionado pelo usuário para exclusão.
6. **Quem chama**: `Form_Exportacao.frm:242,249`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Arq_Temp` (Workbook), `Sh_Temp` (Worksheet), `Lin_1`, `Col_1`, `Ult_Col`, `Campo`, `Item_Exclusao`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Col_Order` (coluna do campo-chave).
11. **Abas/intervalos acessados**: planilha temporária (`Sh_Temp`, de um workbook auxiliar `Arq_Temp` — não a Base).
12. **Pré-condições**: `Campo` deve existir no cabeçalho de `Sh_Temp`.
13. **Passos principais**: ordena `Sh_Temp` pelo campo-chave; localiza bloco de linhas cujo valor é `Item_Exclusao` via `Match`/`CountIfs`; deleta o bloco.
14. **Pós-condições**: item removido da planilha temporária de exportação.
15. **Efeitos colaterais/riscos/evidência**: opera sobre workbook/planilha temporária de exportação, não sobre a Base principal — risco de escopo limitado à sessão de exportação do usuário. Evidência: linhas 1036-1077.

---

#### 47. `Extrair_Info_Colunas_Fixas`
1. **Nome completo**: `Extrair_Info_Colunas_Fixas`
2. **Módulo**: Auxiliar.bas (linha 1081)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Copia uma coluna nomeada (por fórmula, não valor) de uma planilha de origem para uma planilha de destino durante o processo de importação de layouts "Front".
6. **Quem chama**: `Form_Importacao.frm` (4x, linhas 272-284).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `fn` (WorksheetFunction, passado por parâmetro em vez de instanciado localmente — padrão diferente das outras Subs do módulo), `Chave`, `Ult_Lin`, `Sh_Base`, `Lin_Cabecalho_Base`, `Col_Cabecalho_Base`, `Sh_Destino`, `Lin_Cabecalho_Destino`, `Col_Cabecalho_Destino`.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma além dos parâmetros.
11. **Abas/intervalos acessados**: `Sh_Base`/`Sh_Destino` (planilhas de layout Front, não necessariamente a Base principal).
12. **Pré-condições**: `Chave` deve existir nos cabeçalhos de origem e destino.
13. **Passos principais**: localiza colunas via `Match`; se a coluna existir na origem, copia (`PasteSpecial xlPasteFormulas` — preserva fórmula, não converte em valor, diferente da maioria das outras Subs do cluster).
14. **Pós-condições**: coluna copiada (como fórmula) para o destino.
15. **Efeitos colaterais/riscos/evidência**: único procedimento do cluster que recebe `fn` (WorksheetFunction) como parâmetro em vez de instanciar `Set fn = Application.WorksheetFunction` internamente — inconsistência de estilo que sugere ter sido escrito por outra pessoa ou em outra época; cola como fórmula (`xlPasteFormulas`), não como valor — diferente do padrão dominante do resto do cluster, risco de manter dependência viva de referências entre workbooks/planilhas de layout. Evidência: linhas 1081-1093.

---

#### 48. `Ordenar_Lista`
1. **Nome completo**: `Ordenar_Lista`
2. **Módulo**: Auxiliar.bas (linha 1096)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Ordena alfabeticamente (bubble sort simples) uma lista `,`-delimitada de itens, colocando o item "Total" sempre por último.
6. **Quem chama**: `Form_Exportacao.frm:563`. **Atenção**: `BackupCodigo_MainResults.bas:170` também chama `Ordenar_Lista`, mas esse módulo **possui sua própria cópia `Private Sub Ordenar_Lista`** na linha 188 — em VBA, uma chamada a um nome de procedimento dentro do próprio módulo se resolve para a versão local antes de procurar em outros módulos, então essa chamada **não** aciona esta Function de `Auxiliar.bas` (ver detalhamento no risco #4, seção D).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `Filtro_Novo`, `Filtro_Novo_Ordenado` (ByRef, saída).
9. **Retorno**: Function sem valor de retorno nomeado atribuído (saída via `ByRef Filtro_Novo_Ordenado`).
10. **Variáveis/objetos relevantes**: `MyArray` (array resultante do `Split`).
11. **Abas/intervalos acessados**: nenhuma — opera puramente sobre strings em memória.
12. **Pré-condições**: `Filtro_Novo` deve ser uma lista `,`-delimitada.
13. **Passos principais**: `Split` por vírgula; bubble sort (`For a`/`For b`) colocando "Total" sempre por último e demais itens em ordem alfabética ascendente; remonta a string ordenada.
14. **Pós-condições**: `Filtro_Novo_Ordenado` contém a lista ordenada.
15. **Efeitos colaterais/riscos/evidência**: **mesmo nome de procedimento existe em dois módulos diferentes com escopos diferentes** (`Public Function` aqui, `Private Sub` em `BackupCodigo_MainResults.bas`) — armadilha clássica de manutenção: um desenvolvedor que faça `Ctrl+F` por "Ordenar_Lista" e edite a cópia errada (ou assuma que há só uma implementação) pode introduzir bugs sutis, já que o comportamento real depende de qual módulo está chamando; algoritmo bubble sort O(n²) é aceitável apenas para listas curtas (nomes de campo/filtro), não deve ser usado para listas grandes. Evidência: linhas 1096-1116; colisão de nome em BackupCodigo_MainResults.bas:188.

---

### A.3 — Módulo `TK_Functions.bas` (26 procedimentos)

> Nota geral: `TK_Functions.bas` é uma biblioteca heterogênea (utilitários genuínos + rotinas de manutenção + a família `_TK` de reclassificação). Vários procedimentos deste módulo **não têm nenhuma chamada ativa encontrada no dump** — ver ressalva metodológica no topo do documento sobre botões (`Shape.OnAction`) não capturáveis em texto; onde aplicável, distingue-se explicitamente "sem chamador encontrado" (pode ser botão não capturado) de "referenciado apenas em código comentado" (evidência mais forte de órfão/morto).

---

#### 49. `PEGA_A_DESCRICAO_DA_CONTA`
1. **Nome completo**: `PEGA_A_DESCRICAO_DA_CONTA`
2. **Módulo**: TK_Functions.bas (linha 2)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Retorna a descrição textual de uma Classe de Custo (Conta), consultando a tabela Sup_Linhas por correspondência exata.
6. **Quem chama**: `Extracao_Fixed_Revenues.bas:388`, `Extracao_Base_RGM.bas:391` (ambas atribuindo o resultado a `Col_Destino_CC + 1` durante a extração), `TK_Functions.bas:27` (interno, dentro de `PEGA_A_DESCRICAO_DA_CONTA_Independente`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `vCC As String` (Classe de Custo a buscar).
9. **Retorno**: String — descrição da conta, ou `""` se não encontrada.
10. **Variáveis/objetos relevantes**: `vLinha` (linha encontrada via `Match`).
11. **Abas/intervalos acessados**: `Sheets("Sup_Linhas")`, coluna `BZ` (chave de busca) e `CA` (descrição retornada).
12. **Pré-condições**: nenhuma — tratamento de erro cobre o caso de não encontrar.
13. **Passos principais**: `On Error GoTo fim`; `Match(vCC, Sup_Linhas!BZ:BZ, 0)`; retorna `Sup_Linhas!CA<linha>`; se erro, retorna string vazia.
14. **Pós-condições**: N/A (função pura).
15. **Efeitos colaterais/riscos/evidência**: uso de colunas por letra (`BZ`, `CA`) em vez de nome de cabeçalho — diferente do padrão dominante do cluster (que localiza colunas via `Match` pelo nome), tornando esta função frágil a inserção/remoção de colunas em Sup_Linhas. Evidência: linhas 2-10.

---

#### 50. `PEGA_A_DESCRICAO_DA_CONTA_Independente`
1. **Nome completo**: `PEGA_A_DESCRICAO_DA_CONTA_Independente`
2. **Módulo**: TK_Functions.bas (linha 12)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Aplica `PEGA_A_DESCRICAO_DA_CONTA` célula a célula sobre a coluna W da planilha "Base", escrevendo o resultado na coluna X — utilitário manual para popular descrições de conta sob demanda, com validações de pré-condição via `MsgBox`.
6. **Quem chama**: **Nenhuma chamada encontrada** — utilitário standalone de uso manual, ativado provavelmente por atalho/VBE. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `PEGA_A_DESCRICAO_DA_CONTA` (linha 27).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `qtde_linhas` (via `CountIf` de W6:W100000 não-vazio).
11. **Abas/intervalos acessados**: `Sheets("Base")` — coluna `W` (leitura, Classe de Custo) e `X` (escrita, descrição), com checagem de cabeçalho `X5 = "Descricao CC"`.
12. **Pré-condições**: usuário deve ter manualmente preparado a coluna X5 com o cabeçalho exato "Descricao CC"; deve haver ao menos uma CC preenchida em W6:W100000.
13. **Passos principais**: valida `qtde_linhas > 0` e `X5="Descricao CC"` (aborta com `MsgBox` se não); laço `For l = 6 To qtde_linhas` limpando e repreenchendo X.
14. **Pós-condições**: coluna X da Base preenchida com descrições de conta.
15. **Efeitos colaterais/riscos/evidência**: **bug de limite do laço** — usa `qtde_linhas` (a *contagem* de CCs preenchidas) como linha final do laço `For l = 6 To qtde_linhas`, em vez de `5 + qtde_linhas` (a *linha* correspondente); se houver qualquer linha em branco intercalada na coluna W antes do fim dos dados, o laço para prematuramente e deixa linhas sem descrição — bug potencial não coberto por teste algum. Evidência: linhas 12-30, especialmente linha 25.

---

#### 51. `fn_ListAllErrors`
1. **Nome completo**: `fn_ListAllErrors`
2. **Módulo**: TK_Functions.bas (linha 31)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Registra uma linha de log estruturado (procedimento, sheet, linha/coluna Excel, valor de erro, usuário, timestamp) na planilha `tk_Lista_de_erros` — mecanismo de auditoria de erros de fórmula.
6. **Quem chama**: **Nenhuma chamada ativa** — a única referência no dump é um `Call` **comentado** em `Aux_Formulas_Base.bas:466` (dentro de `Form_Segmentos`, ver item 7 da seção A.1). Ou seja, o mecanismo de log estruturado de erros **existe mas está desconectado** do único ponto do cluster (`Form_Segmentos`) onde erro de linha é tratado — lá, o erro vira apenas um `MsgBox` interativo, sem registro em log.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `fn_Name`, `row_VBA`, `sheet_Name`, `row_Excel`, `col_Excel`, `vValor` (todos sem tipo — Variant).
9. **Retorno**: Function sem valor de retorno nomeado atribuído (efeito via escrita direta na planilha de log).
10. **Variáveis/objetos relevantes**: `actual_ROW` (próxima linha livre via `CountA`).
11. **Abas/intervalos acessados**: `Sheets("tk_Lista_de_erros")`, colunas 1-7 (função, linha VBA, sheet, linha Excel, coluna Excel, valor, usuário/timestamp).
12. **Pré-condições**: planilha `tk_Lista_de_erros` deve existir.
13. **Passos principais**: calcula próxima linha livre; grava os 6 primeiros campos; grava `Application.UserName` na coluna 7 e depois **sobrescreve a mesma coluna 7** com `Now()` (bug: a coluna de usuário e a de timestamp deveriam ser colunas separadas, mas o código usa a coluna 7 duas vezes, perdendo o nome do usuário).
14. **Pós-condições**: nova linha de log (com o bug de perda do usuário).
15. **Efeitos colaterais/riscos/evidência**: **bug confirmado**: linhas 40-41 escrevem em `Cells(actual_ROW+1, 7)` duas vezes seguidas (primeiro `Application.UserName`, depois `Now()`), perdendo o dado do usuário — o log final nunca registra quem gerou o erro, só quando; função nunca é chamada de fato (chamada comentada), então este bug é hoje inofensivo, mas se alguém reativar a chamada sem notar o bug, o log ficará incompleto. Evidência: linhas 31-43.

---

#### 52. `convertInterval_to_R1C1`
1. **Nome completo**: `convertInterval_to_R1C1`
2. **Módulo**: TK_Functions.bas (linha 45)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Aparenta ser um experimento/scratch de desenvolvedor para testar a API `Application.ConvertFormula` (conversão A1↔R1C1) usando uma fórmula de exemplo hardcoded — não tem propósito de negócio identificável.
6. **Quem chama**: **Nenhuma chamada encontrada.** Código morto/experimental.
7. **Procedimentos chamados**: nenhum (usa `Application.ConvertFormula`).
8. **Parâmetros**: nenhum.
9. **Retorno**: Variant — resultado de `ConvertFormula` sobre a fórmula literal `"=EC5+H6"`.
10. **Variáveis/objetos relevantes**: nenhuma (tudo hardcoded).
11. **Abas/intervalos acessados**: nenhuma.
12. **Pré-condições**: N/A.
13. **Passos principais**: chama `Application.ConvertFormula` com fórmula fixa de exemplo.
14. **Pós-condições**: N/A.
15. **Efeitos colaterais/riscos/evidência**: **código de teste/experimento deixado em produção**, sem relação com nenhuma regra de negócio real — candidato a remoção na reescrita. Evidência: linhas 45-49.

---

#### 53. `set_formula_CC`
1. **Nome completo**: `set_formula_CC`
2. **Módulo**: TK_Functions.bas (linha 51)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: No painel de validação "Painel_DM", gera uma coluna booleana indicando se cada Classe de Custo (CC) do painel existe na coluna W (23) da Base, via fórmula `COUNTIF`.
6. **Quem chama**: `Aux_Formulas_Base.bas:831` (dentro de `Copiar_Base_Origem`, quando `Col_Destino = "CLASSE CUSTO"`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `lin_ini As Integer`, `col_target As Integer`, `Sh_Destino_DM As Worksheet` (Optional, ByVal).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `x` (contagem de linhas via `Selection.End(xlDown)`).
11. **Abas/intervalos acessados**: `Sh_Destino_DM` ("Painel_DM"), coluna `col_target` (escrita) e `col_target-1` (leitura, CC); fórmula referencia `Base!C23` (coluna W da Base, hardcoded por número).
12. **Pré-condições**: `Sh_Destino_DM` deve ter dados na coluna `col_target-1` a partir de `lin_ini`.
13. **Passos principais**: usa `Select`/`Selection.End(xlDown)` para descobrir a extensão dos dados; aplica `FormulaR1C1 = "=IF(COUNTIF(Base!C23,Painel_DM!RC[-1])>0,true,false)"`; calcula; converte em valor.
14. **Pós-condições**: coluna booleana de validação de CC preenchida no Painel_DM.
15. **Efeitos colaterais/riscos/evidência**: padrão `Select`/`Selection` (estilo "macro gravada") em vez de referência direta a `Range` — mais lento e frágil a mudança de seleção ativa; referência de coluna da Base **hardcoded por número** (`C23`) em vez de nome — se a Base for reestruturada (inserir coluna antes da 23ª), esta fórmula aponta silenciosamente para a coluna errada. Evidência: linhas 51-66.

---

#### 54. `set_formula_CDC`
1. **Nome completo**: `set_formula_CDC`
2. **Módulo**: TK_Functions.bas (linha 68)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Mesmo padrão de `set_formula_CC`, mas para Centro de Custo (CDC), validando existência na coluna 35 (AI) da Base.
6. **Quem chama**: `Aux_Formulas_Base.bas:839` (dentro de `Copiar_Base_Origem`, quando `Col_Destino = "CENTRO CUSTO"`). Também "chamada" de forma **quebrada** por `Teste()` (`TK_Functions.bas:104`: `Call set_formula_CDC(6, 18106, 5)`), passando o literal `5` no lugar do parâmetro `Sh_Destino_DM As Worksheet` — isso causaria erro de tipo (`Type Mismatch`) em tempo de execução se `Teste()` fosse rodada.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `lin_ini As Integer`, `col_target As Integer`, `Sh_Destino_DM As Worksheet` (Optional, ByVal).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: idem `set_formula_CC`.
11. **Abas/intervalos acessados**: `Sh_Destino_DM` ("Painel_DM"); fórmula referencia `Base!C35` (coluna AI, hardcoded por número).
12. **Pré-condições**: idem `set_formula_CC`.
13. **Passos principais**: idêntico a `set_formula_CC`, trocando `Base!C23` por `Base!C35`.
14. **Pós-condições**: coluna booleana de validação de CDC preenchida.
15. **Efeitos colaterais/riscos/evidência**: mesma fragilidade de referência hardcoded por número de coluna; **evidência concreta de código de teste quebrado deixado em produção** (`Teste()` chamaria esta Sub com argumento de tipo incompatível). Evidência: linhas 68-83, chamada quebrada em 102-106.

---

#### 55. `set_formula_CDC_Parte_2`
1. **Nome completo**: `set_formula_CDC_Parte_2`
2. **Módulo**: TK_Functions.bas (linha 85)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Complementa `set_formula_CDC` — resolve, para cada CDC do Painel_DM, a que "grupo" da matriz DropComb ele pertence, via combinação `HLOOKUP`+`VLOOKUP`.
6. **Quem chama**: `Aux_Formulas_Base.bas:840` (dentro de `Copiar_Base_Origem`, mesma condição de `Col_Destino = "CENTRO CUSTO"`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `lin_ini As Integer`, `col_target As Integer`, `Sh_Destino_DM As Worksheet` (Optional, ByVal).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: idem, mais a fórmula `IFERROR(HLOOKUP(VLOOKUP(RC[-2],Sup_Linhas!C22:C23,2,0),DropComb!R2C6:R3C19,2,0),"-")`.
11. **Abas/intervalos acessados**: `Sh_Destino_DM`, `Sheets("Sup_Linhas")` (colunas 22-23, por número), `Sheets("DropComb")` (bloco fixo `R2C6:R3C19`).
12. **Pré-condições**: estrutura de DropComb no bloco fixo referenciado deve estar íntegra.
13. **Passos principais**: idêntico ao padrão das duas Subs anteriores, com a fórmula composta VLOOKUP→HLOOKUP.
14. **Pós-condições**: coluna de "grupo DropComb" do CDC preenchida no Painel_DM.
15. **Efeitos colaterais/riscos/evidência**: referência de **intervalo fixo** `DropComb!R2C6:R3C19` (não localizado dinamicamente por nome) — se a matriz DropComb crescer além da coluna 19 ou mudar de linha, a fórmula não acompanha automaticamente. Evidência: linhas 85-100.

---

#### 56. `Teste`
1. **Nome completo**: `Teste`
2. **Módulo**: TK_Functions.bas (linha 102)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Sub de teste/depuração de desenvolvedor para exercitar `set_formula_CDC` manualmente.
6. **Quem chama**: **Nenhuma chamada encontrada** — executada manualmente via VBE (F5).
7. **Procedimentos chamados**: `set_formula_CDC(6, 18106, 5)`.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: indiretamente via `set_formula_CDC` (mas a chamada quebraria antes de acessar dado algum).
12. **Pré-condições**: N/A.
13. **Passos principais**: uma única chamada.
14. **Pós-condições**: **erro de execução** (`Type Mismatch`) esperado, pois `set_formula_CDC` espera `Worksheet` no 3º parâmetro e recebe o literal `5`.
15. **Efeitos colaterais/riscos/evidência**: **código de teste quebrado deixado em produção** — confirma falta de limpeza/revisão de código antes de versionar o módulo; risco baixo isolado (não é chamado por ninguém), mas é um sintoma de qualidade de processo. Evidência: linhas 102-106.

---

#### 57. `getCellAddress`
1. **Nome completo**: `getCellAddress`
2. **Módulo**: TK_Functions.bas (linha 109)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Utilitário trivial que retorna o endereço A1 de uma célula dado nome de planilha, linha e coluna.
6. **Quem chama**: **Nenhuma chamada ativa** — referenciada apenas dentro de código **comentado** em `UPDATE_Combinacoes_Empresas` (linhas 184, 186), como parte de uma abordagem alternativa (célula a célula) que foi substituída pela abordagem baseada em fórmula (`Reclassificar_Combinacoes_Empresas_TK`).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `sheetName`, `linha`, `coluna` (sem tipo).
9. **Retorno**: String — endereço A1 da célula.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: a planilha nomeada por `sheetName`.
12. **Pré-condições**: `sheetName` deve existir.
13. **Passos principais**: `Sheets(sheetName).Cells(linha, coluna).Address`.
14. **Pós-condições**: N/A.
15. **Efeitos colaterais/riscos/evidência**: código morto — evidência histórica de uma abordagem célula-a-célula (mais lenta) que foi abandonada em favor da abordagem por fórmula em massa; mantida como comentário em vez de removida. Evidência: linhas 109-113, referências comentadas 184-186.

---

#### 58. `UPDATE_Combinacoes_Empresas`
1. **Nome completo**: `UPDATE_Combinacoes_Empresas`
2. **Módulo**: TK_Functions.bas (linha 115)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Sincroniza a lista de combinações de Empresas definida na planilha "DropComb" para dentro do cabeçalho (linha 5) da Base, inserindo/removendo colunas de "slot" conforme a quantidade de combinações mudou, e dispara a reclassificação (`_TK`) apenas se algo de fato mudou.
6. **Quem chama**: `Sheet8.cls:15` — acionada pelo evento `ComboBox1_Change` da aba "Extracao" quando o usuário seleciona a opção contendo o texto "Combinações" no combo.
7. **Procedimentos chamados**: `Reclassificar_Combinacoes_Empresas_TK` (linha 180, condicional).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `linhamaxima` (extensão de DropComb!D3:D100), `qtde_Combinacoes_Empresas`, `qtde_slots` (contagem de colunas "EMPRESA" já existentes no cabeçalho da Base, via `CountIf`), `diff_slots`, `vColuna=133` (**coluna inicial hardcoded** — coluna EC).
11. **Abas/intervalos acessados**: `Sheets("DropComb")` (`D3:D100`, `F2:R2`), `Sheets("Base")` (linha 5, a partir da coluna 133/"EC"; `EC6` e range dependente para checagem de células em branco).
12. **Pré-condições**: usuário deve confirmar `MsgBox vbYesNo` de "Quer prosseguir com essa ação?".
13. **Passos principais**:
    - `MsgBox` de confirmação.
    - Calcula quantos slots de combinação existem hoje na Base vs. quantos são necessários (`diff_slots`).
    - Se precisa de mais slots: insere colunas (`Columns(vColuna).Copy` + `Insert Shift:=xlToRight`) a partir da coluna 133 (EC).
    - Copia os nomes de combinação de DropComb para a linha 5 da Base, célula a célula.
    - Verifica se as células de dados abaixo do cabeçalho estão em branco (`CountBlank` vs. contagem total).
    - Se houve mudança real (`qtde_cells_blank <> qtde_cells_selected`): chama `Reclassificar_Combinacoes_Empresas_TK`.
    - `MsgBox` de sucesso; seleciona `Sheets("Extracao")`.
14. **Pós-condições**: cabeçalho da Base sincronizado com DropComb; reclassificação `_TK` recalculada se necessário.
15. **Efeitos colaterais/riscos/evidência**: **coluna inicial hardcoded (133 = "EC")** — se a estrutura da Base mudar de forma que as colunas de combinação não comecem mais em EC, esta rotina insere dados no lugar errado silenciosamente; a heurística de "mudou ou não" (comparação de células em branco antes/depois) é frágil — não detecta o caso em que o número de combinações é o mesmo mas os *nomes/critérios* mudaram (nesse caso a reclassificação não seria disparada, deixando a Base com classificação desatualizada mas sem nenhum aviso). Evidência: linhas 115-204, coluna hardcoded linha 145.

---

#### 59. `UPDATE_Combinacoes_Proforma`
1. **Nome completo**: `UPDATE_Combinacoes_Proforma`
2. **Módulo**: TK_Functions.bas (linha 206)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Mesmo padrão de `UPDATE_Combinacoes_Empresas`, mas para a matriz de combinações "Proforma" — sem inserção/remoção de colunas (apenas limpa e repreenche o intervalo existente a partir da coluna "PROFORMA" localizada dinamicamente).
6. **Quem chama**: **Nenhuma chamada ativa encontrada.** A única referência é um `Call` **comentado** dentro de `UPDATE_Combinacoes_Empresas` (linha 196: `'Call UPDATE_Combinacoes_Proforma`) — ou seja, apesar de totalmente implementada, **esta Sub não está conectada a nenhum ponto de entrada ativo** (nem botão, nem chamada de outra rotina). Diferente de `UPDATE_Combinacoes_Empresas`, que É acionada pelo ComboBox da aba Extracao — o ComboBox aparentemente só tem as opções "Defined Names" e "Combinações" (Empresas), sem opção equivalente para Proforma.
7. **Procedimentos chamados**: `Reclassificar_Combinacoes_Proforma_TK` (linha 267, condicional).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `linhamaxima` (DropComb!BH3:BH100), `y` (letra da coluna "PROFORMA" na Base, localizada via `Match`).
11. **Abas/intervalos acessados**: `Sheets("DropComb")` (`BH3:BH100`, `BI3:BI<linhamaxima>`, `BK2:CK2`), `Sheets("Base")` (coluna "PROFORMA" localizada por `Match`, a partir da linha 5).
12. **Pré-condições**: `answer = vbYes` está **hardcoded** na linha 212 (sem `MsgBox` de confirmação real, ao contrário de `UPDATE_Combinacoes_Empresas`) — a Sub sempre prossegue como se o usuário tivesse confirmado.
13. **Passos principais**: limpa o range de combinações Proforma existente; copia valores de DropComb para a linha 5, pulando separadores visuais (`"-------..."`); checa mudança via `CountBlank`; se mudou, chama `Reclassificar_Combinacoes_Proforma_TK`; seleciona "Extracao".
14. **Pós-condições**: cabeçalho Proforma da Base sincronizado (se executada).
15. **Efeitos colaterais/riscos/evidência**: **procedimento órfão** — totalmente funcional mas sem nenhum caminho de invocação ativo encontrado, sugerindo funcionalidade planejada e implementada mas nunca conectada à UI (ou desconectada em uma refatoração posterior do ComboBox); `answer = vbYes` hardcoded é resquício de um `MsgBox` de confirmação removido sem limpar a variável — o "Else" que mostraria `"Nada foi feito!"` (linha 279) é código morto inalcançável, já que `answer` nunca é diferente de `vbYes`. Evidência: linhas 206-282, hardcode em 212.

---

#### 60. `GET_value_S_or_N`
1. **Nome completo**: `GET_value_S_or_N`
2. **Módulo**: TK_Functions.bas (linha 284)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Abordagem alternativa (célula a célula, por correspondência de substring) para determinar se uma combinação Empresa×Centro-de-Custo deveria ser "S" ou "N", consultando diretamente a matriz DropComb — abandonada em favor da abordagem por fórmula em massa.
6. **Quem chama**: **Nenhuma chamada ativa** — referenciada apenas em código **comentado** dentro de `UPDATE_Combinacoes_Empresas` (linha 186).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `var1`, `var2` (endereços de célula como string), `linhaMatriz`, `colunaMatriz` (ranges de referência).
9. **Retorno**: String — "S", "N" ou "-" (se `p1="-"`).
10. **Variáveis/objetos relevantes**: `match1`/`match2` (posições encontradas via `InStr` em vez de `Match` exato).
11. **Abas/intervalos acessados**: `Sheets("Base")` (leitura dos valores em `var1`/`var2`), `Sheets("DropComb")` (implícito via `colunaMatriz`/`linhaMatriz` passados por parâmetro).
12. **Pré-condições**: N/A (não executado).
13. **Passos principais**: busca `p2` dentro de cada célula de `linhaMatriz` via `InStr` (substring); busca `p1` dentro de `colunaMatriz`; lê a célula de interseção em DropComb; retorna "S" se o valor for exatamente "=", senão "N".
14. **Pós-condições**: N/A (código morto).
15. **Efeitos colaterais/riscos/evidência**: **correspondência por substring (`InStr`)** em vez de igualdade exata — risco de falso positivo (ex.: "TIM" casando com "TIM FIXO") caso esta função venha a ser reativada; código morto confirmado (só aparece comentado). Evidência: linhas 284-313.

---

#### 61. `UPDATE_Combinacoes_Empresas_parte2`
1. **Nome completo**: `UPDATE_Combinacoes_Empresas_parte2`
2. **Módulo**: TK_Functions.bas (linha 315)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Preenche por `AutoFill` as fórmulas de combinação de Empresa em todo o range `EC6:GD40837` da Base, calcula e converte em valores — parece ser um passo manual complementar/alternativo às reclassificações automáticas, com extensão de linhas **totalmente hardcoded**.
6. **Quem chama**: **Nenhuma chamada encontrada em todo o dump** — nem por outro procedimento, nem por comentário. Presumivelmente executada manualmente pelo desenvolvedor via VBE quando necessário. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma — todos os endereços são literais de string (`"EC6:GD6"`, `"EC6:GD40837"`, `"EC7"`, `"ED6"`).
11. **Abas/intervalos acessados**: `Sheets("Base")`, range fixo `EC6:GD40837` (colunas de combinação de Empresa) e `ED6` em diante.
12. **Pré-condições**: a fórmula-modelo já deve existir em `EC6` antes de rodar (a Sub só propaga via `AutoFill`, não cria a fórmula original).
13. **Passos principais**: `AutoFill` de `EC6` para `EC6:GD6`; `AutoFill` de `EC6:GD6` para `EC6:GD40837`; `.Calculate`; copia/cola como valor a partir de `EC7`; repete para `ED6` em diante.
14. **Pós-condições**: bloco `EC6:GD40837` da Base preenchido com valores fixos de classificação de Empresa.
15. **Efeitos colaterais/riscos/evidência**: **limite de linha hardcoded em 40.837** — se a Base (que já é descrita em outras partes do sistema como tendo dezenas de milhares de linhas) crescer além dessa marca, esta rotina simplesmente não preenche as linhas excedentes, um **bug silencioso de dados incompletos** sem nenhum erro ou aviso ao usuário; classificado como risco crítico na seção D (item D-2). Evidência: linhas 315-340, limite hardcoded na linha 321.

---

#### 62. `CLEAR_Defined_Names`
1. **Nome completo**: `CLEAR_Defined_Names`
2. **Módulo**: TK_Functions.bas (linha 342)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Enumera todos os "Defined Names" (intervalos nomeados) do workbook, classifica cada um por tipo de referência (link de rede, SharePoint, disco local, arquivo externo, referência quebrada `#REF!`/`#N/A`/`#NAME?`) e grava a lista na planilha `ListDefinedNames` — ferramenta de diagnóstico contra o acúmulo patológico de nomes definidos, causa comum de lentidão/corrupção em arquivos `.xlsb` legados.
6. **Quem chama**: `Sheet8.cls:13` — acionada pelo evento `ComboBox1_Change` da aba "Extracao" quando o usuário seleciona a opção contendo "Defined Names".
7. **Procedimentos chamados**: nenhum (usa apenas `ActiveWorkbook.Names`, `Instr`, `Mid`).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `DotNetArray` (ArrayList, instanciado mas nunca populado no caminho ativo — o bloco que o preenche está comentado, linhas 419-433); classificação via `InStr` de padrões: `'\\` (rede não-P&C), `https://timbrasil-my.sharepoint` (SharePoint), `F:\P&C` (rede P&C), `:\` (disco local), `'[` (arquivo externo mesma pasta), mais `#REF!`/`#N/A`/`#NAME?`.
11. **Abas/intervalos acessados**: `Sheets("ListDefinedNames")` (torna visível, limpa `A2:E99999`, repovoa a partir de `A1`).
12. **Pré-condições**: `MsgBox vbYesNo` de confirmação.
13. **Passos principais**: confirmação do usuário; `Application.Calculation = xlCalculationManual`; itera `ActiveWorkbook.Names` (pulando nomes com "xlfn" ou "!" no nome); para cada nome, grava índice/nome/referência simplificada/classificação; `MsgBox` de sucesso; restaura `ScreenUpdating`/`Calculation`.
14. **Pós-condições**: planilha `ListDefinedNames` populada com o diagnóstico completo de todos os nomes definidos do workbook.
15. **Efeitos colaterais/riscos/evidência**: **apenas lista — não deleta** (a lógica de deleção real, incluindo o preenchimento de `DotNetArray`, está inteiramente comentada, linhas 419-445) — a deleção efetiva é feita por `RUN_Apagar_defined_names_definitivamente` (item 63), separadamente, após o usuário marcar manualmente quais nomes deletar na planilha resultante; ou seja, o fluxo de limpeza de Defined Names é **dividido em 2 etapas manuais** (listar aqui, depois marcar "Sim"/"Não" na planilha, depois rodar a outra rotina) — processo propenso a erro humano se a planilha for editada incorretamente entre as etapas. Evidência: linhas 342-456, lógica de deleção comentada 419-445.

---

#### 63. `RUN_Apagar_defined_names_definitivamente`
1. **Nome completo**: `RUN_Apagar_defined_names_definitivamente`
2. **Módulo**: TK_Functions.bas (linha 458)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Executa a exclusão definitiva dos nomes definidos que o usuário marcou com "Sim" (coluna E) na planilha gerada por `CLEAR_Defined_Names` — segunda etapa do processo de limpeza.
6. **Quem chama**: **Nenhuma chamada encontrada** — presumivelmente executada manualmente via VBE após revisão da planilha `Sheet29`. `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum (usa `ActiveWorkbook.Names(...).Delete`).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Sh29=Sheet29`; `possuiAspasSimples` (flag para tratar nomes que contêm `'!`, exigindo aspas simples no início para `Names(...)` resolver corretamente).
11. **Abas/intervalos acessados**: `Sheet29` (colunas 2="Nome" e 5="Sim/Não", conforme layout aparentemente produzido por `CLEAR_Defined_Names` — mas gravado em `Sheets("ListDefinedNames")`, não explicitamente em "Sheet29"; ligação entre os dois **não confirmada diretamente no texto**, `[VALIDAR COM O NEGÓCIO]`).
12. **Pré-condições**: coluna E de `Sh29` deve ter sido preenchida manualmente pelo usuário com "Sim" para os nomes a excluir.
13. **Passos principais**: laço por todas as linhas de `Sh29`; se coluna E = "Sim" e o nome não contém "xlfn"; detecta se o nome tem `'!` embutido (exigindo prefixo `'`); executa `ActiveWorkbook.Names(vName).Delete`; conta exclusões.
14. **Pós-condições**: nomes marcados são removidos do workbook; `MsgBox` com contagem de nomes deletados; `Sh29.Visible = False`.
15. **Efeitos colaterais/riscos/evidência**: **operação destrutiva e irreversível** (exclusão de Defined Names) sem `MsgBox` de confirmação final antes de rodar (a confirmação, se existir, é apenas a decisão prévia de marcar "Sim" na planilha) — um erro de preenchimento da coluna E apaga nomes por engano; exclui `ActiveWorkbook.Names(vName)` sem `On Error` — se um nome já tiver sido removido por outro processo entre a listagem e esta execução, a Sub quebra com erro não tratado no meio do laço, deixando a exclusão parcial. Evidência: linhas 458-497.

---

#### 64. `Testeee`
1. **Nome completo**: `Testeee`
2. **Módulo**: TK_Functions.bas (linha 499)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Utilitário de depuração que percorre todas as planilhas do workbook e lista, na janela Immediate do VBE (`Debug.Print`), todas as células com validação de dados configurada (tipo, sheet, endereço).
6. **Quem chama**: **Nenhuma chamada encontrada** — ferramenta de inspeção manual, uso exclusivo do desenvolvedor via VBE.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `sh` (Worksheet), `cell` (Range) — laço duplo sobre todas as planilhas × todas as células com `SpecialCells(xlCellTypeAllValidation)`.
11. **Abas/intervalos acessados**: **todas** as planilhas do workbook (`ThisWorkbook.Worksheets`).
12. **Pré-condições**: nenhuma.
13. **Passos principais**: laço aninhado planilha→célula com validação; `Debug.Print` do endereço, tipo de validação e nome da planilha.
14. **Pós-condições**: N/A (saída só na janela Immediate, não persistida).
15. **Efeitos colaterais/riscos/evidência**: `SpecialCells(xlCellTypeAllValidation)` lança erro se a planilha não tiver nenhuma célula com validação — não há `On Error Resume Next` ao redor, então a Sub pode abortar no meio da varredura de todas as planilhas dependendo do conteúdo; nome "Testeee" (com 3 "e"s) reforça a natureza de rascunho/debug nunca limpo. Evidência: linhas 499-513.

---

#### 65. `RUN_Atualizar_Dinamica_com_erros`
1. **Nome completo**: `RUN_Atualizar_Dinamica_com_erros`
2. **Módulo**: TK_Functions.bas (linha 515)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Atualiza (`PivotCache.Refresh`) a tabela dinâmica "PivotTable1" localizada em `Sheet29` — provavelmente um painel que resume os erros logados (por `fn_ListAllErrors`, embora essa conexão não seja direta) ou os Defined Names classificados.
6. **Quem chama**: **Nenhuma chamada encontrada.** `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: `Sheet29` (seleciona `G8`), `PivotTables("PivotTable1")` nessa mesma planilha.
12. **Pré-condições**: `Sheet29` deve conter uma tabela dinâmica nomeada "PivotTable1".
13. **Passos principais**: seleciona Sheet29 e a célula G8; `ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh`.
14. **Pós-condições**: pivot atualizado com dados correntes da fonte associada.
15. **Efeitos colaterais/riscos/evidência**: o nome sugere relação com "erros", mas o código não faz nenhuma referência explícita a `tk_Lista_de_erros` — a fonte de dados real do pivot **não é identificável apenas pelo texto VBA** (definida na configuração do PivotCache, fora do escopo de texto), `[NÃO ACESSÍVEL]`/`[VALIDAR COM O NEGÓCIO]`. Evidência: linhas 515-519.

---

#### 66. `UPDATE_aplicar_CDC_por_Referencia`
1. **Nome completo**: `UPDATE_aplicar_CDC_por_Referencia`
2. **Módulo**: TK_Functions.bas (linha 521)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Preenche automaticamente o Centro de Custo (CDC) ausente ("-") de cada linha da Base, buscando uma correspondência em uma tabela DE-PARA alimentada via SQL Server, e marca visualmente (cor de fundo) a origem de cada valor de CDC (original, por referência, ou ainda ausente) — mecanismo de fallback de qualidade de dado.
6. **Quem chama**: **Nenhuma chamada encontrada.** `[INFERÊNCIA]` botão na aba Extracao (mesmo padrão de outras rotinas de "primeiro nível" deste cluster). `[NÃO ACESSÍVEL]`.
7. **Procedimentos chamados**: `Desligar_Tudo`, `Extrair_Base_CDCs_DE_PARA`, `SET_Cor_CdC` (3x, condicional), `PopUp_Tempo_Processamento`, `Ativar_Tudo`.
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `rngCDC` (Range da coluna CENTRO CUSTO); `vChave` (concatenação de colunas 31, 32, 33 e 8 — "AE_AF_AG_H"); `Sh11=Sheet11` (staging DE-PARA).
11. **Abas/intervalos acessados**: `Sheet3` ("Base", colunas "CENTRO CUSTO" e 31/32/33/8), `Sheet11` (coluna 50 em diante — DE-PARA de CDC, populada por `Extrair_Base_CDCs_DE_PARA`).
12. **Pré-condições**: deve haver dados na Base (`A6:A999999` não vazio, senão `MsgBox` e aborta).
13. **Passos principais**:
    - Se Base vazia: `MsgBox` e sai.
    - Senão: `Inicio=Now`; `Desligar_Tudo`; torna `Sh11` visível; localiza coluna CDC; chama `Extrair_Base_CDCs_DE_PARA` (popula DE-PARA via SQL).
    - Laço `For linha = 6 To qtde`: se CDC = "-", monta `vChave` (colunas 31+32+33+8), busca em `Sh11` col. 50; se não encontrado, `SET_Cor_CdC("sem CDC")`; se encontrado e CDC ainda "-", preenche com o valor de referência e `SET_Cor_CdC("com CDC por referencia")`; se já tinha valor, `SET_Cor_CdC("com CDC Real")`.
    - `PopUp_Tempo_Processamento`; `Ativar_Tudo`.
14. **Pós-condições**: coluna CDC da Base preenchida onde possível; células coloridas conforme origem do dado (evidência visual persistente na planilha).
15. **Efeitos colaterais/riscos/evidência**: laço `For linha = 6 To qtde` com `Application.Match` individual por linha (não vetorizado) — para dezenas de milhares de linhas, é uma busca O(n) por linha, potencialmente lento; colore células via `Select` + `Selection.Interior` (padrão "macro gravada" — lento e sujeito a interferência se o usuário interagir com a planilha durante a execução); as cores usadas são hardcoded por código numérico RGB (`Color = 5287936` "com CDC Real", `Color = 49407` "com CDC por referencia") sem constantes nomeadas, dificultando manutenção. Evidência: linhas 521-574.

---

#### 67. `FiltraCampo`
1. **Nome completo**: `FiltraCampo`
2. **Módulo**: TK_Functions.bas (linha 577)
3. **Tipo**: Function
4. **Escopo**: Public
5. **Objetivo**: Utilitário genérico para aplicar um filtro simples (`Campo = 'valor'`) sobre um `ADODB.Recordset` já aberto.
6. **Quem chama**: **Nenhuma chamada encontrada em todo o dump.** Código órfão — pode ter sido usado por uma versão anterior de alguma extração SQL, ou destinado a uso futuro nunca concretizado.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `rstTemp As ADODB.Recordset`, `strCampo As String`, `strcriterio As String`.
9. **Retorno**: `ADODB.Recordset` — o próprio recordset com o filtro aplicado.
10. **Variáveis/objetos relevantes**: nenhuma além dos parâmetros.
11. **Abas/intervalos acessados**: nenhuma planilha — opera sobre um Recordset em memória.
12. **Pré-condições**: `rstTemp` deve estar aberto.
13. **Passos principais**: seta `rstTemp.Filter = strCampo & " = '" & strcriterio & "'"`; retorna o recordset.
14. **Pós-condições**: recordset filtrado em memória.
15. **Efeitos colaterais/riscos/evidência**: concatenação direta de `strcriterio` na cláusula de filtro sem *escaping* de aspas simples — se `strcriterio` contiver um apóstrofo (comum em nomes compostos), o filtro quebra; como não há chamador ativo, o risco é hoje teórico, mas relevante se a função for resgatada no futuro. Evidência: linhas 577-583.

---

#### 68. `Extrair_Base_CDCs_DE_PARA`
1. **Nome completo**: `Extrair_Base_CDCs_DE_PARA`
2. **Módulo**: TK_Functions.bas (linha 585)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Consulta o SQL Server (view `VW_HUBBLE_QUICKDATA_CDC_REFERENCIA`) e carrega a tabela DE-PARA de Centro de Custo na planilha de staging `Sheet11`, usada por `UPDATE_aplicar_CDC_por_Referencia` para preencher CDCs ausentes.
6. **Quem chama**: `TK_Functions.bas:543` — chamada interna, dentro de `UPDATE_aplicar_CDC_por_Referencia`.
7. **Procedimentos chamados**: `AbreConexao`, `FechaConexao` (Conexoes.bas, fora do cluster).
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído (efeito via escrita em `Sh11`).
10. **Variáveis/objetos relevantes**: `Info` (ADODB.Recordset), `Comando` (ADODB.Command, `CommandTimeout=1000`); SQL literal: `"SELECT A.* FROM [BPAM].[dbo].[VW_HUBBLE_QUICKDATA_CDC_REFERENCIA] A"`.
11. **Abas/intervalos acessados**: `Sh11=Sheet11`, colunas 50-51 (limpa dados anteriores, depois `CopyFromRecordset` a partir de `Cells(2,50)`).
12. **Pré-condições**: conectividade SQL Server (BPAM).
13. **Passos principais**: abre conexão; monta e executa comando SQL (`SELECT *` da view de referência); limpa staging anterior em `Sh11` se houver; `Sh11.Cells(2,50).CopyFromRecordset Info`; fecha conexão.
14. **Pós-condições**: `Sh11` (colunas 50-51) contém a tabela DE-PARA de CDC completa, pronta para lookup.
15. **Efeitos colaterais/riscos/evidência**: **tratamento de erro inativo** — a linha `' On Error GoTo tratar_erro` está **comentada** (linha ~602), tornando o rótulo `tratar_erro` (linhas 637-638) código morto/inalcançável; se a query SQL falhar, o erro não é tratado por este bloco (propaga como erro VBA padrão, interrompendo a execução sem a mensagem amigável planejada); a mensagem de erro no bloco morto referencia a variável `Valor`, que **não é definida em nenhum lugar dentro desta Function** — mesmo que o `On Error` fosse reativado, a mensagem exibiria um valor vazio/incorreto (`"Ocorreu um erro ao atualizar " & Valor & "..."`); `SELECT A.*` sem filtro — traz a view inteira, potencialmente cara se a view for grande. Evidência: linhas 585-645, tratamento de erro comentado ~602, variável `Valor` não definida ~638.

---

#### 69. `SET_Cor_CdC`
1. **Nome completo**: `SET_Cor_CdC`
2. **Módulo**: TK_Functions.bas (linha 647)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Aplica cor de preenchimento à célula atualmente selecionada, conforme a origem do valor de CDC ("sem CDC", "com CDC Real", "com CDC por referencia") — codificação visual usada por `UPDATE_aplicar_CDC_por_Referencia`.
6. **Quem chama**: `TK_Functions.bas:557,561,563` — chamadas internas, dentro de `UPDATE_aplicar_CDC_por_Referencia`.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: `flagCDC` (string identificando o caso).
9. **Retorno**: Function sem valor de retorno nomeado atribuído (efeito via `Selection.Interior`).
10. **Variáveis/objetos relevantes**: nenhuma além do parâmetro.
11. **Abas/intervalos acessados**: nenhuma explicitamente — opera sobre `Selection` (célula ativa no momento da chamada, definida pelo chamador via `.Select` antes de invocar esta Function).
12. **Pré-condições**: o chamador deve ter selecionado a célula-alvo antes de chamar esta Function (acoplamento implícito via estado global `Selection`).
13. **Passos principais**: `Select Case` implícito via `If`/`ElseIf` sobre `flagCDC`; aplica `Interior.Color`/`ThemeColor` conforme o caso ("sem CDC" = `ThemeColor xlThemeColorDark1`; "com CDC Real" = `Color 5287936`; "com CDC por referencia" = `Color 49407`).
14. **Pós-condições**: célula selecionada colorida.
15. **Efeitos colaterais/riscos/evidência**: **depende inteiramente do estado global `Selection`** em vez de receber um `Range` por parâmetro — padrão frágil e difícil de testar isoladamente (a função não sabe *qual* célula está colorindo, apenas confia que o chamador selecionou a certa); se `flagCDC` não corresponder a nenhum dos 3 casos, nada acontece silenciosamente (sem `Else`/erro). Evidência: linhas 647-674.

---

#### 70. `SET_Limpar_Cores_CDC`
1. **Nome completo**: `SET_Limpar_Cores_CDC`
2. **Módulo**: TK_Functions.bas (linha 676)
3. **Tipo**: Function
4. **Escopo**: Public (implícito)
5. **Objetivo**: Remove a formatação de cor de fundo da seleção atual — contraparte de limpeza de `SET_Cor_CdC`.
6. **Quem chama**: **Nenhuma chamada encontrada** — código órfão (não é usado nem por `UPDATE_aplicar_CDC_por_Referencia`, que aplica cores mas nunca as limpa via esta função).
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: nenhuma.
11. **Abas/intervalos acessados**: `Selection` (mesmo acoplamento de `SET_Cor_CdC`).
12. **Pré-condições**: usuário/chamador deve ter uma seleção ativa.
13. **Passos principais**: `Selection.Interior.Pattern = xlNone` (e reset de tint/shade).
14. **Pós-condições**: formatação de cor removida da seleção.
15. **Efeitos colaterais/riscos/evidência**: órfã — implica que, uma vez que `UPDATE_aplicar_CDC_por_Referencia` colore células de status de CDC, **não existe rotina automatizada para limpar essas cores** antes de uma nova rodada; se a Base for reprocessada, células antigas "sem CDC" (coloridas) que passam a ter CDC podem manter a cor antiga incorretamente, a menos que a nova cor sobrescreva (o que `SET_Cor_CdC` de fato faz, já que também usa `Pattern=xlSolid`) — risco baixo na prática, mas a existência de uma função de limpeza nunca chamada é indício de fluxo de manutenção incompleto. Evidência: linhas 676-684.

---

#### 71. `Calcular_Comb_Meses_Intervalo_Linha_TK`
1. **Nome completo**: `Calcular_Comb_Meses_Intervalo_Linha_TK`
2. **Módulo**: TK_Functions.bas (linha 686)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Versão `_TK` de `Calcular_Comb_Meses_Intervalo_Linha` (Auxiliar.bas, item 36) — recalcula a fórmula de combinação de meses (`INDIRECT`), mas **sempre para o intervalo fixo `Lin_Inicial=6` até `qtdeLinhasBase+5`**, em vez de receber o intervalo como parâmetro.
6. **Quem chama**: `fx_IFRS16.bas:126` — chamada como parte do pós-processamento de tratamento IFRS16 (`UPDATE_Tratar_IFRS16`), **após** o Empresa/IFRS_Contabil terem sido recalculados para a Base inteira.
7. **Procedimentos chamados**: nenhum.
8. **Parâmetros**: nenhum (diferente da versão em Auxiliar.bas, que recebe `Lin_Inicial`/`Lin_Final`).
9. **Retorno**: N/A.
10. **Variáveis/objetos relevantes**: `Lin_Inicial=6` (hardcoded), `Lin_Final = qtdeLinhasBase+5` (calculado via `CountA` de `A6:A999999`).
11. **Abas/intervalos acessados**: `Sheet3` — bloco de colunas de combinação de meses.
12. **Pré-condições**: assume que a linha 6 é sempre o início dos dados (cabeçalho na linha 5) — **premissa de layout fixo**, diferente da busca dinâmica por "LIN_BASE" usada na versão de `Auxiliar.bas`.
13. **Passos principais**: mesma lógica de `Calcular_Comb_Meses_Intervalo_Linha` (fórmula `SUM(INDIRECT(...))`, bake-in via `.Copy`/`PasteSpecial`), mas com o intervalo de linhas sempre recalculado do zero (linha 6 até a última linha com dado na coluna A).
14. **Pós-condições**: **toda** a coluna de combinação de meses da Base recalculada (não apenas um subconjunto, ao contrário da versão parametrizada de Auxiliar.bas).
15. **Efeitos colaterais/riscos/evidência**: **DUPLICAÇÃO CONFIRMADA COM PREMISSA DIVERGENTE** — a versão de `Auxiliar.bas:466` (`Calcular_Comb_Meses_Intervalo_Linha`) localiza o cabeçalho dinamicamente via `Match("LIN_BASE", ...)` (linha 472) e recebe `Lin_Inicial`/`Lin_Final` como parâmetros explícitos do chamador; esta versão `_TK` **assume que a linha do cabeçalho é sempre 5 e os dados começam sempre na linha 6** (`Lin_Inicial = 6`, linha 691) e recalcula sempre a Base inteira. Se a estrutura da Base for alterada (por exemplo, inserir uma linha de metadado acima do cabeçalho, deslocando "LIN_BASE" para outra posição), a versão de `Auxiliar.bas` continua funcionando corretamente (busca dinâmica) enquanto **esta versão `_TK` quebra silenciosamente** (recalcula o intervalo errado, sem erro visível). Evidência: linhas 686-714 (esta versão) vs. Auxiliar.bas linhas 466-472 (versão dinâmica) — ver também RN-070 e risco D-1.

---

#### 72. `Reclassificar_Combinacoes_IFRS_Contabil_TK`
1. **Nome completo**: `Reclassificar_Combinacoes_IFRS_Contabil_TK`
2. **Módulo**: TK_Functions.bas (linha 716)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Versão `_TK` de `Reclassificar_Combinacoes_IFRS_Contabil` (Auxiliar.bas, item 41) — mesma lógica de montagem de fórmula via `Criar_Formula_Filtros`, mas com intervalo de linhas sempre recalculado internamente (`Lin_Inicial=6` fixo) em vez de recebido por parâmetro.
6. **Quem chama**: `fx_IFRS16.bas:128` — **único chamador em todo o dump**. Diferente das outras duas reclassificações `_TK`, esta **não é chamada por nenhuma Sub `UPDATE_Combinacoes_*` dentro do próprio `TK_Functions.bas`** (não existe um `UPDATE_Combinacoes_IFRS_Contabil` equivalente no módulo) — ou seja, o único caminho de invocação desta versão `_TK` é o pós-processamento IFRS16.
7. **Procedimentos chamados**: `Criar_Formula_Filtros` (Auxiliar.bas — a única peça de lógica genuinamente compartilhada entre as duas famílias, ver item 43).
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9` ("DropComb"); `Lin_Inicial=6` (hardcoded, linha 728); `Ult_Linh` calculado via `CountA(sh.Cells(1,Col_IFRS_Contabil).EntireColumn) + Lin_Cabecalho - 2` (linha 784) — **fórmula de cálculo do fim do range diferente** da usada na versão de Auxiliar.bas (que recebe `Ult_Linh` como parâmetro do chamador).
11. **Abas/intervalos acessados**: `Sheet3` ("IFRS_Contabil"), `Sheet9` (colunas "IFRS_Contabil*").
12. **Pré-condições**: mesma premissa de layout fixo (linha 6 = início dos dados) da versão `_TK` de combinação de meses.
13. **Passos principais**: idêntico à versão de Auxiliar.bas (loop de leitura de critérios, exclusão de separadores visuais, `Criar_Formula_Filtros`, fechamento de `ELSE`), mas resolvendo `Lin_Inicial`/`Ult_Linh` internamente em vez de recebê-los.
14. **Pós-condições**: bloco de colunas "IFRS_Contabil" da Base recalculado (potencialmente para uma extensão diferente da que seria calculada pela versão de Auxiliar.bas, dependendo de onde estejam os dados).
15. **Efeitos colaterais/riscos/evidência**: **DUPLICAÇÃO CONFIRMADA** — mesma lógica de negócio (reclassificação IFRS_Contabil) implementada 2x com fórmulas de extensão de range **diferentes** (uma parametrizada e dinâmica em `Auxiliar.bas:781-857`, outra hardcoded/auto-calculada aqui) — qualquer correção na regra de negócio de classificação IFRS precisa ser replicada manualmente nos dois lugares, e o histórico do código já mostra que isso nem sempre acontece (a versão de Auxiliar.bas foi ajustada para dividir cálculo em partes/casos especiais que não têm equivalente aqui). Evidência: linhas 716-794 (esta versão) vs. Auxiliar.bas 781-857; único chamador fx_IFRS16.bas:128.

---

#### 73. `Reclassificar_Combinacoes_Proforma_TK`
1. **Nome completo**: `Reclassificar_Combinacoes_Proforma_TK`
2. **Módulo**: TK_Functions.bas (linha 797)
3. **Tipo**: Function (usada como Sub)
4. **Escopo**: Public
5. **Objetivo**: Versão `_TK` de `Reclassificar_Combinacoes_Proforma` (Auxiliar.bas, item 42), mesmo padrão de premissa de layout fixo.
6. **Quem chama**: `fx_IFRS16.bas:127` (pós-tratamento IFRS16), `TK_Functions.bas:267` (interno, dentro de `UPDATE_Combinacoes_Proforma`, condicional a mudança detectada) — **porém `UPDATE_Combinacoes_Proforma` em si não tem nenhum chamador ativo** (item 59), então na prática **o único caminho de execução real desta versão `_TK` hoje é via `fx_IFRS16.bas`**.
7. **Procedimentos chamados**: `Criar_Formula_Filtros` (Auxiliar.bas).
8. **Parâmetros**: nenhum.
9. **Retorno**: Function sem valor de retorno nomeado atribuído.
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9`; `Lin_Inicial=6` (hardcoded, linha 809); `Ult_Linh` via `CountA(...) + Lin_Cabecalho - 2` (linha 865, mesma fórmula da versão IFRS_Contabil_TK).
11. **Abas/intervalos acessados**: `Sheet3` ("Proforma"), `Sheet9` (colunas "Proforma*").
12. **Pré-condições**: idem item 72.
13. **Passos principais**: idêntico à versão de Auxiliar.bas, com `Lin_Inicial`/`Ult_Linh` resolvidos internamente.
14. **Pós-condições**: bloco de colunas "Proforma" recalculado.
15. **Efeitos colaterais/riscos/evidência**: mesma duplicação de premissa documentada no item 72, aplicada a "Proforma"; adicionalmente, `Application.StatusBar = ""` ao final (linha 873) — diferente das versões de `Auxiliar.bas`, que não limpam a `StatusBar` explicitamente ao final de cada reclassificação individual (apenas a rotina "pai" o faz) — inconsistência menor de comportamento entre as duas famílias. Evidência: linhas 797-875.

---

#### 74. `Reclassificar_Combinacoes_Empresas_TK`
1. **Nome completo**: `Reclassificar_Combinacoes_Empresas_TK`
2. **Módulo**: TK_Functions.bas (linha 877)
3. **Tipo**: Sub
4. **Escopo**: Public (implícito)
5. **Objetivo**: Versão `_TK` de `Reclassificar_Combinacoes_Empresas` (Auxiliar.bas, item 40) — **sem a divisão em 2 partes** presente na versão original (aqui a fórmula é montada e aplicada de uma vez só ao bloco inteiro de colunas "EMPRESA").
6. **Quem chama**: `fx_IFRS16.bas:125` (pós-tratamento IFRS16), `TK_Functions.bas:180` (interno, dentro de `UPDATE_Combinacoes_Empresas`, condicional a mudança detectada — este é o único caminho `_TK` que **tem** um chamador de primeiro nível ativo confirmado, via `Sheet8.cls:15`).
7. **Procedimentos chamados**: `Criar_Formula_Filtros` (Auxiliar.bas).
8. **Parâmetros**: nenhum.
9. **Retorno**: N/A (Sub, não Function — diferente das outras duas versões `_TK`, que são `Function`).
10. **Variáveis/objetos relevantes**: `Sh_Aux=Sheet9`; `Lin_Inicial=6` (hardcoded, linha 887); `Ult_Linh = CountA(sh.Range("A1").EntireColumn) + 3` (linha 891 — **terceira fórmula diferente** de cálculo de fim de range, distinta tanto da versão de Auxiliar.bas quanto das outras duas versões `_TK` de IFRS/Proforma); execução condicionada a `qtdeLinhasBase > 0` (linha 889 — única das 6 rotinas de reclassificação, entre as 2 famílias, que verifica explicitamente se há dados antes de processar).
11. **Abas/intervalos acessados**: `Sheet3` ("EMPRESA"), `Sheet9` (colunas "Empresa*").
12. **Pré-condições**: mesma premissa de layout fixo; adicionalmente, só executa se `qtdeLinhasBase > 0`.
13. **Passos principais**: mesma lógica de montagem de fórmula (sem divisão em 2 partes) de `Reclassificar_Combinacoes_Empresas` de Auxiliar.bas; aplica ao bloco inteiro de colunas "EMPRESA" de uma vez; bake-in via atribuição direta.
14. **Pós-condições**: bloco de colunas "EMPRESA" recalculado (se havia dados).
15. **Efeitos colaterais/riscos/evidência**: **tripla divergência** em relação à versão "canônica": (a) premissa de layout fixo vs. dinâmico (mesmo risco documentado nos itens 71-73); (b) **ausência da divisão em 2 partes** que a versão de Auxiliar.bas implementa (presumivelmente por razão de performance/limite de fórmula) — se essa divisão foi necessária para a Base em produção, esta versão `_TK` pode gerar uma fórmula única excessivamente longa/lenta para o mesmo volume de combinações, sem o benefício da otimização aplicada à versão irmã; (c) fórmula de cálculo de `Ult_Linh` própria e distinta das outras duas versões `_TK` do mesmo módulo. É o exemplo mais forte de divergência de implementação dentro da própria família `_TK`, evidenciando que nem as 3 versões `_TK` entre si estão alinhadas, além de divergirem da família `Auxiliar.bas`. Evidência: linhas 877-953, especialmente ausência de divisão em partes (comparar com Auxiliar.bas 757-763) e fórmula de `Ult_Linh` linha 891.

---


## 10.3 Cluster Refresh / Validação / IFRS16 (23 procedimentos — 9 módulos)
# A) CATÁLOGO POR PROCEDIMENTO

## Módulo `Refresh_DP_Segmento.bas`

### 1. Refresh_Base_Segmento

1. **Nome completo:** `Refresh_Base_Segmento`
2. **Módulo:** `Refresh_DP_Segmento.bas`
3. **Tipo:** Function (não usa o valor de retorno da função — comportamento de "Sub disfarçada de Function", padrão recorrente neste projeto)
4. **Escopo:** `Public` (declarado `Public Function`) — `Refresh_DP_Segmento.bas:4`
5. **Objetivo (negócio):** Reconstruir do zero a tabela `DP_Segmento` (Sheet21) a partir do SQL Server — mapeamento de Classe Custo + Centro Custo para Segmento, Abertura_1, Empresa e Diretorias Gerenciais N1-N3, usada como fonte de lookup para os campos `SEG_N2_DESTINO`/`A1_DESTINO` de Ajustes.
6. **Quem chama (evidência):** `Auxiliar.bas:571` (`Call Refresh_Base_Segmento`, dentro de `Sub Refresh_Base_Aux`). Nenhuma outra ocorrência de chamada encontrada no `grep` completo do diretório.
7. **Procedimentos chamados:** `AbreConexao` (`Refresh_DP_Segmento.bas:18`, definido em `Conexoes.bas:9`), `FechaConexao` (`Refresh_DP_Segmento.bas:83`, definido em `Conexoes.bas:47`). Não chama nenhum dos outros 8 módulos do escopo.
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum valor atribuído ao nome da função — `Function` usada apenas como agrupador de Sub.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet21`), `fn` (=`Application.WorksheetFunction`), `conn`/`Comando`/`Info` (ADODB), `TB_BANCO_SQL` (nome da tabela SQL, lido de `Sheet24!D20`), `Sheet_Oculta` (guarda visibilidade original da aba).
11. **Abas/intervalos/arquivos acessados:** Leitura: SQL Server `BPAM.dbo.[<tabela em Sheet24!D20>]`. Escrita: `Sheet21` (DP_Segmento), colunas 1 em diante, a partir da linha 1 (cabeçalho) e 2 (dados) — `Refresh_DP_Segmento.bas:42,56-64`.
12. **Pré-condições:** conexão SQL Server disponível; célula `Sheet24!D20` contendo nome de tabela válido; usuário com permissão de leitura na tabela do banco `BPAM`.
13. **Passos principais:** (1) `sh.ShowAllData` para remover filtros — `:12`; (2) torna `Sheet21` visível — `:38`; (3) abre conexão e executa `SELECT CONCAT(...) AS 'CHAVE', ...` ordenado por 6,7,8,9,10,4,5,2,3 — `:44-48`; (4) grava cabeçalhos e dados via `CopyFromRecordset` — `:58-64`; (5) `sh.Calculate` e restaura visibilidade original — `:69-70`; (6) fecha conexão — `:83`.
14. **Pós-condições:** `Sheet21` (DP_Segmento) contém o snapshot mais recente do mapeamento Classe Custo+Centro Custo→Segmento/Abertura_1/Empresa/Diretorias, a partir da tabela SQL configurada em `Sheet24!D20`. Visibilidade original da aba é restaurada.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Bloco `tratar_erro` (`:79-80`, MsgBox de erro de rede com contato `dfigsilva@timbrasil.com.br`) **nunca é alcançado** porque `On Error GoTo tratar_erro` está comentado (linha `' On Error GoTo tratar_erro`, `:33`) — dead code de tratamento de erro, confirmado por leitura direta. `CommandTimeout = 1000` sem retry (`:50`). Nenhuma verificação de que o `Recordset` retornou linhas antes de assumir sucesso. É o mais simples dos 5 módulos de refresh (bom candidato a primeiro a portar/testar).

---

## Módulo `Refresh_De_X_Para.bas`

### 2. Refresh_Base_De_Para_Ref_Cruzadas

1. **Nome completo:** `Refresh_Base_De_Para_Ref_Cruzadas`
2. **Módulo:** `Refresh_De_X_Para.bas`
3. **Tipo:** Function (padrão "Sub disfarçada", valor de retorno não usado)
4. **Escopo:** `Public` — `Refresh_De_X_Para.bas:3`
5. **Objetivo (negócio):** Importar de um arquivo Excel externo (`Bases_DE_PARA.xlsx`) as tabelas de responsabilidade cruzada (Ref Cruzada) e de centro de custo, que definem regras de redirecionamento excepcional de resultado entre diretorias/centros de custo quando a alocação padrão via `Sup_Linhas` não é suficiente.
6. **Quem chama (evidência):** `Auxiliar.bas:572` (`Call Refresh_Base_De_Para_Ref_Cruzadas`, dentro de `Refresh_Base_Aux`). Nenhuma outra chamada encontrada.
7. **Procedimentos chamados:** `Desligar_Tudo` (`:8`, `Auxiliar.bas:128`), `Extrair_Linhas_Ref_Cruzada` (5 vezes: `:28,40,68,80,114`, procedimento nº 3 abaixo, no mesmo módulo), `Ativar_Tudo` (`:210`, `Auxiliar.bas:69`).
8. **Parâmetros:** nenhum (usa variáveis globais/módulo `Arq`, `Diret`, `Col_1`, `Col_2` sem `Dim` explícito para a maioria, exceto `Col_1, Col_2 As Variant` — `:4`).
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `Arq`/`Diret` (nome/caminho do arquivo externo, lidos de `Sheet24!D15`/`D14`), `Sh_Destino`/`Sh_Origem` (reatribuídos a cada bloco), `Ult_Lin`, `Col_Chave`.
11. **Abas/intervalos/arquivos acessados:** Entrada externa: arquivo `Sheet24!D14`+`D15`, abas `Resp. Cruzada` (`:25`), `Resp_Cruzada_CdC` (`:37`), `CC_IFRS_Contabil` (`:65`), `CC BD` (`:77`), `Base CDC` (`:111`). Saída interna: `Ref_Cruzada_1` (`:22`), `Ref_Cruzada_2` (`:34`), `CC BD` (`:62,74`), `Sup_Linhas`/`Sheet15` (`:46-56`, lookup `OFFSET`/`MATCH` em M2:T + `Sheets("Sup_Linhas")` linha `:108`, Q1 em diante).
12. **Pré-condições:** arquivo `Diret & Arq` acessível na rede no momento da execução; abas de origem com os nomes exatos esperados; `Sup_Linhas` (Sheet15) já populada o suficiente para o lookup `OFFSET/MATCH` em `:51-56` funcionar (dependência de ordem — ver seção C).
13. **Passos principais:** (1) `Desligar_Tudo`; (2) abre `Diret & Arq` somente leitura (`:17`); (3) para cada par origem/destino, chama `Extrair_Linhas_Ref_Cruzada` copiando colunas por nome de cabeçalho; (4) recalcula `IFRS_Contabil` em `CC BD` via `INDEX/MATCH` (`:92-97`); (5) monta chave concatenada (`:102`); (6) copia `Base CDC`→`Sup_Linhas` col Q (`:111-115`); (7) aplica renomeação de Empresa hardcoded e regra Fiber WTTx/Fiber Live (`:128-150`); (8) ordena por 6 colunas (`:166-184`); (9) reconstrói chaves (`:186-190`); (10) fecha o arquivo externo sem salvar (`:205`); (11) `Ativar_Tudo`.
14. **Pós-condições:** `Ref_Cruzada_1`, `Ref_Cruzada_2`, `CC BD` e `Sup_Linhas` (colunas Q em diante) atualizados com os dados mais recentes do arquivo `Bases_DE_PARA.xlsx`.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Nenhum `On Error` na função principal — se o arquivo externo não abrir, a macro trava sem mensagem amigável. Regras de negócio de renomeação de Empresa **hardcoded** em VBA: `TPAR→TPAR_HO`, `Fiber_OG→Fiber Ongoing`, `INT_OG→Intelig Ongoing`, `INT_RR→Intelig Rural`, `Fiber_RR→Fiber Rural`, `CRC→Bloqueado (CRC)` (`:143-148`). Regra `Fiber WTTx` vs `Fiber Live`: `If UCase(Classe)="NO UBB" Then EMPRESA="Fiber WTTx" Else "Fiber Live"` aplicada apenas quando Empresa="FIBER" (`:129-137`), em loop célula-a-célula (baixo desempenho para base grande). `Windows(Arq).Close SaveChanges:=False` sem checar se o `Workbooks.Open` teve sucesso.

### 3. Extrair_Linhas_Ref_Cruzada

1. **Nome completo:** `Extrair_Linhas_Ref_Cruzada`
2. **Módulo:** `Refresh_De_X_Para.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito, sem modificador) — `Refresh_De_X_Para.bas:216`
5. **Objetivo (negócio):** Rotina genérica de cópia: para cada coluna do cabeçalho da planilha destino, localiza a coluna de mesmo nome na planilha origem e copia os valores, tratando erros `#N/A`/vazios — permite reusar o mesmo código para os 5 pares origem/destino do procedimento nº 2.
6. **Quem chama (evidência):** `Refresh_De_X_Para.bas:28,40,68,80,114` (5 chamadas, todas dentro de `Refresh_Base_De_Para_Ref_Cruzadas`, mesmo módulo). Nenhuma chamada de outro módulo.
7. **Procedimentos chamados:** `HandleNAErrorIsErrorWithCleaning` (`:244`, procedimento nº 4 abaixo, para cada célula copiada).
8. **Parâmetros:** `Sh_Destino`, `Sh_Origem`, `Arq`, `Col_Inicial` (saída — recebe `Sh_Origem.Range(Range_Destino).Column`, `:225`), `Range_Destino` (entrada, ex. `"C1"`).
9. **Retorno:** N/A (Sub); `Col_Inicial` funciona como parâmetro de saída por passagem `ByRef` (padrão default do VBA).
10. **Variáveis/objetos relevantes:** `fn` (WorksheetFunction), `Col_Chave`, `Col_Origem`, `Ult_Lin_Atual`, `Ult_Lin_Max` (acumula o maior número de linhas entre as colunas processadas — variável de módulo, não `Dim`, risco de arrastar valor entre chamadas).
11. **Abas/intervalos/arquivos acessados:** Lê `Sh_Origem` (linha 1 = cabeçalho, coluna a partir de `Range_Destino`); escreve em `Sh_Destino`, mesma posição de coluna cujo cabeçalho bate com o nome em `Sh_Origem`.
12. **Pré-condições:** `Sh_Destino` já deve ter, na linha 1, os nomes de coluna que se deseja importar (a função só copia colunas cujo cabeçalho já existe no destino — `If Col_Chave = "" Then Exit Do`, `:231`).
13. **Passos principais:** (1) `Sh_Origem.ShowAllData` (`:222`); (2) localiza `Col_Inicial` pela célula `Range_Destino` (`:225`); (3) loop `Do` por colunas do destino até achar célula de cabeçalho vazia; (4) para cada coluna com nome existente na origem (`fn.CountIfs`, `:233`), copia valores e aplica `HandleNAErrorIsErrorWithCleaning` linha a linha (`:243-246`).
14. **Pós-condições:** Colunas de `Sh_Destino` cujo cabeçalho bate com `Sh_Origem` ficam com os valores mais recentes da origem, sem `#N/A` (substituído por "Not Found"/"Error").
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** `On Error Resume Next` seguido de `On Error GoTo 0` ao redor de `ShowAllData` (`:221-223`) — suprime qualquer erro dessa chamada silenciosamente. Loop célula-a-célula (`For Lin = 2 To Ult_Lin_Max`, `:243`) é O(n) por coluna e pode ser lento em bases grandes. Variável `Ult_Lin_Max` não é resetada entre colunas nem entre chamadas (não há `Dim`/reinício explícito visível no trecho lido) — risco de comportamento cumulativo entre colunas dentro da mesma chamada.

### 4. HandleNAErrorIsErrorWithCleaning

1. **Nome completo:** `HandleNAErrorIsErrorWithCleaning`
2. **Módulo:** `Refresh_De_X_Para.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Refresh_De_X_Para.bas:255`
5. **Objetivo (negócio):** Padronizar tratamento de erro de célula após lookup: transforma `#N/A` em texto legível "Not Found", outros erros em "Error", e limpa células vazias — evita que erros de fórmula "vazem" para as tabelas de suporte copiadas.
6. **Quem chama (evidência):** `Refresh_De_X_Para.bas:244` (dentro de `Extrair_Linhas_Ref_Cruzada`, mesmo módulo). Nenhuma outra chamada encontrada no `grep` do diretório.
7. **Procedimentos chamados:** nenhum (usa apenas funções nativas VBA `IsError`, `CVErr`).
8. **Parâmetros:** `targetCell` (referência de célula, passada por `ByRef` default).
9. **Retorno:** N/A (Sub); efeito é escrito diretamente em `targetCell.Value`.
10. **Variáveis/objetos relevantes:** nenhuma variável de módulo própria além do parâmetro.
11. **Abas/intervalos/arquivos acessados:** a célula individual passada como parâmetro (qualquer aba, definida pelo chamador).
12. **Pré-condições:** `targetCell` deve ser uma referência de célula válida já avaliada (contém valor ou erro).
13. **Passos principais:** (1) `If IsError(targetCell.Value)` — se erro `#N/A` (`CVErr(xlErrNA)`) grava `"Not Found"`, senão grava `"Error"` (`:257-263`); (2) `ElseIf targetCell.Value = ""` então `ClearContents` (`:264-265`).
14. **Pós-condições:** célula nunca contém um erro nativo do Excel após a chamada — ou tem texto substituto, ou está vazia, ou mantém o valor original.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Substituir todo erro não-`#N/A` genericamente por `"Error"` (`:262`) **oculta a causa raiz** do erro (`#VALUE!`, `#REF!`, `#DIV/0!` etc. todos viram o mesmo texto "Error") — dificulta diagnóstico. Sem log de qual célula/linha gerou o erro original.

---

## Módulo `Refresh_Sup_Linhas.bas`

### 5. Refresh_Base_Suporte_Linhas

1. **Nome completo:** `Refresh_Base_Suporte_Linhas`
2. **Módulo:** `Refresh_Sup_Linhas.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Refresh_Sup_Linhas.bas:4`
5. **Objetivo (negócio):** Reconstruir a tabela mestre `Sup_Linhas` (Sheet15) — o dicionário central que traduz combinações de Aberturas 2-8, Classe Custo, Diretoria Gerencial N1-N3 e Segmento a partir do SQL Server, incluindo cálculo de "IFRS_Contabil" e listas de KPI/Versão.
6. **Quem chama (evidência):** `Auxiliar.bas:573` (`Call Refresh_Base_Suporte_Linhas`, dentro de `Refresh_Base_Aux`). Nenhuma outra chamada encontrada.
7. **Procedimentos chamados:** `AbreConexao`/`FechaConexao` (`Conexoes.bas`); `Atualizar_Lista_KPI_Versao` (`:267`, procedimento nº 7 abaixo, mesmo módulo).
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet15`), `Sh_Aux` (=`Sheet2`), `Sh11` (=`Sheet11`), `TB_BANCO_SQL` (`Sheet24!D19`), `TB_BANCO_SQL_AUX_SUP_LIN` (`Sheet24!D21`), `Col_Destino` (cursor de coluna incrementado manualmente por bloco), `Qtd_Colunas`.
11. **Abas/intervalos/arquivos acessados:** Leitura SQL: `BPAM.dbo.TB_HUBBLE_DBS_CC` (query explícita, `:65`), tabela em `Sheet24!D19` (`:126-141`, ainda que o nome literal `TB_HUBBLE_CLUSTER_RPD_NEW` apareça comentado como referência histórica em `:30`), `Sheet24!D19` novamente em `:213`, tabela em `Sheet24!D21` (`:240`). Escrita: `Sheet15` em 4 blocos de colunas consecutivos (a partir da coluna 4, depois `+Qtd_Colunas+14` repetido duas vezes, depois `+Qtd_Colunas+5`); `Sheet11` (coluna "EMPRESA", `:272-286`).
12. **Pré-condições:** conexão SQL disponível; `Sheet24!D19`/`D21` com nomes de tabela válidos; `Sheet2` já contendo colunas `"Contas IFRS"`/`"Efeito"` para o bloco comentado de IFRS_Contabil (bloco está desativado — ver item 15).
13. **Passos principais:** (1) `ShowAllData`; (2) bloco 1: `SELECT DISTINCT ABERTURA_2..8, LINHA_BD, CLASSE CUSTO, IFRS_CONTABIL FROM TB_HUBBLE_DBS_CC` (`:61-67`), grava e cria coluna `AA_Chave` (fórmula concatenando IFRS_Contabil+Aberturas, `:86-89`); (3) bloco 2: `SELECT DISTINCT EMPRESA, ABERTURA_2_SEG, SEGMENTO, ABERTURA_1, CLASSE CUSTO, '-' AS CENTRO CUSTO` com `LEFT JOIN TB_HUBBLE_DBS_CC` (`:131-141`), cria `AA_Chave_Segmento` (`:160-163`); (4) ordena tudo por 10 colunas (`:182-206`); (5) bloco 3: lista distinta de `ORGANIC` (`:212-215`); (6) bloco 4: `DIRETORIA N1/N2/N3 GERENCIAL` filtrado por `TIPOLOGIA` do primeiro registro da tabela principal (`:238-242`), com fórmulas de numeração sequencial condicional (`:262-263`); (7) `Atualizar_Lista_KPI_Versao`; (8) lista de `EMPRESA` para `Sheet11` (`:274-286`); (9) `sh.Calculate` e restaura visibilidade.
14. **Pós-condições:** `Sheet15` (Sup_Linhas) com 4 blocos de colunas atualizados e ordenados; `Sheet11` com lista de empresas; `Sheet15` também recebe (via chamada interna) a lista de KPI/Versão.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Mesmo padrão de `tratar_erro` morto (`On Error GoTo tratar_erro` comentado, `:35`; bloco `:303-304` nunca alcançado). SQL montado por concatenação de string sem parametrização (`:61-67,131-141,212-215,238-242`) — risco moderado de injeção se `TB_BANCO_SQL`/`TB_BANCO_SQL_AUX_SUP_LIN` alguma vez vierem de entrada não controlada (hoje vêm de célula de configuração). Posicionamento de colunas por deslocamento aritmético (`Col_Destino = Col_Destino + Qtd_Colunas + 14`, `:114,210`) — número mágico `14`/`5` sem explicação, extremamente frágil a mudanças na query. Bloco de cálculo `IFRS_Contabil` via `INDEX/MATCH` em `Sh_Aux`(Sheet2) está **inteiramente comentado** (`:102-110`) — a coluna `IFRS_Contabil` hoje vem diretamente do SQL (bloco 1, `:64`), não mais calculada localmente; isto é evidência de uma migração de lógica (de fórmula Excel para SQL) parcialmente limpa no código (comentário morto remanescente). Lista de versões "Preview 1-6, Pré-Closing, AJ_Pré-Closing" hardcoded dentro de `Atualizar_Lista_KPI_Versao` (ver procedimento 7). Ordenação usa `xlPinYin`.

### 6. Extrair_Valid_Lin

1. **Nome completo:** `Extrair_Valid_Lin`
2. **Módulo:** `Refresh_Sup_Linhas.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Refresh_Sup_Linhas.bas:312`
5. **Objetivo (negócio):** Popular `Sheet25` com as combinações válidas e distintas de Aberturas 2-8, Classe Custo e Diretorias N1-N3, excluindo `ABERTURA_2='EBITDA'` — é a base de referência usada para restringir os dropdowns dependentes de "Main Results" (ver `BackupCodigo_MainResults.bas`).
6. **Quem chama (evidência):** `Auxiliar.bas:575` (`Call Extrair_Valid_Lin`, última etapa de `Refresh_Base_Aux`). Nenhuma outra chamada encontrada.
7. **Procedimentos chamados:** `AbreConexao`/`FechaConexao` (`Conexoes.bas`).
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet25`), `TB_BANCO_SQL_AUX_SUP_LIN` (`Sheet24!D21`).
11. **Abas/intervalos/arquivos acessados:** Leitura SQL: tabela em `Sheet24!D21`. Escrita: `Sheet25`, a partir da coluna 2 (`:341-365`).
12. **Pré-condições:** conexão SQL disponível; `Sheet24!D21` válida.
13. **Passos principais:** (1) `ShowAllData`; (2) `SELECT DISTINCT ABERTURA_2..8, CLASSE CUSTO, DIRETORIA N1/N2/N3 GERENCIAL, CLASSE FROM <tabela> WHERE ABERTURA_2 <> 'EBITDA' ORDER BY 1,2,3,4,5,6,7,9,10,11,8` (`:343-349`); (3) grava cabeçalhos e dados via `CopyFromRecordset` (`:359-365`).
14. **Pós-condições:** `Sheet25` contém a base de combinações válidas mais recente, consumida por `BackupCodigo_MainResults.bas` (quando/se ativo — ver risco correspondente).
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Sem bloco `tratar_erro` nesta função (diferente das demais de refresh) — se a query falhar, o erro VBA não tratado interrompe a execução sem mensagem amigável. `CommandTimeout = 1000`. Exclusão hardcoded de `EBITDA` (`:348`) — regra de negócio embutida no filtro SQL, sem explicação de por que essa abertura é excluída da validação.

### 7. Atualizar_Lista_KPI_Versao

1. **Nome completo:** `Atualizar_Lista_KPI_Versao`
2. **Módulo:** `Refresh_Sup_Linhas.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` (implícito, sem modificador) — `Refresh_Sup_Linhas.bas:376`
5. **Objetivo (negócio):** Obter do SQL Server todas as combinações reais de `KPI + Versão` existentes na base RPD, complementadas por uma lista fixa de versões de preview/fechamento, para alimentar o filtro de KPI/Versão usado em toda a ferramenta.
6. **Quem chama (evidência):** `Refresh_Sup_Linhas.bas:267` (dentro de `Refresh_Base_Suporte_Linhas`, mesmo módulo); `Extracao_Sheet_Ajustes.bas:14` (`Call Atualizar_Lista_KPI_Versao`, fora do escopo dos 9 módulos, mas confirma reuso em outro fluxo de extração).
7. **Procedimentos chamados:** `AbreConexao`/`FechaConexao`; `Atualizar_Lista_KPI_Versao_Interna` (`:434`, procedimento nº 8, mesmo módulo).
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet15`), `TB_BANCO_SQL` (`Sheet24!D19`), `Col_Destino` (posição da coluna "KPI_VERSAO" já existente, localizada por `fn.Match`).
11. **Abas/intervalos/arquivos acessados:** Leitura SQL: tabela em `Sheet24!D19`. Escrita: `Sheet15`, colunas a partir de "KPI_VERSAO" (localizada dinamicamente, `:396`).
12. **Pré-condições:** `Sheet15` já deve ter uma coluna de cabeçalho literalmente chamada `"KPI_VERSAO"` (senão `fn.Match` gera erro não tratado).
13. **Passos principais:** (1) monta SQL com `UNION` entre a query real (`KPI + ' > ' + VERSÃO`, ordenada com "DIVULGADO" primeiro) e 7 linhas fixas hardcoded — `'ACT > Preview 1'` a `'ACT > Preview 6'`, `'ACT > Pré-Closing'`, `'ACT > AJ_Pré-Closing'` (`:398-413`); (2) grava resultado em `Sheet15`; (3) chama `Atualizar_Lista_KPI_Versao_Interna`.
14. **Pós-condições:** `Sheet15` (coluna KPI_VERSAO e adjacentes) contém todas as combinações reais do banco mais as versões de preview/fechamento fixas, prontas para serem complementadas pela etapa "Interna".
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** **Regra de negócio hardcoded**: o calendário de fechamento (Preview 1 a 6, Pré-Closing, AJ_Pré-Closing) está embutido diretamente na string SQL (`:404-411`) — se o calendário de fechamento mudar (nº de previews, novo marco), requer alteração de código VBA, não de configuração. Sem tratamento de erro próprio (depende do `tratar_erro` do módulo, que está morto — ver procedimento 5, item 15).

### 8. Atualizar_Lista_KPI_Versao_Interna

1. **Nome completo:** `Atualizar_Lista_KPI_Versao_Interna`
2. **Módulo:** `Refresh_Sup_Linhas.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Refresh_Sup_Linhas.bas:439`
5. **Objetivo (negócio):** Complementar a lista de KPI/Versão com combinações que já existem na base histórica (`Sheet3`) mas ainda não estão cadastradas em `Sheet15`, e ordenar a lista final com prioridade: Divulgado (0) < demais (1) < Pré-Closing (2).
6. **Quem chama (evidência):** `Refresh_Sup_Linhas.bas:434` (dentro de `Atualizar_Lista_KPI_Versao`, mesmo módulo). Chamada adicionalmente, fora do escopo dos 9 módulos, por: `Extracao_Base_MOCKUP_RGM.bas:13`, `Extracao_SQL_Hubble.bas:13`, `Extracao_Base_Consolidad.bas:14`, `Auxiliar.bas:22`, `Extracao_Base_Quick_Data.bas:14`, `Extracao_Base_Other_Inco.bas:14`, `Extracao_Base_1009.bas:13`, `Extracao_Base_RGM.bas:13`, `Extracao_Fixed_Revenues.bas:13` — evidência de que esta rotina é reusada como etapa padrão de pós-processamento em praticamente todas as extrações do sistema, não só no pipeline de refresh.
7. **Procedimentos chamados:** nenhum procedimento customizado (apenas funções nativas `fn.Match`, `fn.CountA`, `fn.CountIfs`, `Sort`).
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet15`), `Sh_Base` (=`Sheet3`), `Sep=" > "`, `Col_KPI_VERSAO`, `Ordernar` (flag, com erro de digitação de "Ordenar", nunca lida após ser setada — variável morta, ver item 15).
11. **Abas/intervalos/arquivos acessados:** Leitura: `Sheet3` (colunas "KPI"/"VERSÃO", a partir da linha marcada "LIN_BASE" na coluna A). Escrita: `Sheet15`, colunas KPI_VERSAO/KPI/VERSÃO e coluna auxiliar de ordenação à esquerda.
12. **Pré-condições:** `Sheet3` (Main Results) deve conter uma célula com o texto literal `"LIN_BASE"` na coluna A, e colunas de cabeçalho `"KPI"`/`"VERSÃO"`.
13. **Passos principais:** (1) para cada linha de `Sheet3`, monta `KPI_VERSAO = KPI & " > " & Versão`; se ainda não existir em `Sheet15`, adiciona no fim (`:457-468`); (2) calcula coluna auxiliar de ordem: `0` se "DIVULGADO", `2` se "PRÉ-CLOSING", senão `1` (`:479`); (3) torna a aba visível, ordena por 3 colunas (`:484-499`); (4) recalcula a coluna de ordem com fórmula de numeração sequencial condicional (`:501`); (5) restaura visibilidade; (6) `sh.Range(Range_Order).EntireColumn.Calculate` (`:532`).
14. **Pós-condições:** `Sheet15` com lista de KPI/Versão completa (SQL + valores reais da base histórica não cadastrados) e ordenada pela prioridade de negócio (Divulgado < demais < Pré-Closing).
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Variável `Ordernar` é setada em `:466` mas **nunca lida** — código morto (indício de lógica condicional removida sem limpar a variável). Bloco grande de código comentado (`:506-529`) documentando uma versão anterior da lógica de ordenação — dívida de documentação/histórico não limpo. Sem tratamento de erro próprio.

---

## Módulo `Refresh_Drop_Comb.bas`

### 9. Refresh_Drop_Comb_Hubble

1. **Nome completo:** `Refresh_Drop_Comb_Hubble`
2. **Módulo:** `Refresh_Drop_Comb.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Refresh_Drop_Comb.bas:3`
5. **Objetivo (negócio):** Copiar a aba "DropComb" de um arquivo externo de configuração para dentro do Quick Data — fonte das listas de itens dos combos/dropdowns da tela FRONT.
6. **Quem chama (evidência):** `Auxiliar.bas:574` (`Call Refresh_Drop_Comb_Hubble`, dentro de `Refresh_Base_Aux`). Nenhuma outra chamada encontrada.
7. **Procedimentos chamados:** `Desligar_Tudo` (`:7`), `Ativar_Tudo` (`:37`) — ambos em `Auxiliar.bas`.
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `Arq` (`Sheet24!D9`), `Diret` (`Sheet24!D7`), `Sh_Origem`/`Sh_Destino` (=`Sheet9`).
11. **Abas/intervalos/arquivos acessados:** Entrada externa: arquivo em `Sheet24!D7`+`D9`, aba `"DropComb"`. Saída: `Sheet9` (DropComb interna), todas as células.
12. **Pré-condições:** arquivo externo acessível na rede; aba `"DropComb"` existente no arquivo com esse nome exato.
13. **Passos principais:** (1) `Desligar_Tudo`; (2) abre arquivo externo somente leitura, `Editable:=False` (`:13`); (3) `Sh_Origem.Cells.Copy` → `Sh_Destino.Cells.PasteSpecial xlPasteAll` (`:25-27`, cola tudo — valores, fórmulas, formatos, validações); (4) fecha arquivo externo sem salvar (`:32`); (5) `Ativar_Tudo`.
14. **Pós-condições:** `Sheet9` (DropComb) espelha integralmente a aba "DropComb" do arquivo externo mais recente.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Nenhum `On Error` em toda a função — se o arquivo ou a aba não existirem, a macro trava com erro VBA não tratado, sem mensagem amigável (diferente dos módulos de refresh SQL, que ao menos têm um bloco de erro, ainda que morto). `PasteSpecial xlPasteAll` (não `xlPasteValues`) replica formatação/comentários/validações da origem, podendo ser lento e "sujar" a formatação da aba interna com qualquer resíduo da planilha de origem.

---

## Módulo `Limpeza_Base_Ajustes.bas`

### 10. Limpar_Ajustes

1. **Nome completo:** `Limpar_Ajustes`
2. **Módulo:** `Limpeza_Base_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Limpeza_Base_Ajustes.bas:3`
5. **Objetivo (negócio):** Wrapper de UI para o usuário disparar a limpeza/resincronização completa da planilha de Ajustes manuais, com confirmação visual de conclusão.
6. **Quem chama (evidência):** Nenhuma chamada encontrada no `grep` completo do diretório (nem em `.bas`, nem em `.frm`) — **evidência de botão dedicado fora do dump de texto** (shape/form control com `OnAction="Limpar_Ajustes"`, não recuperável deste dump).
7. **Procedimentos chamados:** `Desligar_Tudo` (`:5`), `Processo_Limpar_Ajustes` (`:6`, procedimento nº 11, mesmo módulo), `Ativar_Tudo` (`:7`) — todos em `Auxiliar.bas` exceto o segundo.
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** nenhuma variável própria de destaque.
11. **Abas/intervalos/arquivos acessados:** indiretamente, `Sheet13` (via `Processo_Limpar_Ajustes`) e `Sheet13.Calculate` (`:8`).
12. **Pré-condições:** nenhuma verificada explicitamente no código.
13. **Passos principais:** (1) `Desligar_Tudo`; (2) `Processo_Limpar_Ajustes`; (3) `Ativar_Tudo`; (4) `Sheet13.Calculate`; (5) `MsgBox "Processo concluído com sucesso!"` (`:10`).
14. **Pós-condições:** `Sheet13` (AJUSTES) limpa e com fórmulas de destino recalculadas; usuário informado via MsgBox.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Sem `On Error` — se `Processo_Limpar_Ajustes` falhar no meio, `Ativar_Tudo` (que religa `ScreenUpdating`/`DisplayAlerts`/`EnableEvents`) pode não ser executado, deixando a planilha em estado "travado" (eventos desligados) até o usuário reabrir o arquivo ou rodar outra macro que religue.

### 11. Processo_Limpar_Ajustes

1. **Nome completo:** `Processo_Limpar_Ajustes`
2. **Módulo:** `Limpeza_Base_Ajustes.bas`
3. **Tipo:** Function (retorno não usado)
4. **Escopo:** `Public` — `Limpeza_Base_Ajustes.bas:15`
5. **Objetivo (negócio):** Núcleo da limpeza/resincronização de Ajustes: apaga ~25 campos de entrada digitados/calculados (Versão, Exercício, Aberturas, meses, etc.) e reescreve ~12 fórmulas de campos de destino (Classe/Centro Custo, FY, Ref Cruzada, Empresa/Diretoria/Segmento Destino) alinhadas às tabelas de suporte mais recentes.
6. **Quem chama (evidência):** `Limpeza_Base_Ajustes.bas:6` (dentro de `Limpar_Ajustes`, mesmo módulo); `Form_Importacao.frm:494` (`Call Processo_Limpar_Ajustes`, dentro do fluxo de botão `B_Ok_Click` do formulário de Importação — confirma execução automática ao importar a aba "AJUSTES" de um arquivo/período anterior, antes de recarregar os dados brutos, conforme contexto em `Form_Importacao.frm:477-494`).
7. **Procedimentos chamados:** `Comando_Limpar` (procedimento nº 12, chamado ~25 vezes), `Comando_Refazer_Formula` (procedimento nº 13, chamado ~12 vezes) — ambos no mesmo módulo.
8. **Parâmetros:** nenhum.
9. **Retorno:** nenhum.
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet13`), `Sh_Sup_Lin` (=`Sheet15`), `Lin_Cabecalho`/`Col_Inicial` (localizados pela célula que contém "VERSÃO"), `Ult_Lin` (heurística: `Match` de "" na coluna Ref Cruzada +50, decrementada até achar dado real, `:29-34`).
11. **Abas/intervalos/arquivos acessados:** Leitura/escrita: `Sheet13` (AJUSTES). Lookup (fórmulas escritas referenciam): `Sup_Linhas` (Sheet15), `Ref_Cruzada_1`, `Ref_Cruzada_2`, `CC BD`, `Sheet21` (DP_Segmento, referenciado por nome dinâmico `Sheet21.Name`).
12. **Pré-condições:** `Sheet13` deve conter célula com texto "VERSÃO" (cabeçalho); tabelas de suporte (`Sup_Linhas`, `Ref_Cruzada_1/2`, `CC BD`, `DP_Segmento`) devem estar atualizadas — dependência de ordem com o pipeline `Refresh_Base_Aux` (ver seção C).
13. **Passos principais:** (1) localiza cabeçalho e `Ult_Lin`; (2) `ShowAllData`; (3) limpa ~25 campos via `Comando_Limpar` (VERSÃO, EXERCICIO, IFRS_CONTABIL, ORGANIC, ABERTURA_1..8, SEGMENTO, LINHA_BD, EMPRESA, CLASSE, Diretoria N1-N3, Jan-Dez, CC_Destino, CDC_Destino); (4) reescreve fórmulas via `Comando_Refazer_Formula` para Chave_BD, CLASSE CUSTO, CENTRO CUSTO, FY, Ref Cruzada, EMPRESA DESTINO, CLASSE DESTINO, SEG_N2_DESTINO, A1_DESTINO, DIRETORIA N1-N3 DESTINO, A2-A8_DESTINO, Gr. BD 2-8 - Destino (regras de negócio detalhadas na seção B).
14. **Pós-condições:** `Sheet13` com entradas do usuário zeradas e todas as fórmulas de destino recalculadas com base nas tabelas de suporte vigentes no momento da execução.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Todas as fórmulas de negócio (cascata Ref Cruzada, SEG_N2_DESTINO, "Labour Cost"→"W/o Fiber" etc.) estão **hardcoded como strings de fórmula Excel dentro do VBA** (`:141-219`) — ponto de maior risco de perda de conhecimento na reescrita; precisam ser decodificadas 1:1 (ver seção B). `Ult_Lin` calculado por heurística frágil (`Do...Loop` decrementando, `:31-34`). Sem `On Error` na função principal além do `ShowAllData` (`:38-42`).

### 12. Comando_Limpar

1. **Nome completo:** `Comando_Limpar`
2. **Módulo:** `Limpeza_Base_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Limpeza_Base_Ajustes.bas:247`
5. **Objetivo (negócio):** Helper genérico: localizar uma coluna pelo nome do cabeçalho e limpar todo o intervalo de dados abaixo dela.
6. **Quem chama (evidência):** `Limpeza_Base_Ajustes.bas:47,52,57,60,66,72,77,82,87,92,98,117,123,128` (14 pontos de chamada, todos dentro de `Processo_Limpar_Ajustes`, mesmo módulo). Sendo `Private`, não pode ser chamado por outros módulos — confirmado pela ausência de qualquer ocorrência fora deste arquivo no `grep`.
7. **Procedimentos chamados:** nenhum (apenas `fn.Match`).
8. **Parâmetros:** `sh` (planilha), `Campo` (nome do cabeçalho a buscar), `Ult_Lin`, `Lin_Cabecalho`.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `Col` (posição encontrada via `fn.Match`).
11. **Abas/intervalos/arquivos acessados:** a planilha `sh` recebida por parâmetro (na prática, sempre `Sheet13`).
12. **Pré-condições:** `Campo` deve existir literalmente na linha `Lin_Cabecalho` de `sh` (senão `fn.Match` gera erro `#N/A` não tratado, que propagaria como erro VBA não capturado).
13. **Passos principais:** (1) `Col = fn.Match(Campo, sh.Cells(Lin_Cabecalho,1).EntireRow, 0)`; (2) `sh.Range(sh.Cells(Lin_Cabecalho+1,Col), sh.Cells(Ult_Lin,Col)).ClearContents`.
14. **Pós-condições:** coluna correspondente a `Campo` fica vazia da linha `Lin_Cabecalho+1` até `Ult_Lin`.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Nenhum tratamento de erro — se `Campo` não existir no cabeçalho, a macro-mãe (`Processo_Limpar_Ajustes`) para com erro VBA não tratado. Reuso limpo e correto (boa prática pontual dentro de um módulo com muitos outros riscos).

### 13. Comando_Refazer_Formula

1. **Nome completo:** `Comando_Refazer_Formula`
2. **Módulo:** `Limpeza_Base_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Limpeza_Base_Ajustes.bas:256`
5. **Objetivo (negócio):** Helper genérico: localizar uma coluna pelo nome do cabeçalho e reaplicar uma fórmula `FormulaR1C1` padrão em todo o intervalo de dados abaixo dela.
6. **Quem chama (evidência):** `Limpeza_Base_Ajustes.bas:136,144,152,158,169,179,185,199,210,223,231,239` (12 pontos, todos dentro de `Processo_Limpar_Ajustes`, mesmo módulo). `Private` — nenhuma chamada externa possível/encontrada.
7. **Procedimentos chamados:** nenhum (apenas `fn.Match`).
8. **Parâmetros:** `sh`, `Campo`, `Ult_Lin`, `Lin_Cabecalho`, `Formula` (string de fórmula `R1C1`).
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `Col` (posição encontrada via `fn.Match`).
11. **Abas/intervalos/arquivos acessados:** a planilha `sh` recebida por parâmetro (na prática, sempre `Sheet13`).
12. **Pré-condições:** `Campo` deve existir na linha `Lin_Cabecalho`; `Formula` deve ser sintaticamente válida como `FormulaR1C1` e referenciar corretamente colunas relativas ao contexto de `Campo`.
13. **Passos principais:** (1) `Col = fn.Match(Campo, ...)`; (2) `sh.Range(...).FormulaR1C1 = Formula`.
14. **Pós-condições:** coluna correspondente a `Campo` passa a conter a fórmula informada em todas as linhas de dados.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Nenhum tratamento de erro. Como as fórmulas passadas por `Processo_Limpar_Ajustes` são strings longas montadas por concatenação (referenciando números de coluna calculados dinamicamente, ex. `Col_CC`, `Col_CDC`), qualquer erro de cálculo de offset nessas strings produziria uma fórmula sintaticamente válida mas **semanticamente errada** sem nenhum aviso — risco silencioso.

---

## Módulo `Lista_Validacao_Ajustes.bas`

### 14. Atualizar_Ajustes_Lista_Validacao_Geral

1. **Nome completo:** `Atualizar_Ajustes_Lista_Validacao_Geral`
2. **Módulo:** `Lista_Validacao_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Lista_Validacao_Ajustes.bas:3`
5. **Objetivo (negócio):** Wrapper de UI: dispara a reconstrução completa das listas de validação (dropdowns) da planilha de Ajustes e confirma ao usuário.
6. **Quem chama (evidência):** Nenhuma chamada encontrada no `grep` completo do diretório — **evidência de botão dedicado fora do dump de texto**, provavelmente na planilha AJUSTES, próximo ao botão de `Limpar_Ajustes`.
7. **Procedimentos chamados:** `Processo_Atuliza_Lista_Validacao_Geral_Ajustes` (`:5`, procedimento nº 15, mesmo módulo).
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** nenhuma própria de destaque.
11. **Abas/intervalos/arquivos acessados:** indiretamente, `Sheet13` e `Sheet15` (via procedimento chamado).
12. **Pré-condições:** nenhuma verificada explicitamente.
13. **Passos principais:** (1) `Processo_Atuliza_Lista_Validacao_Geral_Ajustes`; (2) `MsgBox "Validação atualizada com sucesso!"` (`:6`).
14. **Pós-condições:** listas de validação de `Sheet13` reconstruídas; usuário informado.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Sem `On Error`; sem `Desligar_Tudo`/`Ativar_Tudo` ao redor da chamada (diferente de `Limpar_Ajustes`) — a rotina interna roda com `ScreenUpdating`/eventos ligados, potencialmente mais lenta e com tela "piscando" durante a execução.

### 15. Processo_Atuliza_Lista_Validacao_Geral_Ajustes

1. **Nome completo:** `Processo_Atuliza_Lista_Validacao_Geral_Ajustes` (nome contém erro de digitação original — "Atuliza" em vez de "Atualiza"; reproduzido exatamente como no código-fonte)
2. **Módulo:** `Lista_Validacao_Ajustes.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `Lista_Validacao_Ajustes.bas:10`
5. **Objetivo (negócio):** Reconstrói todas as listas de validação (dropdowns) de `Sheet13` (AJUSTES): 6 campos fixos ligados a intervalos nomeados (Empresa, Abertura_2, Classe, Diretoria N1-N3) e todos os campos dinâmicos a partir de "Abertura_3" até o início do bloco de meses, calculados a partir dos valores distintos existentes em `Sup_Linhas`.
6. **Quem chama (evidência):** `Lista_Validacao_Ajustes.bas:5` (dentro de `Atualizar_Ajustes_Lista_Validacao_Geral`, mesmo módulo). `Private` — nenhuma chamada externa possível/encontrada.
7. **Procedimentos chamados:** nenhum procedimento customizado (bubble sort e concatenação implementados inline, `:110-129` — duplicando lógica também presente em `BackupCodigo_MainResults.bas:188-207` e `Auxiliar.bas:1096`, sem reuso).
8. **Parâmetros:** nenhum.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `sh` (=`Sheet13`), `Sh_Base` (=`Sheet15`), `Ult_Lin` (mínimo hardcoded em 140, `:27`), `List_Campos_Validados` (registro dos campos já tratados, para não duplicar no loop dinâmico), `Qtd_Filtros` (limitado a 747, `:123`).
11. **Abas/intervalos/arquivos acessados:** Leitura: `Sheet15` (Sup_Linhas, valores distintos por coluna); intervalos nomeados fixos `CB_EMPRESA`, `CB_ABERTURA_2`, `CB_Classe`, `CB_N1_Gerencial`, `CB_N2_Gerencial_Linha`, `CB_N3_Gerencial_Linha`. Escrita: `Data Validation` das colunas de `Sheet13`, de "Abertura_3" até a coluna anterior a "Jan".
12. **Pré-condições:** `Sheet13` deve conter célula com texto "Versão" (localizada por `Cells.Find`); os intervalos nomeados fixos devem existir no workbook.
13. **Passos principais:** (1) localiza `Lin_Cabecalho`/`Col_1` pela célula "Versão"; (2) `Ult_Lin = max(140, última linha com dado na coluna Versão)`; (3) aplica validação fixa nas 6 colunas-chave (Empresa, Abertura_2, Classe, Diretoria N1-N3) via intervalos nomeados (`:32-83`); (4) loop `Do Until` de "Abertura_3" até achar "Jan": para cada coluna ainda não tratada, monta lista de valores distintos de `Sup_Linhas`, ordena (bubble sort manual) com "Total" sempre por último, corta em 747 itens, e aplica como `Data Validation` (`:89-141`).
14. **Pós-condições:** todas as colunas-chave e dinâmicas de `Sheet13` com listas de validação atualizadas conforme os valores vigentes em `Sup_Linhas`.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** **Limite hardcoded de 747 itens** (`:123`, reflete o limite técnico do Excel para listas de validação por valores separados por vírgula) implementado como **corte silencioso** — se `Sup_Linhas` tiver mais de 747 valores distintos numa coluna, os excedentes são descartados sem aviso ao usuário. `Ult_Lin` mínimo hardcoded em 140 (`:27`) sem explicação — número mágico. Bubble sort O(n²) manual duplicado em pelo menos 3 lugares do código-fonte (ver item 7). Concatenação de string para checar duplicidade (`InStr`, `:102`) é O(n) por item — algoritmo global O(n²) sobre valores distintos, pode degradar com crescimento de `Sup_Linhas`. Sem tratamento de erro.

---

## Módulo `Front_Processos.bas`

### 16. Atualizar_Front_Formula_Padrao

1. **Nome completo:** `Atualizar_Front_Formula_Padrao`
2. **Módulo:** `Front_Processos.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Front_Processos.bas:3`
5. **Objetivo (negócio):** Reaplicar a fórmula padrão de referência em células de uma planilha-modelo (tipo FRONT) que deveriam conter fórmula mas foram sobrescritas com dado fixo pelo usuário — mecanismo de "correção estrutural" de relatórios template.
6. **Quem chama (evidência):** Nenhuma chamada encontrada no `grep` completo do diretório — **evidência de botão genérico reutilizável** presente em várias planilhas de relatório, cada uma apontando para esta mesma macro (uso de `ActiveSheet` reforça esse padrão). [NÃO ACESSÍVEL: confirmação do(s) botão(ões)/shape(s) reais está fora do dump de texto].
7. **Procedimentos chamados:** `PopUp_Tempo_Processamento` (`:87`, `Auxiliar.bas:138`), `Ativar_Tudo` (`:88,94`, `Auxiliar.bas:69`). Também usa (sem `Call`) a variável global `Tit_Msg` (função, `Auxiliar.bas:202`) como parâmetro de MsgBox (`:15,98`).
8. **Parâmetros:** nenhum (opera sobre `ActiveSheet`).
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `sh` (=`ActiveSheet`), `F_Padrao` (fórmula de referência, lida da célula ao lado de "Bkp fórmula"), `Lin_1`/`Lin_2` (delimitam blocos contíguos de células candidatas), `Total_Verif`/`Qtd_Verif` (progresso, exibido na barra de status).
11. **Abas/intervalos/arquivos acessados:** `ActiveSheet` — célula com texto "Bkp fórmula" (guarda a fórmula padrão), região a partir de "Menu de seleção", coluna com valor iniciando em "Aux".
12. **Pré-condições:** `ActiveSheet` deve seguir a "estrutura padrão do arquivo default" (texto do próprio MsgBox de erro, `:98`) — deve conter as células/textos "Bkp fórmula", "Menu de seleção" e uma coluna iniciando com "Aux".
13. **Passos principais:** (1) confirmação `MsgBox vbYesNo` (`:13-18`); (2) localiza `F_Padrao` (`:31-35`); (3) localiza faixa de linhas/colunas candidatas (`:39-47`); (4) para cada célula candidata, verifica se a fórmula contém simultaneamente os marcadores de texto `R3C="FÓRMULA"` e `R3C="DADO"` (checagem de "tipo de célula" codificada dentro da própria fórmula da planilha); se sim, marca início/fim de um bloco contíguo e reaplica `F_Padrao` a esse bloco (`:59-82`); (5) `PopUp_Tempo_Processamento`; (6) `Ativar_Tudo`.
14. **Pós-condições:** todas as células candidatas identificadas como "deveriam ser fórmula" voltam a conter a fórmula padrão `F_Padrao`.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Convenção de "célula mágica" (`R3C="FÓRMULA"`/`R3C="DADO"` embutidos na própria fórmula da planilha) é frágil e mistura metadado de estrutura com lógica de cálculo — típico ponto de quebra silenciosa se a linha 3 de referência for alterada. `On Error GoTo fim` genérico (`:22,92`) — captura qualquer erro e exibe MsgBox crítico, mas sem detalhar qual célula falhou. Barra de status como único feedback de progresso, sem log persistente.

### 17. Atualizar_Validacao_Linhas_Geral

1. **Nome completo:** `Atualizar_Validacao_Linhas_Geral`
2. **Módulo:** `Front_Processos.bas`
3. **Tipo:** Sub
4. **Escopo:** `Public` (implícito) — `Front_Processos.bas:104`
5. **Objetivo (negócio, conforme intenção declarada no código):** Reconstruir listas de validação (dropdowns) de colunas conhecidas (EMPRESA, ORGANIC, VISAO, IFRS_CONTABIL, PROFORMA) na planilha ativa, resolvendo o nome do intervalo de validação correspondente a cada uma.
6. **Quem chama (evidência):** Nenhuma chamada encontrada no `grep` completo do diretório — mesmo padrão de botão genérico dedicado, fora do dump de texto.
7. **Procedimentos chamados:** `Desligar_Tudo` (`:109`), `Ativar_Tudo` (`:139`) — `Auxiliar.bas`.
8. **Parâmetros:** nenhum (opera sobre `ActiveSheet`).
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `sh` (=`ActiveSheet`), `Chave_Validacao` (nome do intervalo de validação resolvido por coluna, ex. `"CB_Empresa_DropComb"`, `"CB_Organic"`, `"CB_Visao"`, `"CB_IFRS_CONTABIL_DropComb"`, `"CB_PROFORMA_DropComb"`).
11. **Abas/intervalos/arquivos acessados:** `ActiveSheet` — linha de cabeçalho localizada por coluna iniciando em "Aux".
12. **Pré-condições:** `ActiveSheet` deve seguir a estrutura padrão com linha de cabeçalho identificável por "Aux*".
13. **Passos principais:** (1) `Desligar_Tudo`; (2) localiza `Lin_Cabecalho`/`Ult_Col`/`Ult_Lin`; (3) loop por todas as colunas do cabeçalho, resolvendo `Chave_Validacao` conforme o nome da coluna (`:122-126`); (4) **bloco `If/Else` de aplicação da validação está vazio** — apenas comentários (`:128-135`), nenhuma ação de fato ocorre; (5) `Ativar_Tudo`; (6) `MsgBox "Processo concluído com sucesso!"` (`:141`).
14. **Pós-condições (reais, não as pretendidas):** nenhuma alteração efetiva nas listas de validação da planilha — o procedimento identifica as colunas e os nomes de intervalo corretos, mas não os aplica.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** **Achado crítico**: a função exibe `MsgBox "Processo concluído com sucesso!"` (`:141`) **mesmo não fazendo nada** — o corpo do `Else` que aplicaria a validação está vazio (`:132-135`, apenas comentário `'Não faz nada, apenas ignora.` no ramo IF, e nada no ramo ELSE). Isto é **enganoso para o usuário de negócio**, que acredita que a validação foi atualizada. Recomenda-se tratar como funcionalidade incompleta/abandonada — **[VALIDAR COM O NEGÓCIO]** se este botão ainda está exposto na UI e se algum processo dependeu dele silenciosamente não fazendo efeito.

---

## Módulo `BackupCodigo_MainResults.bas`

### 18. Worksheet_Change

1. **Nome completo:** `Worksheet_Change`
2. **Módulo:** `BackupCodigo_MainResults.bas`
3. **Tipo:** Sub (procedimento de evento — assinatura `Private Sub Worksheet_Change(ByVal Target As Range)`, idêntica à de um evento nativo de objeto Planilha do Excel)
4. **Escopo:** `Private` — `BackupCodigo_MainResults.bas:2`
5. **Objetivo (negócio, conforme lógica interna):** Validar automaticamente em cascata os dropdowns de aberturas da planilha "Main Results": ao alterar uma célula de abertura, restringir dinamicamente as opções válidas das colunas dependentes seguintes, com base nas combinações reais cadastradas em `Sheet25`.
6. **Quem chama (evidência) — ACHADO ESTRUTURAL CRÍTICO:** `grep -rn "Worksheet_Change" vba_dump_tmp/` retorna **apenas a própria linha de declaração** (`BackupCodigo_MainResults.bas:2`) — nenhuma outra ocorrência em todo o diretório (33 arquivos `.bas`/`.cls`/`.frm`). Isso é evidência formal de que **nenhum procedimento do código VBA chama este evento explicitamente** — o que é esperado para um `Worksheet_Change`, pois ele deveria ser disparado **automaticamente pelo Excel** quando uma célula é alterada. Porém, esse disparo automático **só ocorre se o procedimento estiver no módulo de código do objeto Planilha correspondente** (ex.: dentro do `.cls` de `Sheet3`), e **não** dentro de um módulo `.bas` padrão como `BackupCodigo_MainResults.bas`. Comparando com os 33 arquivos do dump: todos os `Worksheet_Change`/`Worksheet_BeforeDoubleClick` legítimos e ativos encontrados neste projeto estão em arquivos `.cls` (ex.: `Sheet8.cls:35`, `Worksheet_BeforeDoubleClick`); **nenhum outro `Worksheet_Change` ativo foi encontrado em nenhum `.cls`** do dump. O nome do próprio módulo — **"BackupCodigo"** ("código de backup") — reforça essa leitura. **Conclusão formal:** neste estado (isolado em módulo `.bas`), o procedimento está **estruturalmente inerte** — não é chamado por nenhuma outra rotina (confirmado por grep) e não está posicionado onde o Excel o dispararia automaticamente (confirmado pela ausência de qualquer `Worksheet_Change` equivalente nos arquivos `.cls` do dump). **[VALIDAR COM O NEGÓCIO]**: confirmar se este código foi de fato desligado intencionalmente (ex. por problema de performance) ou se existe uma cópia idêntica/similar viva no code-behind real de `Sheet3` que não foi capturada por este dump de texto (o dump cobre os módulos conforme extraídos; uma divergência entre o projeto VBA "vivo" no `.xlsb` e este dump não pode ser descartada sem inspeção direta do arquivo binário).
7. **Procedimentos chamados:** `Processo_Validaca_Linha` (`:113`, procedimento nº 19, mesmo módulo) — condicionado a `Valida_Hierarquia = "SIM"` (célula `L32`); `Desligar_Tudo`/`Ativar_Tudo` (`Auxiliar.bas`, `:72,118,129`).
8. **Parâmetros:** `Target As Range` (assinatura padrão do evento `Worksheet_Change`).
9. **Retorno:** N/A (Sub/evento).
10. **Variáveis/objetos relevantes:** `Valida_Hierarquia` (lida de `Me.Range("L32")`, controla se a validação roda), `Lin_Cabecalho`, `Col_Inicio`/`Col_Fim` (faixa de colunas de abertura, de "Abertura_2" até a última coluna preenchida na linha de cabeçalho), `Chaves(2000)` (array fixo — ver risco), `Col_Diretorias`.
11. **Abas/intervalos/arquivos acessados:** `Me` (a planilha dona do código — presumivelmente Main Results/Sheet3, dado o nome do módulo); leitura de `Sheet25` (via `Sh_Base`, passado a `Processo_Validaca_Linha`) para resolver combinações válidas.
12. **Pré-condições:** célula `L32` (rotulada "Ligar validação de linhas automáticas", localizada por `Cells.Find` em `:13-15`) deve conter `"SIM"`; a célula alterada (`Target`) deve estar dentro do range de colunas de abertura (`Col_Inicio-1` a `Col_Fim`) e em linha abaixo do cabeçalho.
13. **Passos principais:** (1) localiza cabeçalho e checa `L32` (`:13-18`, `Exit Sub` se não for "SIM"); (2) determina faixa de colunas de abertura e de diretorias; (3) determina se `Target` é uma célula única ou um range (`:52-62`); (4) se dentro da faixa válida, para cada linha alterada: monta array de chaves (coluna+valor já preenchido, `:81-89`), busca linhas de `Sh_Base` (Sheet25) cujos valores batem com todas as chaves (`:91-109`); (5) chama `Processo_Validaca_Linha` para aplicar a validação resultante nas colunas ainda vazias da linha.
14. **Pós-condições (quando ativo):** colunas de abertura ainda vazias na(s) linha(s) alterada(s) recebem uma lista de validação restrita às combinações realmente existentes em `Sheet25`.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Array fixo `Dim Chaves(2000)` (`:6`) — limite hardcoded; estouraria (erro de índice) se uma linha tivesse mais de 2000 colunas de abertura preenchidas simultaneamente (cenário improvável, mas não validado/tratado). Complexidade O(n×m): para cada linha alterada, varre toda `Sh_Base` (Sheet25) múltiplas vezes (`:92-109`) — se ativo e disparado a cada `Change` (inclusive colagens de múltiplas linhas), pode ser extremamente lento; **hipótese razoável [VALIDAR COM O NEGÓCIO]** para explicar por que este código pode ter sido desativado (ver item 6). `On Error GoTo Termina` genérico (`:20`), sem log. Este é o achado mais importante do módulo para a documentação formal: **não presumir que a validação automática de hierarquia é comportamento vivo do sistema atual sem confirmação do negócio.**

### 19. Processo_Validaca_Linha

1. **Nome completo:** `Processo_Validaca_Linha`
2. **Módulo:** `BackupCodigo_MainResults.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `BackupCodigo_MainResults.bas:136`
5. **Objetivo (negócio):** Para cada coluna de abertura ainda vazia na linha, calcular a lista de valores possíveis (interseção das linhas de `Sheet25` que batem com a chave já preenchida) e aplicá-la como `Data Validation` (lista suspensa) na célula.
6. **Quem chama (evidência):** `BackupCodigo_MainResults.bas:113` (dentro de `Worksheet_Change`, mesmo módulo). `Private` — nenhuma chamada externa possível/encontrada. Herda, portanto, o mesmo status estrutural inerte do procedimento nº 18 (só é alcançado se `Worksheet_Change` disparar, o que hoje não ocorre automaticamente).
7. **Procedimentos chamados:** `Ordenar_Lista` (`:170` — resolve para a versão **local/Private** deste mesmo módulo, `:188`, procedimento nº 20, por precedência de escopo do VBA: um procedimento `Private` de mesmo nome no módulo chamador tem prioridade sobre uma versão `Public` de outro módulo, como a existente em `Auxiliar.bas:1096`).
8. **Parâmetros:** `Sh_Base` (=Sheet25, passado pelo chamador), `Lin_Cabecalho`, `Lin`, `Linhas_Chaves` (string de números de linha separados por vírgula), `Col_Inicio`, `Col_Fim`.
9. **Retorno:** N/A (Sub).
10. **Variáveis/objetos relevantes:** `Filtro_Novo`/`Filtro_Novo_Ordenado` (lista de valores válidos, antes/depois de ordenar), `MyArray` (array de valores após `Split`).
11. **Abas/intervalos/arquivos acessados:** `Me` (planilha corrente, para `.Cells(Lin,y).Validation`); `Sh_Base` (Sheet25) para resolver os valores possíveis por coluna.
12. **Pré-condições:** `Linhas_Chaves` deve ter sido calculada previamente pelo chamador (`Worksheet_Change`); coluna de validação (`Col_Validacao`) deve existir como cabeçalho em `Sh_Base`.
13. **Passos principais:** (1) para cada coluna `y` de `Col_Inicio` a `Col_Fim` cuja célula esteja vazia (`:143`); (2) se `Linhas_Chaves=""`, `Filtro_Novo="Total"`; senão monta lista de valores distintos encontrados nas linhas de `Sh_Base` indicadas (`:148-158`); (3) ordena via `Ordenar_Lista` (exceto quando é só "Total"); (4) `Me.Cells(Lin,y).Validation.Delete` seguido de `.Add Type:=xlValidateList ... Formula1:=Filtro_Novo_Ordenado` (`:174-179`).
14. **Pós-condições:** cada célula de abertura vazia na linha recebe uma `Data Validation` do tipo lista, com os valores possíveis calculados.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Sem tratamento de erro próprio. Depende de `Sh_Base` (Sheet25) estar atualizada (gerada por `Extrair_Valid_Lin`, procedimento nº 6) — dependência de ordem de pipeline. Estruturalmente inerte hoje pelas mesmas razões do procedimento nº 18.

### 20. Ordenar_Lista (versão local de BackupCodigo_MainResults.bas)

1. **Nome completo:** `Ordenar_Lista`
2. **Módulo:** `BackupCodigo_MainResults.bas`
3. **Tipo:** Sub
4. **Escopo:** `Private` — `BackupCodigo_MainResults.bas:188`
5. **Objetivo (negócio):** Ordenar alfabeticamente uma lista de valores separados por vírgula, garantindo que o valor "Total" fique sempre por último — usado para apresentar as opções de dropdown de forma previsível ao usuário.
6. **Quem chama (evidência):** `BackupCodigo_MainResults.bas:170` (dentro de `Processo_Validaca_Linha`, mesmo módulo — resolve para esta versão local por regra de escopo VBA, não para `Auxiliar.bas:1096`). **Nota de desambiguação formal:** existe uma **terceira** implementação homônima, `Public Function Ordenar_Lista` em `Auxiliar.bas:1096`, chamada separadamente por `Form_Exportacao.frm:563` — são três procedimentos de código distintos (dois `Sub`/`Function` diferentes, mais uma versão inline em `Lista_Validacao_Ajustes.bas:110-129` sem nome de procedimento próprio) implementando a mesma lógica de bubble-sort com "Total" por último, sem reuso entre módulos — oportunidade de consolidação clara para a reescrita.
7. **Procedimentos chamados:** nenhum (bubble sort manual inline).
8. **Parâmetros:** `Filtro_Novo` (entrada, string separada por vírgula), `Filtro_Novo_Ordenado` (saída, por `ByRef` default).
9. **Retorno:** N/A (Sub); resultado devolvido via parâmetro `ByRef` `Filtro_Novo_Ordenado`.
10. **Variáveis/objetos relevantes:** `MyArray` (array via `Split`), `TempTxt1`/`TempTxt2` (swap temporário).
11. **Abas/intervalos/arquivos acessados:** nenhuma (opera só sobre strings em memória).
12. **Pré-condições:** `Filtro_Novo` deve ser uma string válida separada por vírgulas.
13. **Passos principais:** (1) `Split(Filtro_Novo, ",")`; (2) bubble sort duplo-loop: troca posições se o item à direita for "Total" OU (não for "Total" E for alfabeticamente menor) — efetivamente ordena alfabeticamente com "Total" forçado ao final (`:190-200`); (3) remonta string separada por vírgula (`:202-206`).
14. **Pós-condições:** `Filtro_Novo_Ordenado` contém a mesma lista, ordenada alfabeticamente, com "Total" (se presente) na última posição.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Bubble sort O(n²) — aceitável para listas pequenas de dropdown, mas é código duplicado (ver item 6) sem qualquer reuso, aumentando a superfície de manutenção. Sem tratamento de erro.

---

## Módulo `fx_IFRS16.bas`

### 21. UPDATE_Tratar_IFRS16

1. **Nome completo:** `UPDATE_Tratar_IFRS16`
2. **Módulo:** `fx_IFRS16.bas`
3. **Tipo:** Function (retorno não usado pelo chamador — `Call UPDATE_Tratar_IFRS16(Chave)`, `Extracao_Base_1009.bas:16`)
4. **Escopo:** `Public` (implícito, sem modificador) — `fx_IFRS16.bas:3`
5. **Objetivo (negócio):** Implementar a **regra contábil de tratamento IFRS16** (arrendamentos/leasing): identificar, na base histórica, as linhas de origem "Base_1009" cujo Centro de Custo pertence à lista de contratos tratados por IFRS16, reclassificá-las para a Classe Custo/Centro de Custo de destino definidos em `Sup_Linhas`, inverter o sinal dos valores e substituir as linhas originais pelas tratadas.
6. **Quem chama (evidência):** `Extracao_Base_1009.bas:16` (`Call UPDATE_Tratar_IFRS16(Chave)`, onde `Chave="Base_1009"`, `:5`) — chamada ativa, dentro de `Extrair_Base_1009`. `Extracao_SQL_Hubble.bas:16` contém `'Call UPDATE_Tratar_IFRS16` — **chamada comentada/inativa**, evidência de que já se cogitou (ou se descontinuou) rodar este tratamento também após a extração da base Hubble.
7. **Procedimentos chamados:** `Limpar_Base_Historica` (`:105`, `Auxiliar.bas:496`, chamado com `Campo="Fonte", Chave="DELETAR"`); `Form_IFRS_Contabil` (`:123`, `Aux_Formulas_Base.bas:753`); `Form_Empresa` (`:124`, `Aux_Formulas_Base.bas:174`); `Reclassificar_Combinacoes_Empresas_TK` (`:125`, `TK_Functions.bas:877`); `Calcular_Comb_Meses_Intervalo_Linha_TK` (`:126`, `TK_Functions.bas:686`); `Reclassificar_Combinacoes_Proforma_TK` (`:127`, `TK_Functions.bas:797`); `Reclassificar_Combinacoes_IFRS_Contabil_TK` (`:128`, `TK_Functions.bas:716`).
8. **Parâmetros:** `vBase_a_procurar` (string — prefixo do campo "Fonte" a filtrar; valor real observado: `"Base_1009"`, via `Chave` em `Extracao_Base_1009.bas:5,16`).
9. **Retorno:** nenhum valor de retorno é atribuído ao nome da função nem consumido pelo chamador.
10. **Variáveis/objetos relevantes:** `Sh_Base` (=`Sheet3`), `Sh_Extracao` (=`Sheet8`), `Sh_SupLinhas` (=`Sheet15`), `Sh_AuxIFRS16` (=`Sheet28`), `arr_Linhas_a_apagar` (ArrayList — criado mas não efetivamente usado para a exclusão real, ver item 15), `xx` (contador de linhas puladas por erro de mapeamento), `vCC_DE`/`vCC_PARA`, `vCDC_DE`/`vCDC_PARA`, `vEmpresa`, `vChave`.
11. **Abas/intervalos/arquivos acessados:** Leitura: `Sh_Extracao!I25` (flag "Sim"/"Não", `:12`); `Sh_Base` (Sheet3), range `A6:AV<última linha>` (`:22-23`), colunas relevantes por número: **coluna 2** ("Fonte", filtro de prefixo e marcação "DELETAR"), **coluna 8** (Empresa, `:47`), **coluna 23** (Classe Custo "DE", `:26,40,45`), **coluna 31/32/33** (chave Diretoria N1/N2/N3, `:48`), **coluna 35** (Centro de Custo "DE", `:40,46`). Lookup em `Sh_SupLinhas` (Sheet15): **coluna 71** (chave de busca do CC "DE", `Match(vCC_DE, Sh_SupLinhas.Columns(71),0)`, `:49`), **coluna 73** (CC "PARA", `Sh_SupLinhas.Cells(vLin_Alvo,73)`, `:53`; também usada como coluna de busca em outro `Match`, `:27`), **coluna 15** (chave de busca do CDC, formato `"CTCEL" & vChave`, `:66`), **coluna 22** (CDC "PARA", `:67`), célula **`BX1`** (multiplicador de sinal, setado para `-1`, `:77-78`). Escrita: `Sh_AuxIFRS16` (Sheet28, staging — colunas 1 a 48 copiadas de `Sh_Base`, mais ajuste nas colunas 2/23/35 e colunas de meses a partir de "AJ"), de volta em `Sh_Base` (linhas reinseridas ao final).
12. **Pré-condições:** `Sh_Extracao!I25 = "Sim"`; `Sh_SupLinhas` colunas 71/73/15/22 previamente preenchidas com o De-Para de IFRS16 — **não populadas por nenhuma macro de refresh analisada** (ver item 15/risco crítico); célula `Sh_SupLinhas!BX1` disponível para receber o valor `-1`.
13. **Passos principais:** (1) checa `I25="Sim"` (`:12`); (2) limpa e prepara `Sh_AuxIFRS16` (`:15-20`); (3) 1ª passada sobre `Sh_Base`: para linhas cujo "Fonte" comece com `vBase_a_procurar`, resolve `vCC_DE`→`Sh_SupLinhas.Columns(73)` via `Match`; se achar, marca `Sh_Base!B<linha>="DELETAR"` (`:24-33`); (4) 2ª passada: para as mesmas linhas, valida colunas-chave sem `#N/A` (senão pula e incrementa `xx`, `:40-43`); resolve CC "DE"→"PARA" via `Sh_SupLinhas.Columns(71)`→`73` (`:45-53`); copia linha inteira para `Sh_AuxIFRS16`, ajusta coluna 23 (CC PARA) e sufixa coluna 2 (Fonte) com `"_IFRS16 Tratado"` (`:55-63`); se Empresa ≠ "5G" e ≠ "METIS_CZ", resolve CDC "DE"→"PARA" via chave `"CTCEL"&vChave` em `Sh_SupLinhas.Columns(15)`→`22` (`:65-69`); (5) multiplica valores mensais (a partir da coluna "AJ") por `-1` lido de `Sh_SupLinhas!BX1` (`:76-86`); (6) copia linhas tratadas de `Sh_AuxIFRS16` de volta ao final de `Sh_Base` (`:90-99`); (7) exclui linhas originais via `Limpar_Base_Historica(Campo:="Fonte", Chave:="DELETAR")` (`:103-105`); (8) reaplica fórmulas de enriquecimento (`Form_IFRS_Contabil`, `Form_Empresa`) e reclassificações genéricas (`:122-128`); (9) `Sh_AuxIFRS16.Visible=False`; `MsgBox "Linhas IFRS16 tratadas com sucesso!"` (`:129-130`).
14. **Pós-condições:** linhas de origem "Base_1009" com Centro de Custo IFRS16 mapeado ficam reclassificadas (CC/CDC de destino, Fonte sufixada, valores com sinal invertido) e substituem as linhas originais na base histórica (`Sh_Base`); `Sh_AuxIFRS16` volta a ficar oculta.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** **Achado crítico [VALIDAR COM O NEGÓCIO]**: `grep` confirma que as colunas 71/73/15/22 de `Sh_SupLinhas` e a célula `BX1` **só são referenciadas neste módulo em todo o dump** (nenhuma das macros de refresh analisadas — `Refresh_Sup_Linhas.bas`, `Refresh_De_X_Para.bas` — escreve nessas posições) — indício forte de que o De-Para de IFRS16 é **mantido manualmente por um usuário de negócio diretamente na planilha `Sup_Linhas`**, fora de qualquer automação rastreável neste código-fonte. Ponto de risco alto para a reescrita (ver seção D). Exceção `Empresa <> "5G" E <> "METIS_CZ"` **hardcoded** no VBA (`:65`). Linhas com erro de mapeamento (`#N/A` em CC/CDC/Empresa/Chave) são **silenciosamente puladas**, só contabilizadas em `xx`, sem log de qual CC/linha falhou (`:40-43,133-137`). `arr_Linhas_a_apagar` (ArrayList) é criado e populado (`:54`) mas a exclusão real é feita por `Limpar_Base_Historica` via filtro de texto "DELETAR", não pela lista — **código morto residual**. Uso de índices de coluna fixos por número (23,35,71,73,15,22, etc.) sem indireção por nome de cabeçalho — diferente de outros módulos do sistema que usam `fn.Match` — extremamente frágil a reordenação de colunas nas planilhas envolvidas. Loop duplo sobre a mesma base grande (`:24-33` e `:35-73`) — ineficiente. Mensagens ao usuário: sucesso (`:130`), tratamento desabilitado (`:136`), aviso de N linhas puladas por erro (`:134`).

### 22. fx_Tratar_KPI

1. **Nome completo:** `fx_Tratar_KPI`
2. **Módulo:** `fx_IFRS16.bas`
3. **Tipo:** Function (esta, diferente das demais do cluster, **usa de fato** o valor de retorno — atribui a `fx_Tratar_KPI` nas duas ramificações, `:146,148`)
4. **Escopo:** `Public` (implícito) — `fx_IFRS16.bas:143`
5. **Objetivo (negócio):** Resolver o "KPI" efetivo de uma linha: usa o KPI de destino explícito se informado; senão, extrai o KPI a partir da string combinada `"KPI > Versão"`.
6. **Quem chama (evidência):** `grep -rn "fx_Tratar_KPI" vba_dump_tmp/` retorna apenas as linhas internas da própria função (`:143,146,148`) — **nenhuma chamada `Call`/uso em outro procedimento VBA foi encontrada em todo o dump**. **[NÃO IDENTIFICADO]** um chamador VBA explícito. **Inferência [VALIDAR COM O NEGÓCIO]**: a assinatura (`vKPI_Versao`, `vKPI_Destino`) e a convenção de nome `fx_` (mesmo prefixo do módulo `fx_IFRS16`) são compatíveis com uso como **User Defined Function (UDF)** chamada diretamente em fórmula de célula do Excel (ex. `=fx_Tratar_KPI(...)`) — cenário plausível dado que `Sup_Linhas` (Sheet15) possui colunas "KPI_VERSAO"/"KPI"/"VERSÃO" geradas por `Atualizar_Lista_KPI_Versao` (procedimento nº 7). **[NÃO ACESSÍVEL]**: o conteúdo de fórmulas de célula não está incluído neste dump de texto (só código `.bas`/`.cls`/`.frm`), portanto essa hipótese não pôde ser confirmada nem refutada por grep.
7. **Procedimentos chamados:** nenhum (apenas `Split`, função nativa VBA).
8. **Parâmetros:** `vKPI_Versao` (string combinada `"KPI > Versão"`), `vKPI_Destino` (override opcional).
9. **Retorno:** String — o KPI de destino (`vKPI_Destino`, se `Len(vKPI_Destino) > 0`) ou a parte antes de `" > "` de `vKPI_Versao` (`Split(vKPI_Versao," > ")(0)`).
10. **Variáveis/objetos relevantes:** nenhuma além dos parâmetros.
11. **Abas/intervalos/arquivos acessados:** nenhum diretamente (função pura sobre os parâmetros recebidos) — se usada como UDF, o "acesso" à planilha ocorre indiretamente via a célula-fórmula que a invoca (fora do escopo verificável deste dump).
12. **Pré-condições:** `vKPI_Versao`, se `vKPI_Destino` estiver vazio, deve conter o separador `" > "` (senão `Split(...)(0)` funciona mas indica ausência do padrão esperado; `Split(...)(1)`, usado pela função irmã nº 23, geraria erro de índice se o separador não existir).
13. **Passos principais:** (1) `If Len(vKPI_Destino) > 0 Then fx_Tratar_KPI = vKPI_Destino`; (2) `Else fx_Tratar_KPI = Split(vKPI_Versao, " > ")(0)`.
14. **Pós-condições:** retorna o nome do KPI, priorizando um valor de destino explícito sobre o valor derivado da string combinada.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** Sem tratamento de erro — se `vKPI_Versao` não contiver `" > "`, `Split(...)(0)` ainda retorna algo (a string inteira), mas de forma potencialmente incorreta sem aviso. Função pura, sem efeitos colaterais em planilha.

### 23. fx_Tratar_Versao

1. **Nome completo:** `fx_Tratar_Versao`
2. **Módulo:** `fx_IFRS16.bas`
3. **Tipo:** Function (usa o retorno de fato, mesma observação do procedimento nº 22)
4. **Escopo:** `Public` (implícito) — `fx_IFRS16.bas:153`
5. **Objetivo (negócio):** Resolver a "Versão" efetiva de uma linha: usa a versão de destino explícita se informada; senão, extrai a versão a partir da string combinada `"KPI > Versão"`.
6. **Quem chama (evidência):** `grep -rn "fx_Tratar_Versao" vba_dump_tmp/` retorna apenas as linhas internas da própria função (`:153,156,158`) — **nenhuma chamada externa encontrada**. Mesma situação e mesma hipótese `[VALIDAR COM O NEGÓCIO]`/`[NÃO ACESSÍVEL]` do procedimento nº 22 (provável UDF usada em fórmula de célula, não confirmável neste dump).
7. **Procedimentos chamados:** nenhum (apenas `Split`).
8. **Parâmetros:** `vKPI_Versao`, `vVersao_Destino` (override opcional).
9. **Retorno:** String — a versão de destino (`vVersao_Destino`, se preenchida) ou a parte após `" > "` de `vKPI_Versao` (`Split(vKPI_Versao," > ")(1)`).
10. **Variáveis/objetos relevantes:** nenhuma além dos parâmetros.
11. **Abas/intervalos/arquivos acessados:** nenhum diretamente (mesma observação do procedimento nº 22).
12. **Pré-condições:** se `vVersao_Destino` estiver vazio, `vKPI_Versao` **deve** conter o separador `" > "`, pois `Split(...)(1)` gera erro de índice fora do limite (`Subscript out of range`) se o separador não existir — diferente da função irmã (`fx_Tratar_KPI`, que usa índice `(0)`, sempre válido mesmo sem separador).
13. **Passos principais:** (1) `If Len(vVersao_Destino) > 0 Then fx_Tratar_Versao = vVersao_Destino`; (2) `Else fx_Tratar_Versao = Split(vKPI_Versao, " > ")(1)`.
14. **Pós-condições:** retorna o nome da versão, priorizando um valor de destino explícito.
15. **Efeitos colaterais/tratamento de erro/mensagens/risco/evidência:** **Sem tratamento de erro** — ao contrário da função irmã (nº 22), esta é vulnerável a erro de execução (`Subscript out of range`) se `vKPI_Versao` não contiver `" > "` e `vVersao_Destino` estiver vazio. Risco a testar explicitamente na reescrita (caso de borda: KPI sem versão cadastrada).

---


## 10.4 Cluster UI / Forms (35 procedimentos/eventos + 30 planilhas sem código)
# (A) CATÁLOGO POR PROCEDIMENTO/EVENTO

## A.1 — Form_Importacao.frm (11 procedimentos)

### 1. `B_Cancel_Click` — Form_Importacao.frm:9
1. **Nome completo:** `B_Cancel_Click`
2. **Módulo:** Form_Importacao.frm (UserForm)
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Cancelar a operação de importação em andamento; fechar o arquivo externo que havia sido aberto para leitura, se houver, e devolver o controle à planilha principal.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no botão `B_Cancel`). Não há chamada explícita no código (evidência: `grep -rn "B_Cancel_Click"` retorna somente a declaração, Form_Importacao.frm:9).
7. **Procedimentos chamados:** `Ativar_Tudo` (Auxiliar.bas:69) — linha 29.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.TB_Arq` (TextBox, inferido), `Me.TB_Diret` (TextBox, inferido). Variáveis globais implícitas `Arq`, `Diret` (não declaradas com `Dim` — escopo módulo/projeto por omissão do VBA).
11. **Abas/intervalos/arquivos acessados:** `Windows(Arq)` — a janela do arquivo externo aberto pelo `B_Ler_Fronts_Click`, se existir. `ThisWorkbook` (reativação).
12. **Pré-condições:** Form aberto. `Arq` pode estar vazio (usuário cancelou antes de escolher arquivo) ou preenchido (arquivo já aberto para leitura).
13. **Passos principais:** (i) lê `TB_Arq`/`TB_Diret`; (ii) se `Arq` não vazio, tenta ativar a janela do arquivo e pergunta ao usuário (MsgBox Sim/Não) se deseja fechá-la sem salvar; (iii) descarrega o form (`Unload Me`); (iv) reativa `ThisWorkbook`; (v) chama `Ativar_Tudo` para restaurar `ScreenUpdating`/`DisplayAlerts`/`EnableEvents`/`StatusBar`.
14. **Pós-condições:** Form fechado; arquivo externo fechado (se o usuário confirmou) sem salvar (`SaveChanges:=False`); tela do Excel reativada.
15. **Efeitos colaterais / erro / mensagens / risco:** Usa `On Error GoTo Fim_Cod` para tolerar o caso de a janela do arquivo já não existir mais (silencioso, sem mensagem de erro ao usuário). Mensagem de confirmação: "Deseja fechar o arquivo [...] já aberto!?" (linha 19). Risco baixo — comportamento defensivo adequado. `SaveChanges:=False` descarta silenciosamente qualquer alteração feita no arquivo externo (esperado, pois é aberto `ReadOnly:=True` no fluxo normal, mas se o usuário editou manualmente o arquivo aberto, a alteração é perdida sem aviso específico).

### 2. `B_Incluir_Click` — Form_Importacao.frm:33
1. **Nome completo:** `B_Incluir_Click`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Mover os layouts/abas marcados na lista geral para a lista de seleção ("carrinho" de itens a importar), evitando duplicados.
6. **Quem chama / gatilho:** Gatilho direto por clique do usuário no botão `B_Incluir`. Também chamado programaticamente por `LB_Layouts_Geral_DblClick` (Form_Importacao.frm:732, `Call B_Incluir_Click`) — duplo-clique num item da lista geral produz o mesmo efeito.
7. **Procedimentos chamados:** Nenhum (apenas manipulação de coleção `ListBox`).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Geral` (ListBox, multi-seleção — evidência: `.Selected(x)` em loop, linha 36), `Me.LB_Layouts_Select` (ListBox destino).
11. **Abas/intervalos/arquivos acessados:** Nenhum (opera só sobre os `ListBox` em memória).
12. **Pré-condições:** `LB_Layouts_Geral` populado (normalmente após `B_Ler_Fronts_Click`).
13. **Passos principais:** Para cada item marcado (`Selected = True`) em `LB_Layouts_Geral`: verifica (case-insensitive, via `UCase`) se já existe em `LB_Layouts_Select`; se não existir, adiciona (`AddItem`); em seguida desmarca o item na lista de origem.
14. **Pós-condições:** `LB_Layouts_Select` contém a união dos itens previamente selecionados com os recém-incluídos, sem duplicados; itens de origem ficam desmarcados.
15. **Efeitos colaterais / erro / mensagens / risco:** Duas linhas de ajuste de altura dos ListBox (`.Height = 118`) estão **comentadas** (linhas 47–48) — ou seja, o redimensionamento visual que existia em versão anterior foi desativado; comportamento atual da altura dos controles é herdado do design estático do form ([NÃO ACESSÍVEL] o valor atual exato). Nenhuma mensagem de erro; nenhuma validação de limite de itens.

### 3. `B_Excluir_Click` — Form_Importacao.frm:52
1. **Nome completo:** `B_Excluir_Click`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Remover da lista de seleção os itens marcados pelo usuário.
6. **Quem chama / gatilho:** Gatilho por clique do usuário no botão `B_Excluir`. Também chamado por `LB_Layouts_Select_DblClick` (Form_Importacao.frm:738) — duplo-clique num item da lista de seleção remove esse item.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Select` (ListBox).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** `LB_Layouts_Select` com ao menos um item marcado (senão o laço simplesmente não remove nada).
13. **Passos principais:** Laço com `GoTo Inicio` (repetição manual, não `For` clássico decrescente): percorre `LB_Layouts_Select`, remove (`RemoveItem`) o primeiro item marcado encontrado e reinicia a varredura do zero até não sobrar nenhum marcado.
14. **Pós-condições:** Todos os itens que estavam marcados são removidos da lista de seleção.
15. **Efeitos colaterais / erro / mensagens / risco: ** Padrão de remoção via `GoTo Inicio` é funcional mas ineficiente em O(n²) para muitos itens marcados simultaneamente — irrelevante em volume normal de uso (dezenas de abas), mas é um "code smell" a não replicar na reescrita. Mesmas linhas de redimensionamento comentadas (linhas 64–65) do item anterior.

### 4. `B_Ler_Fronts_Click` — Form_Importacao.frm:69
1. **Nome completo:** `B_Ler_Fronts_Click`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Abrir (ou ativar, se já aberto) o arquivo externo escolhido pelo usuário e listar todas as suas abas/layouts disponíveis para importação (excluindo abas de suporte internas).
6. **Quem chama / gatilho:** Gatilho por clique do usuário no botão `B_Ler_Fronts`. Não referenciado programaticamente em nenhum outro ponto (evidência: grep só retorna a declaração e uma linha **comentada** de exemplo de auto-preenchimento para um usuário específico — Form_Importacao.frm:768).
7. **Procedimentos chamados:** `Carregar_Sheets_Suporte` (Auxiliar.bas:582), `PopUp_Tempo_Processamento` (Auxiliar.bas:138).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.TB_Arq`, `Me.TB_Diret` (leitura), `Me.LB_Layouts_Geral` (populado), `Me.LB_Layouts_Select` (limpo). Variável global `Sheets_Suporte` (definida em `Auxiliar.bas:2` como `Global`).
11. **Abas/intervalos/arquivos acessados:** Arquivo externo indicado por `Diret & Arq` (abertura via `Workbooks.Open`, `ReadOnly:=True`, `UpdateLinks:=False`); `Sheet3` (usado para filtrar seu próprio nome da lista de suporte, linha 86).
12. **Pré-condições:** `TB_Arq` deve estar preenchido (validado — ver campo 15).
13. **Passos principais:** (i) limpa as duas ListBox; (ii) valida se há arquivo selecionado; (iii) marca `Inicio` (para medição de tempo); (iv) chama `Carregar_Sheets_Suporte` e remove `Sheet3.Name` da lista de nomes de suporte; (v) tenta ativar a janela do arquivo (`Windows(Arq).Activate`) — se falhar, cai no label `Abrir_Arq` e abre o arquivo via `Workbooks.Open`; (vi) percorre todas as `Worksheets` do arquivo aberto e adiciona à `LB_Layouts_Geral` os nomes que não estão na lista de sheets de suporte; (vii) reativa `ThisWorkbook` (duas vezes, redundante); (viii) chama `PopUp_Tempo_Processamento`.
14. **Pós-condições:** `LB_Layouts_Geral` populada com os nomes das abas elegíveis para importação do arquivo externo, que permanece aberto (em memória, modo leitura) até `B_Ok_Click` ou `B_Cancel_Click`.
15. **Efeitos colaterais / erro / mensagens / risco:** Validação: `MsgBox "Selecione um arquivo para prosseguir..."` (vbCritical) se `TB_Arq` vazio (linha 79). Tratamento de erro ao abrir arquivo: `MsgBox "Ocorreu um erro ao abrir este arquivo!"` (vbCritical, linha 120) — mensagem genérica, não expõe a causa real (nome/erro do sistema). Uso de `On Error GoTo -1` para limpar o handler de erro ativo antes de reatribuir outro (padrão correto em VBA, mas pouco comum — indica desenvolvedor experiente). `PopUp_Tempo_Processamento` sempre reporta "Extração concluída com sucesso!!!" mesmo aqui, onde a operação é apenas uma leitura de nomes de abas — mensagem genérica reaproveitada de outro contexto, texto pode confundir o usuário (RN relacionada, ver seção B).

### 5. `B_Ok_Click` — Form_Importacao.frm:125
1. **Nome completo:** `B_Ok_Click`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Executar a importação efetiva dos layouts/abas escolhidos pelo usuário, adaptando cada aba ao padrão interno do Quick Data (cabeçalhos, fórmulas, formatação, agrupamentos) — é o procedimento de negócio mais complexo do form (≈580 linhas, 125–704).
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no botão `B_Ok`). Não referenciado em nenhum outro lugar do código (evidência: grep só retorna a declaração).
7. **Procedimentos chamados:** `Desligar_Tudo` (Auxiliar.bas:128), `Extrair_Info_Colunas_Fixas` (Auxiliar.bas:1081) — 5x (chaves EMPRESA/ORGANIC/VISAO/IFRS_CONTABIL/PROFORMA), `Processo_Limpar_Ajustes` (Limpeza_Base_Ajustes.bas:15), `Calcular_Comb_Meses_Intervalo` (Auxiliar.bas:378), `Form_Ref_Organic` (Aux_Formulas_Base.bas:788), `Reclassificar_Combinacoes_Empresas` (Auxiliar.bas:606), `Reclassificar_Combinacoes_IFRS_Contabil` (Auxiliar.bas:781), `Reclassificar_Combinacoes_Proforma` (Auxiliar.bas:860), `Form_Calcular_FY` (Aux_Formulas_Base.bas:22), `Ativar_Tudo` (Auxiliar.bas:69), `PopUp_Tempo_Processamento` (Auxiliar.bas:138).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Select` (fonte da lista de importação), `Me.TB_Arq`. Objetos de planilha: `WB` (ThisWorkbook), `Arq_Temp` (workbook externo), `Sh_Base`/`Sh_Destino` (worksheets de origem/destino), `fn` (`Application.WorksheetFunction`).
11. **Abas/intervalos/arquivos acessados:** Arquivo externo (`Windows(Arq)`), todas as abas escolhidas nele; internamente `Sheet13` (branch "AJUSTES", linha 490), `Sheet3` (branch "BASE", linha 574), e abas novas criadas dinamicamente para os demais layouts ("FRONT >>" como âncora de posição, linhas 200-209). Campos-chave usados como âncoras textuais: "Aux Empresa", "Menu*seleção*", "Coluna 5", "Bkp fórmula oficial", "Fórmulas:", "VERSÃO", "LIN_BASE", "Ref Cruzada".
12. **Pré-condições:** Ao menos um item em `LB_Layouts_Select` (validado); arquivo externo ainda aberto (referenciado via `Windows(Arq)`).
13. **Passos principais:** (i) monta a lista de layouts escolhidos e pede confirmação (MsgBox Sim/Não com a lista); (ii) `Desligar_Tudo` (desativa ScreenUpdating/Alerts/Events/Calculation); (iii) para cada layout selecionado, determina destino único (evita nomes de aba duplicados, incrementando sufixo numérico) e ramifica em 3 casos: **(a)** aba é um "Front de Quick Data" (identificado pela presença de cabeçalho "Menu de seleção" e "Aux Empresa") → cria/copia aba nova, remapeia cabeçalhos, fórmulas, agrupamentos de linha/coluna, formatos e fórmula padrão; **(b)** aba é "AJUSTES" → copia campos específicos para `Sheet13`, preservando colunas de exceção (`Campos_Aux`); **(c)** aba é a Base (nome igual a `Sheet3.Name`) → copia colunas do tipo `FLOAT`/`VARCHAR` para `Sheet3`, monta campo "Fonte" como `"IMPORT_QD - " & ...`, e dispara o recálculo de combinações (Organic, Empresas, IFRS Contábil, Proforma, FY); **(d)** caso não seja nenhum dos anteriores → cria aba nova copiando integralmente formatos/fórmulas/largura de coluna e todos os `Shapes` (imagens/botões) da aba de origem; (iv) fecha o arquivo externo sem salvar; (v) descarrega o form, seleciona `Sheet8`, chama `Ativar_Tudo` e `PopUp_Tempo_Processamento`.
14. **Pós-condições:** Novas abas (ou dados mesclados em `Sheet13`/`Sheet3`) presentes no workbook principal; arquivo externo fechado; usuário de volta em `Sheet8`.
15. **Efeitos colaterais / erro / mensagens / risco:** Validação: `MsgBox "Não foi selecionada nenhum Front!"` (vbCritical) se lista vazia (linha 137). Confirmação Sim/Não antes de iniciar (linha 141). **Risco alto**: nenhuma barra de progresso real, apenas `Application.StatusBar` (linha 174) — processo pode levar minutos sem feedback visual central. **Risco alto**: lógica fortemente dependente de rótulos de texto fixos no arquivo de origem (ver Seção D, risco #4) — qualquer variação de layout do arquivo externo pode gerar erro no meio do laço `For` (linha 154) sem rollback (abas já copiadas antes do erro permanecem no workbook). Há um `On Error GoTo proximaLinha_` (linha 211) que pula para a próxima iteração do laço em caso de erro pontual (linha 690, label `proximaLinha_`), mascarando falhas silenciosamente. Fórmula "padrão" é redescoberta em tempo de execução via `Cells.Find(What:="Bkp fórmula oficial"...)` (linhas 371-375) — se essa célula-âncora não existir no destino, gera erro não tratado.

### 6. `B_Pesq_Arq_Click` — Form_Importacao.frm:707
1. **Nome completo:** `B_Pesq_Arq_Click`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Abrir o diálogo nativo do Windows para o usuário escolher o arquivo a importar, e popular os campos de arquivo/diretório do form.
6. **Quem chama / gatilho:** Gatilho por clique do usuário no botão `B_Pesq_Arq`.
7. **Procedimentos chamados:** `GetArquivo` (Aux_Leitura_Nome_Arqs.bas:52).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.TB_Arq.Text`, `Me.TB_Diret.Text` (escrita).
11. **Abas/intervalos/arquivos acessados:** `Sheet8` indiretamente (via `GetArquivo`, que grava a última pasta usada em uma célula de `Sheet8`, conforme lido em `Aux_Leitura_Nome_Arqs.bas:58`).
12. **Pré-condições:** Nenhuma.
13. **Passos principais:** (i) inicializa `Diret_e_Arq = "*"`; (ii) chama `GetArquivo` (abre `Application.FileDialog` nativo); (iii) se o valor mudou (usuário não cancelou), separa manualmente o caminho completo em diretório + nome de arquivo procurando a última ocorrência de `"\"` caractere a caractere (laço `Do...Loop`); (iv) grava os dois valores em `TB_Arq`/`TB_Diret`.
14. **Pós-condições:** `TB_Arq`/`TB_Diret` preenchidos com o arquivo escolhido (ou inalterados se o usuário cancelou o diálogo).
15. **Efeitos colaterais / erro / mensagens / risco:** Parsing manual de caminho de arquivo (laço caractere a caractere) em vez de usar `InStrRev`/`Split` — funcional, porém uma escolha de implementação frágil/antiga a não replicar. Nenhuma mensagem de erro própria (delegada a `GetArquivo`).

### 7. `LB_Layouts_Geral_DblClick` — Form_Importacao.frm:730
1. **Nome completo:** `LB_Layouts_Geral_DblClick`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (DblClick de ListBox)
4. **Escopo:** Private
5. **Objetivo:** Atalho de usabilidade — duplo-clique num item da lista geral equivale a selecioná-lo e clicar em "Incluir".
6. **Quem chama / gatilho:** Gatilho por interação do usuário (duplo-clique em `LB_Layouts_Geral`).
7. **Procedimentos chamados:** `B_Incluir_Click` (linha 33).
8. **Parâmetros:** `ByVal Cancel As MSForms.ReturnBoolean` (parâmetro padrão do evento DblClick, não utilizado no corpo).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Geral` (implícito, via `B_Incluir_Click`).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Item da lista geral sob o cursor no duplo-clique.
13. **Passos principais:** Delega inteiramente a `B_Incluir_Click`.
14. **Pós-condições:** Idênticas às de `B_Incluir_Click`.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum próprio. Nota de UX: o item sob duplo-clique só é movido se estava com `Selected = True` no momento — em ListBox padrão do VBA um duplo-clique único já seleciona o item antes dos handlers disparar, então o comportamento funciona na prática, mas depende desse comportamento implícito do controle (não validado explicitamente no código).

### 8. `LB_Layouts_Select_DblClick` — Form_Importacao.frm:736
1. **Nome completo:** `LB_Layouts_Select_DblClick`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (DblClick de ListBox)
4. **Escopo:** Private
5. **Objetivo:** Atalho de usabilidade — duplo-clique num item da lista de seleção equivale a selecioná-lo e clicar em "Excluir" (remover da seleção).
6. **Quem chama / gatilho:** Gatilho por interação do usuário (duplo-clique em `LB_Layouts_Select`).
7. **Procedimentos chamados:** `B_Excluir_Click` (linha 52).
8. **Parâmetros:** `ByVal Cancel As MSForms.ReturnBoolean` (não utilizado).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Select`.
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Item sob o cursor no duplo-clique.
13. **Passos principais:** Delega a `B_Excluir_Click`.
14. **Pós-condições:** Idênticas às de `B_Excluir_Click`.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum próprio. **Ver RN-109 (Seção B)**: este form usa duplo-clique para remover, enquanto o Form_Exportacao usa clique único no evento equivalente — inconsistência de padrão de interação entre telas semelhantes.

### 9. `TB_Arq_KeyPress` — Form_Importacao.frm:742
1. **Nome completo:** `TB_Arq_KeyPress`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (KeyPress de TextBox)
4. **Escopo:** Private
5. **Objetivo:** Impedir que o usuário digite manualmente no campo de nome de arquivo, forçando o uso do botão de pesquisa (garante que o valor sempre venha de um caminho de arquivo real, validado pelo diálogo do Windows).
6. **Quem chama / gatilho:** Gatilho automático do VBA a cada tecla pressionada com o foco em `TB_Arq`.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** `ByVal KeyAscii As MSForms.ReturnInteger`.
9. **Retorno:** Nenhum (Sub); mas o parâmetro `Cancel` é setado (ver observação no campo 15 — na verdade cancela via `Cancel = True`, uma variável não declarada localmente que, por escopo do VBA, é interpretada como variável solta, não como o parâmetro de evento — ver risco).
10. **Controles/variáveis relevantes:** `Me.TB_Arq` (implícito).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Foco no campo `TB_Arq` e usuário pressiona uma tecla.
13. **Passos principais:** Exibe `MsgBox` crítico explicando que a digitação não é permitida e orientando o uso do botão de pesquisa; tenta cancelar a tecla via `Cancel = True`.
14. **Pós-condições:** Mensagem exibida a cada tecla pressionada (potencialmente repetitiva/irritante se o usuário insistir).
15. **Efeitos colaterais / erro / mensagens / risco:** **Risco médio de bug latente**: a assinatura do evento `TB_Arq_KeyPress` no MSForms padrão não expõe um parâmetro `Cancel` — o cancelamento de tecla em `KeyPress` de TextBox do VBA se faz setando `KeyAscii = 0`, não `Cancel = True`. A linha `Cancel = True` (linha 747) portanto **não cancela a tecla digitada**; cria (ou atribui a) uma variável implícita chamada `Cancel` no escopo do módulo, sem efeito sobre o evento. Evidência: linha 747, comparado à assinatura declarada na linha 742 (só `KeyAscii`, sem `Cancel` na lista de parâmetros). Isso é uma inconsistência técnica que sugere que o campo, na prática, **provavelmente permanece editável mesmo com o aviso** — merece validação funcional direta no arquivo original antes de replicar o comportamento na nova UI. Classificar como [VALIDAR COM O NEGÓCIO]/teste funcional.

### 10. `TB_Diret_KeyPress` — Form_Importacao.frm:751
1. **Nome completo:** `TB_Diret_KeyPress`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (KeyPress de TextBox)
4. **Escopo:** Private
5. **Objetivo:** Idêntico ao anterior, mas para o campo de diretório.
6. **Quem chama / gatilho:** Gatilho automático do VBA a cada tecla pressionada com foco em `TB_Diret`.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** `ByVal KeyAscii As MSForms.ReturnInteger`.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.TB_Diret` (implícito).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Foco no campo `TB_Diret`.
13. **Passos principais:** Mesmo padrão de `TB_Arq_KeyPress`: MsgBox de aviso + `Cancel = True`.
14. **Pós-condições:** Mensagem exibida.
15. **Efeitos colaterais / erro / mensagens / risco:** Mesma observação técnica do item 9 (linha 756) — `Cancel` não é parâmetro válido do evento `KeyPress`, o bloqueio de digitação pode não funcionar de fato como pretendido. [VALIDAR COM O NEGÓCIO].

### 11. `UserForm_Activate` — Form_Importacao.frm:760
1. **Nome completo:** `UserForm_Activate`
2. **Módulo:** Form_Importacao.frm
3. **Tipo:** Evento (Activate de UserForm)
4. **Escopo:** Private
5. **Objetivo:** Inicializar o estado do form sempre que ele é exibido/ativado, limpando as duas listas.
6. **Quem chama / gatilho:** Disparado automaticamente pelo VBA quando `Form_Importacao.Show` é executado (chamado por `Importar_Fronts`, Auxiliar.bas:56-59) — gatilho por evento de ciclo de vida do form, não por interação direta do usuário, mas consequência indireta do clique que invoca `Importar_Fronts`.
7. **Procedimentos chamados:** Nenhum ativo. Bloco inteiro de auto-preenchimento para um usuário específico (`Environ("UserName") = "F8044606"`) está **comentado** (linhas 765-769) — antigo atalho de desenvolvedor/teste, hoje inativo.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Layouts_Geral`, `Me.LB_Layouts_Select` (ambos limpos).
11. **Abas/intervalos/arquivos acessados:** Nenhum (o bloco comentado referenciaria um caminho de rede fixo `F:\P&C 2021\Reports\Front Ends\...` — evidência de ambiente de desenvolvimento/uso específico de um analista).
12. **Pré-condições:** Form sendo exibido.
13. **Passos principais:** `Clear` em ambas as ListBox.
14. **Pós-condições:** Form pronto para uso, sem arquivo/layout pré-selecionado (no comportamento ativo atual).
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum no código ativo. Risco informativo baixo: o código comentado revela um caminho de rede e usuário (`F8044606`) hardcoded historicamente usado para acelerar testes — não é executado hoje, mas é evidência de dependência de ambiente de rede específico em versões passadas do processo (`F:\P&C 2021\...`).

---

## A.2 — Form_Exportacao.frm (11 procedimentos)

### 1. `B_CANCELAR_Click` — Form_Exportacao.frm:10
1. **Nome completo:** `B_CANCELAR_Click`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Fechar o form de exportação sem executar nenhuma ação.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no botão `B_CANCELAR`).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum além do próprio form.
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Nenhuma.
13. **Passos principais:** `Unload Me`.
14. **Pós-condições:** Form fechado; nenhum estado de ambiente (ScreenUpdating/Events) foi alterado, pois nenhuma etapa de exportação chegou a rodar `Desligar_Tudo`.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhuma confirmação é pedida (diferente do `B_Cancel_Click` do Form_Importacao, que pergunta se deve fechar o arquivo externo) — aqui é aceitável pois este cancelar não interage com arquivos externos abertos. Comportamento simples e seguro.

### 2. `B_Incluir_Click` — Form_Exportacao.frm:14
1. **Nome completo:** `B_Incluir_Click`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Mover os itens (abas "Front" ou combinações "KPI > Versão > Ano") marcados na lista geral para a lista de seleção, evitando duplicados.
6. **Quem chama / gatilho:** Gatilho por clique do usuário no botão `B_Incluir`. Também chamado por `LB_Versoes_Geral_DblClick` (linha 475).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Versoes_Geral` (ListBox origem, multi-seleção), `Me.LB_Versoes_Select` (ListBox destino).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** `LB_Versoes_Geral` populado (via `Carregar_ListBox`).
13. **Passos principais:** Mesma lógica de `B_Incluir_Click` do Form_Importacao: para cada item marcado, verifica duplicidade (case-insensitive) e adiciona se novo; desmarca o item de origem. Adicionalmente redimensiona `LB_Versoes_Geral.Height`/`LB_Versoes_Select.Height` para `118` (linhas 28-29 — **ativo aqui**, diferente do Form_Importacao onde o equivalente está comentado).
14. **Pós-condições:** `LB_Versoes_Select` atualizado sem duplicados.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhuma mensagem. Nota de inconsistência: este form ativa o redimensionamento de altura das listas via código (linhas 28-29), enquanto o Form_Importacao tem a mesma lógica comentada — sugere manutenção divergente entre os dois formulários apesar de serem estruturalmente semelhantes.

### 3. `B_Excluir_Click` — Form_Exportacao.frm:33
1. **Nome completo:** `B_Excluir_Click`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Remover da lista de seleção os itens marcados.
6. **Quem chama / gatilho:** Gatilho por clique do usuário no botão `B_Excluir`. Também chamado por `LB_Versoes_Select_Click` (linha 481) — **aqui o gatilho equivalente é clique único, não duplo-clique** (ver campo 15 e RN-109).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Versoes_Select`.
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Item marcado em `LB_Versoes_Select`.
13. **Passos principais:** Mesmo padrão `GoTo Inicio` de remoção iterativa do Form_Importacao; redimensiona alturas para 118 ao final (linhas 45-46).
14. **Pós-condições:** Itens marcados removidos da seleção.
15. **Efeitos colaterais / erro / mensagens / risco:** **RN-109**: `LB_Versoes_Select_Click` (evento de **clique simples**, linha 479) chama este procedimento diretamente — ou seja, no Form_Exportacao, um único clique num item já selecionado da lista de seleção **o remove imediatamente**, sem confirmação e sem exigir duplo-clique como no Form_Importacao. Risco de remoção acidental (usuário clica só para inspecionar/re-selecionar o item e ele desaparece da lista).

### 4. `B_Ok_Click` — Form_Exportacao.frm:50
1. **Nome completo:** `B_Ok_Click`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Click de CommandButton)
4. **Escopo:** Private
5. **Objetivo:** Executar a exportação efetiva: gera arquivo(s) `.xlsb` com os "Fronts" (abas de relatório) ou com a "Base" (dados consolidados filtrados), conforme a página ativa do `Page_Frame` e o modo (único/separado) escolhido.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no botão `B_Ok`).
7. **Procedimentos chamados:** `GetPasta` (Aux_Leitura_Nome_Arqs.bas:103), `Desligar_Tudo` (Auxiliar.bas:128), `Excluir_Campos_Selecao_Front` (Form_Exportacao.frm:430 — 2x, modo Front único e separado), `Excluir_Item_Especifico` (Auxiliar.bas:1036 — 2x, remoção de EBITDA e "IFRS Itália"), `Ativar_Tudo` (Auxiliar.bas:69).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.OP_Arq_Unico`, `Me.OP_Arq_Sep` (OptionButton, par exclusivo — evidência: `.Value = 0`/`.Value = True` tratados como par booleano nas linhas 57 e 106/153/340/349), `Me.LB_Versoes_Select`, `Me.Page_Frame.SelectedItem.Index` (0 = Front, 1 = Base).
11. **Abas/intervalos/arquivos acessados:** `Sheet3` (Base, alias `Sh_Base`), abas do `ThisWorkbook` correspondentes aos nomes escolhidos (modo Front), pasta de destino escolhida pelo usuário via `GetPasta`, workbooks temporários criados via `Workbooks.Add`.
12. **Pré-condições:** Tipo de exportação (único/separado) escolhido; ao menos uma versão/front selecionado; pasta de destino selecionável.
13. **Passos principais:** (i) valida tipo de exportação escolhido (linha 57-61); (ii) monta string de versões escolhidas e valida não-vazio (linhas 65-74); (iii) monta texto de confirmação conforme página ativa (Front/Base) e pede confirmação Sim/Não (linha 82); (iv) pede pasta de destino via `GetPasta`, valida não-vazio (linhas 90-97); (v) `Desligar_Tudo`; (vi) **branch Front** (`Page_Frame.Index = 0`): se "Arquivo único", copia todas as abas escolhidas para um novo workbook, converte fórmulas em valores, remove campos de seleção de menu (`Excluir_Campos_Selecao_Front`), remove abas não escolhidas, salva como `FRONT_QD - <timestamp>.xlsb`; se "Arquivos separados", repete o processo aba a aba, um arquivo por aba (`FRONT_QD - <aba> - <timestamp>.xlsb`); (vii) **branch Base** (`Page_Frame.Index = 1`): copia a base consolidada para workbook temporário, remove linhas de EBITDA e de "IFRS Itália", ordena por KPI/Versão/Ano, remove linhas fora das combinações escolhidas, remove colunas fora da lista `Manter_Colunas`, converte colunas de meses em valores; se "único" salva tudo em `BASE_QD - <timestamp>.xlsb`; se "separado", usa `AutoFilter` por KPI/Versão/Ano e gera um arquivo por combinação (`BASE_QD - <KPI_Versão_Ano> - <timestamp>.xlsb`); (viii) `Ativar_Tudo`, `MsgBox` de sucesso, `Unload Me`.
14. **Pós-condições:** Um ou mais arquivos `.xlsb` gravados na pasta escolhida; workbooks temporários fechados sem salvar as cópias intermediárias no workbook principal; form descarregado.
15. **Efeitos colaterais / erro / mensagens / risco:** Validações com `MsgBox vbCritical`: tipo de exportação não escolhido (linha 58), nenhuma versão selecionada (linha 71), nenhuma pasta selecionada (linha 94) — todas com a frase "O processo será cancelado automaticamente!". **Risco crítico (RN-107)**: `If Dir(Diret & Nome_Arq) <> "" Then Kill (Diret & Nome_Arq)` (linhas 147, 175, 401) apaga **sem confirmação** um arquivo pré-existente de mesmo nome antes de salvar o novo — nome inclui timestamp até o minuto (`YYYYMMDD_HHMM`), portanto duas exportações no mesmo minuto colidem e a primeira é sobrescrita silenciosamente. Mensagem final única e genérica: "Processo concluído com sucesso!" (linha 423), sem resumo de quantos arquivos/onde foram salvos (RN-115). Nenhuma barra de progresso durante o processamento (pode envolver múltiplos `Workbooks.Add`/`Sort`/`AutoFilter`/`SaveAs` em sequência).

### 5. `Excluir_Campos_Selecao_Front` — Form_Exportacao.frm:430
1. **Nome completo:** `Excluir_Campos_Selecao_Front`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Sub (procedimento auxiliar, não é evento)
4. **Escopo:** Private
5. **Objetivo:** Remover, de uma cópia exportada de um "Front", as linhas/colunas de controle de menu de seleção que só fazem sentido dentro do workbook original (ex.: o menu superior de filtros), deixando o arquivo exportado "limpo" para o destinatário final.
6. **Quem chama / gatilho:** Chamado por `B_Ok_Click` — 2x (linha 129, modo arquivo único; linha 171, modo arquivos separados).
7. **Procedimentos chamados:** Nenhum (usa apenas `Application.WorksheetFunction`).
8. **Parâmetros:** `Sh_Temp` (worksheet a ser limpa, passada por referência implícita — VBA `ByRef` por padrão).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum controle de form; variáveis locais `Lin_Cabecalho_Base`, `Col_Cabecalho_Base`.
11. **Abas/intervalos/arquivos acessados:** `Sh_Temp` (worksheet temporária recém-copiada). Âncoras textuais: "Menu*seleção*" (linha 440), "Aux Empresa" (linha 441), "Coluna 5*" (linha 451).
12. **Pré-condições:** `Sh_Temp` deve ser uma cópia de um "Front" no padrão Quick Data (com cabeçalho "Menu de seleção").
13. **Passos principais:** (i) verifica se a aba tem o cabeçalho de menu de seleção; se sim, (ii) localiza a linha/coluna de âncora "Aux Empresa"/"Menu de seleção"; (iii) calcula o intervalo de linhas de controle (usa um contador `Z` que para após 2 linhas vazias consecutivas a partir de "Coluna 5", linhas 454-458); (iv) apaga (`EntireRow.Delete`/`EntireColumn.Delete`) as linhas e colunas de controle.
14. **Pós-condições:** `Sh_Temp` sem as linhas/colunas de menu de seleção — pronta para exportação como arquivo final.
15. **Efeitos colaterais / erro / mensagens / risco:** Se a aba não tiver o cabeçalho esperado, a sub simplesmente não faz nada (sem erro, sem aviso) — comportamento correto para abas que não são "Fronts" de Quick Data. Lógica de parada do laço (2 linhas vazias consecutivas) é frágil a variações de layout — mesma classe de risco do `B_Ok_Click` do Form_Importacao.

### 6. `ListBox1_DblClick` — Form_Exportacao.frm:467
1. **Nome completo:** `ListBox1_DblClick`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (DblClick de ListBox)
4. **Escopo:** Private
5. **Objetivo:** [NÃO IDENTIFICADO COM CERTEZA] — pelo nome do procedimento chamado (`Registrar_Sheet`), aparenta ser uma funcionalidade de "registrar"/marcar uma aba a partir de uma lista (`ListBox1`), mas o procedimento-alvo não existe em nenhum arquivo do dump.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (duplo-clique em `ListBox1`) — **porém não há evidência de que `ListBox1` seja populado ou mesmo visível em nenhum outro procedimento deste form** (não aparece em `Carregar_ListBox`, `UserForm_Activate`, nem em nenhum outro handler). Pode ser controle órfão/legado ainda presente no design visual mas não mais alimentado por código.
7. **Procedimentos chamados:** `Registrar_Sheet` — **procedimento não encontrado em nenhum arquivo do dump** (evidência: `grep -rn "Registrar_Sheet" vba_dump_tmp/` retorna apenas esta linha de chamada, Form_Exportacao.frm:469; nenhuma declaração `Sub Registrar_Sheet` ou `Function Registrar_Sheet` em lugar nenhum). Classificado como referência morta/quebrada — ver RN-114 e Risco Crítico #1 (Seção D).
8. **Parâmetros:** `ByVal Cancel As MSForms.ReturnBoolean` (não utilizado).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.ListBox1` — controle cuja função e conteúdo não puderam ser determinados a partir do código disponível. [NÃO IDENTIFICADO].
11. **Abas/intervalos/arquivos acessados:** [NÃO IDENTIFICADO] (depende do que `Registrar_Sheet` faria, mas esse procedimento não existe no código-fonte disponível).
12. **Pré-condições:** `Me.ListBox1.Value <> ""` (linha 469) — ou seja, algum item precisa estar selecionado/valor presente.
13. **Passos principais:** Se `ListBox1.Value` não vazio, chama `Registrar_Sheet`.
14. **Pós-condições:** [NÃO IDENTIFICADO] — dependeria da execução de `Registrar_Sheet`, que geraria erro de compilação/execução ("Sub ou Function não definida") caso este código chegasse a rodar em um ambiente VBA que compile sob demanda, ou erro de compilação antecipado se o VBA compilar o projeto inteiro antes de qualquer execução (mais provável, dado que VBA compila o módulo inteiro ao rodar qualquer macro do projeto).
15. **Efeitos colaterais / erro / mensagens / risco:** **RISCO CRÍTICO**: referência a procedimento inexistente. Em VBA clássico, isso normalmente causa erro de compilação ("Sub ou Function não definida") **assim que qualquer macro do projeto é executada** (o compilador VBA valida o projeto inteiro antes de rodar), o que sugere fortemente que **este código nunca é de fato executado na prática** — ou porque `ListBox1` está oculto/não utilizado na versão atual do form, ou porque o arquivo `.xlsb` de produção já não compila mais este trecho (possível resíduo de uma refatoração incompleta em que o `ListBox1` foi descontinuado mas o handler não foi removido). Recomenda-se checar no arquivo original `.xlsb` (fora do escopo deste dump de texto) se `ListBox1` sequer existe mais no formulário renderizado, e se o projeto VBA compila sem erro — teste sugerido na Seção D.

### 7. `LB_Versoes_Geral_DblClick` — Form_Exportacao.frm:473
1. **Nome completo:** `LB_Versoes_Geral_DblClick`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (DblClick de ListBox)
4. **Escopo:** Private
5. **Objetivo:** Atalho — duplo-clique num item da lista geral equivale a incluí-lo na seleção.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (duplo-clique em `LB_Versoes_Geral`).
7. **Procedimentos chamados:** `B_Incluir_Click` (linha 14).
8. **Parâmetros:** `ByVal Cancel As MSForms.ReturnBoolean` (não utilizado).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Versoes_Geral` (implícito).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Item sob o cursor.
13. **Passos principais:** Delega a `B_Incluir_Click`.
14. **Pós-condições:** Idênticas às de `B_Incluir_Click`.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum próprio. Simetria correta com o comportamento do Form_Importacao para a operação de "incluir" (ambos usam duplo-clique aqui) — a divergência de padrão está apenas no lado "excluir" (ver item 3 e RN-109).

### 8. `LB_Versoes_Select_Click` — Form_Exportacao.frm:479
1. **Nome completo:** `LB_Versoes_Select_Click`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Click, não DblClick, de ListBox)
4. **Escopo:** Private
5. **Objetivo:** Remover imediatamente da seleção o item clicado (ver risco de UX).
6. **Quem chama / gatilho:** Gatilho por interação do usuário — **clique único** (não duplo) em qualquer item de `LB_Versoes_Select`.
7. **Procedimentos chamados:** `B_Excluir_Click` (linha 33).
8. **Parâmetros:** Nenhum (evento `Click` padrão de ListBox não recebe parâmetros).
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Versoes_Select` (implícito).
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** Qualquer clique que resulte em item marcado (`Selected = True`) em `LB_Versoes_Select`.
13. **Passos principais:** Delega a `B_Excluir_Click`, que remove todo item atualmente marcado.
14. **Pós-condições:** Item(ns) marcado(s) removido(s) da lista de seleção.
15. **Efeitos colaterais / erro / mensagens / risco:** **RN-109 / Risco médio**: como o evento é `Click` (não `DblClick`), **qualquer clique simples num item da lista de seleção já o remove** — inconsistente com o padrão duplo-clique usado no Form_Importacao para a mesma finalidade, e mais agressivo do ponto de vista de UX (maior chance de remoção acidental).

### 9. `Page_Frame_Change` — Form_Exportacao.frm:485
1. **Nome completo:** `Page_Frame_Change`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Change de MultiPage)
4. **Escopo:** Private
5. **Objetivo:** Recarregar as listas de itens disponíveis sempre que o usuário alterna entre a página "Exportar Front" e a página "Exportar Base".
6. **Quem chama / gatilho:** Gatilho automático do VBA ao trocar de página no `Page_Frame` (interação do usuário clicando na aba da página).
7. **Procedimentos chamados:** `Carregar_ListBox` (Form_Exportacao.frm:502).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.Page_Frame` (MultiPage, 2 páginas — evidência: `.SelectedItem.Index` comparado a 0 e 1 em `B_Ok_Click`, linhas 79-80 e 104/185).
11. **Abas/intervalos/arquivos acessados:** Indireto, via `Carregar_ListBox`.
12. **Pré-condições:** Form aberto.
13. **Passos principais:** Chama `Carregar_ListBox`.
14. **Pós-condições:** Listas atualizadas conforme a página ativa.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum próprio.

### 10. `UserForm_Activate` — Form_Exportacao.frm:492
1. **Nome completo:** `UserForm_Activate`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Evento (Activate de UserForm)
4. **Escopo:** Private
5. **Objetivo:** Inicializar o form sempre na página "Exportar Front" e carregar a lista correspondente.
6. **Quem chama / gatilho:** Disparado automaticamente pelo VBA quando `Form_Exportacao.Show` é executado (chamado por `Exportar`, Auxiliar.bas:50-52).
7. **Procedimentos chamados:** `Carregar_ListBox` (linha 497).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.Page_Frame.SelectedItem.Index` (forçado para `0`, linha 495).
11. **Abas/intervalos/arquivos acessados:** `ThisWorkbook` (ativado, linha 494); indireto via `Carregar_ListBox`.
12. **Pré-condições:** Form sendo exibido.
13. **Passos principais:** Ativa `ThisWorkbook`; força a página 0 ("Exportar Front") como ativa; chama `Carregar_ListBox`.
14. **Pós-condições:** Form exibido sempre iniciando no modo "Exportar Front", com a lista geral já populada.
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhum. Regra de negócio implícita: o modo padrão de exportação é sempre "Front", nunca "Base" — usuário precisa clicar manualmente para trocar de modo a cada abertura do form (não há memória da última escolha).

### 11. `Carregar_ListBox` — Form_Exportacao.frm:502
1. **Nome completo:** `Carregar_ListBox`
2. **Módulo:** Form_Exportacao.frm
3. **Tipo:** Sub (procedimento auxiliar, não é evento)
4. **Escopo:** Private
5. **Objetivo:** Popular a `LB_Versoes_Geral` com a lista de itens disponíveis, de acordo com a página ativa (nomes de abas "Front" ou combinações únicas "KPI > Versão > Ano" da Base), e ajustar a cor de fundo do form conforme o modo.
6. **Quem chama / gatilho:** Chamado por `Page_Frame_Change` (linha 487) e por `UserForm_Activate` (linha 497).
7. **Procedimentos chamados:** `Carregar_Sheets_Suporte` (Auxiliar.bas:582, apenas no modo Front), `Ordenar_Lista` (Auxiliar.bas:1096).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Me.LB_Versoes_Geral`, `Me.LB_Versoes_Select` (ambos limpos no início), `Me.BackColor`, `Me.Page_Frame.BackColor`, `Me.Label1/2/3.BackColor`, `Me.Frame1.BackColor`, `Me.OP_Arq_Sep.BackColor`, `Me.OP_Arq_Unico.BackColor` (todos recoloridos conforme o modo — evidência de uso de cor como sinalizador visual do modo ativo, `&HE0E0E0` para Front, `&HF1E6DC` para Base).
11. **Abas/intervalos/arquivos acessados:** Modo Front: todas as `Worksheets` de `ThisWorkbook`, exceto as de suporte (`Sheets_Suporte`). Modo Base: `Sheet3`, faixa a partir da linha "LIN_BASE", colunas "KPI", "VERSÃO", "EXERCICIO".
12. **Pré-condições:** `Me.Page_Frame.SelectedItem.Index` definido (0 ou 1).
13. **Passos principais:** (i) limpa as duas ListBox; (ii) se página 0 (Front): carrega sheets de suporte e monta lista de nomes de abas não-suporte; (iii) se página 1 (Base): localiza cabeçalho "LIN_BASE" em `Sheet3`, monta lista de combinações únicas "KPI > Versão > Ano" (evita duplicados via string `Versao_Cadastrada`); (iv) ordena a lista (`Ordenar_Lista`) e popula `LB_Versoes_Geral`; (v) aplica a cor de fundo correspondente ao modo em todos os controles estáticos do form.
14. **Pós-condições:** `LB_Versoes_Geral` populada e ordenada; aparência do form reflete visualmente o modo ativo (cor de fundo).
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhuma mensagem de erro. Não há tratamento explícito para o caso de `Sheet3` estar vazia/sem dados (o `fn.Match` pode falhar silenciosamente para 0 itens — [NÃO ACESSÍVEL] confirmar comportamento exato sem executar).

---

## A.3 — Form_Tratamento_Opcoes.frm (3 procedimentos)

### 1. `UserForm_Initialize` — Form_Tratamento_Opcoes.frm:9
1. **Nome completo:** `UserForm_Initialize`
2. **Módulo:** Form_Tratamento_Opcoes.frm
3. **Tipo:** Evento (Initialize de UserForm)
4. **Escopo:** Private
5. **Objetivo:** Montar e exibir uma grade somente-leitura com o cabeçalho das opções de extração configuradas em `Sheet8` (faixa `B94:S112`) — aparenta ser uma tela de consulta/depuração das fontes de dados habilitadas para extração.
6. **Quem chama / gatilho:** Disparado automaticamente pelo VBA na primeira referência ao form (ex.: `Form_Tratamento_Opcoes.Show` ou qualquer acesso a `Form_Tratamento_Opcoes`). **Nenhum ponto do dump de código chama `Form_Tratamento_Opcoes.Show` ou referencia `Form_Tratamento_Opcoes` externamente** (evidência: `grep -rn "Form_Tratamento_Opcoes" vba_dump_tmp/` retorna apenas ocorrências dentro do próprio arquivo — linhas 1, 35-38). **O ponto de entrada externo deste form é [NÃO IDENTIFICADO]** no código disponível — pode estar associado a um botão/shape na planilha cujo `OnAction` não está capturado neste dump (a atribuição de macro a um Shape fica armazenada fora do texto VBA, na definição do objeto de planilha).
7. **Procedimentos chamados:** Nenhum procedimento externo — o próprio `Form_Tratamento_Opcoes.Show` é chamado de dentro deste evento (linha 38), o que é incomum (form se auto-exibe dentro do seu próprio `Initialize`, antes mesmo de terminar de montar todo o conteúdo — ver campo 15).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `ListBox1` (multi-coluna, `ColumnCount = 10`, `ColumnWidths = "40;80;80;80;80;80;80;120;120;120"` — evidência direta, linhas 27-28). Variáveis: `rngExtracoes` (Range `Sh8.Range("B94:S112")`), `cabecalhos` (string concatenada), `colunas_cabecalho` (array via `Split`), `dados` (string concatenada das linhas com "Sim" na primeira coluna).
11. **Abas/intervalos/arquivos acessados:** `Sheet8`, intervalo fixo `B94:S112`.
12. **Pré-condições:** `Sheet8!B94:S112` deve conter, na primeira linha do range, os cabeçalhos das colunas de configuração, e nas linhas seguintes, "Sim"/outro valor na primeira coluna indicando se a extração está habilitada.
13. **Passos principais:** (i) monta a string de cabeçalhos pulando as colunas de índice 2, 3 e acima de 11 (`qtdeColunas`); (ii) separa em array via `Split(cabecalhos, ";")` e popula a linha 0 do `ListBox1` coluna a coluna; (iii) redimensiona o form (`Width=600`, `Height=200`) e o `ListBox1` (`Width=564`); (iv) **exibe o form** (`Form_Tratamento_Opcoes.Show`, linha 38) — **antes** de terminar o restante da rotina; (v) *após* o `Show` (código que só roda quando o `Show` for não-modal, ou após o usuário fechar o form, se for modal — comportamento depende da modalidade do form, [NÃO ACESSÍVEL] sem o `.frx`), percorre novamente `rngExtracoes` e monta a string `dados` só com as linhas cuja primeira coluna é "Sim"; (vi) `Debug.Print dados`.
14. **Pós-condições:** Form exibido com o cabeçalho das colunas de configuração na primeira linha do `ListBox1`. A lista de itens habilitados ("Sim") é calculada em `dados` mas **nunca é escrita de volta no `ListBox1` nem em nenhum outro controle visível** — só vai para a janela Immediate (`Debug.Print`), inacessível ao usuário final.
15. **Efeitos colaterais / erro / mensagens / risco:** **RN-110 / Risco médio**: a etapa mais "útil" do form do ponto de vista de negócio (mostrar quais extrações estão habilitadas) nunca chega à tela — indício de funcionalidade incompleta ou de uma reformulação que ficou pela metade. **Chamar `.Show` no meio do próprio `Initialize`** é uma prática atípica em VBA (o `Initialize` normalmente prepara o form e quem chama de fora decide quando mostrar) — sugere que este form pode ter sido adaptado/colado de outro contexto sem refatoração completa. Dimensões do form e da lista são hardcoded (600x200 / largura 564) — não responsivas.

### 2. `get_ListaDeExtracoesSelecionadas_cabecalho` — Form_Tratamento_Opcoes.frm:59
1. **Nome completo:** `get_ListaDeExtracoesSelecionadas_cabecalho`
2. **Módulo:** Form_Tratamento_Opcoes.frm
3. **Tipo:** Function (mas sem instrução `<nome> = ...` explícita atribuindo valor de retorno — ver campo 9)
4. **Escopo:** Implícito Public (não declarado `Private`, diferente dos demais procedimentos do form — evidência: ausência da palavra `Private` na linha 59, ao contrário de `UserForm_Initialize` que é `Private Sub`).
5. **Objetivo:** Aparentemente recalcular/expor o cabeçalho das colunas de extração (mesma lógica de montagem de `cabecalhos` do `UserForm_Initialize`, com pequena variação nos índices de coluna pulados) e a lista de linhas habilitadas ("Sim") — provavelmente pensada para ser chamada externamente (dado o escopo público e o prefixo `get_`), mas não usada em lugar nenhum do código disponível.
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump** (evidência: `grep -rn "get_ListaDeExtracoesSelecionadas_cabecalho" vba_dump_tmp/` retorna apenas a própria declaração). Código morto ou função utilitária pensada para uso futuro/manual (via janela Immediate) nunca integrada à interface.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Declarada como `Function`, mas o corpo **nunca atribui um valor ao nome da função** (não há linha `get_ListaDeExtracoesSelecionadas_cabecalho = ...`) — portanto retorna sempre `Empty`/`Null` padrão do VBA para uma Function sem atribuição. Efetivamente se comporta como um `Sub` disfarçado de `Function`.
10. **Controles/variáveis relevantes:** Nenhum controle de form. Variáveis locais `Sh8`, `rngExtracoes`, `cabecalhos`, `dados`, `rngExtracoes_Filtrado` (`Set` no final, também não retornado/exposto).
11. **Abas/intervalos/arquivos acessados:** `Sheet8`, intervalo `B94:S112`.
12. **Pré-condições:** Mesmas do item anterior.
13. **Passos principais:** Monta `cabecalhos` (pulando colunas 2 e 3), depois monta `dados` (linhas com "Sim" na primeira coluna, pulando colunas 2 e 3), ambos só para `Debug.Print`.
14. **Pós-condições:** Nenhuma alteração persistente; apenas saída no Immediate Window.
15. **Efeitos colaterais / erro / mensagens / risco:** Risco baixo (função não usada, sem efeito colateral em dados de negócio), porém é um indício adicional de retrabalho/duplicação de lógica dentro do mesmo form (a mesma extração de cabeçalho é feita 2x, com pequenas diferenças de índice de coluna pulada entre este método e o `UserForm_Initialize` — linha 18 do Initialize pula colunas `2, 3 ou >11`; aqui, linha 67, pula apenas `2 ou 3`). Isso é uma **divergência funcional sutil** entre as duas implementações que pode gerar confusão sobre qual é a lógica "correta" caso alguém reative este código no futuro.

### 3. `get_ListaDeExtracoesSelecionadas_corpo` — Form_Tratamento_Opcoes.frm:94
1. **Nome completo:** `get_ListaDeExtracoesSelecionadas_corpo`
2. **Módulo:** Form_Tratamento_Opcoes.frm
3. **Tipo:** Function (mesmo padrão do item anterior — sem atribuição de retorno)
4. **Escopo:** Implícito Public.
5. **Objetivo:** Extrair apenas o "corpo" (linhas com "Sim") das opções de extração, sem montar o cabeçalho — provavelmente pensada para ser combinada com `get_ListaDeExtracoesSelecionadas_cabecalho` em algum consumidor externo não implementado.
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump** (evidência: grep retorna apenas a declaração, Form_Tratamento_Opcoes.frm:94). Código morto/não integrado.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Declarada como `Function`, mas sem atribuição ao nome — retorna vazio por padrão. Mesmo padrão de "Function disfarçada de Sub" do item anterior.
10. **Controles/variáveis relevantes:** Nenhum controle de form. Variáveis locais `Sh8`, `rngExtracoes`, `dados`, `rngExtracoes_Filtrado`.
11. **Abas/intervalos/arquivos acessados:** `Sheet8`, intervalo `B94:S112`.
12. **Pré-condições:** Mesmas dos itens anteriores.
13. **Passos principais:** Percorre `rngExtracoes`, monta `dados` com as linhas cuja primeira coluna é "Sim" (pulando colunas 2 e 3), `Debug.Print dados`.
14. **Pós-condições:** Nenhuma persistente.
15. **Efeitos colaterais / erro / mensagens / risco:** Risco baixo, mesma observação de código não utilizado/duplicado do item anterior. Em conjunto, os 3 procedimentos deste form sugerem uma funcionalidade de "relatório de configuração de extrações" que foi iniciada mas nunca finalizada/exposta ao usuário — recomenda-se **[VALIDAR COM O NEGÓCIO]** se essa visualização ainda é um requisito desejado para a nova interface.

---

## A.4 — Sheet8.cls (2 procedimentos)

### 1. `ComboBox1_Change` — Sheet8.cls:10
1. **Nome completo:** `ComboBox1_Change`
2. **Módulo:** Sheet8.cls (planilha "home"/painel principal)
3. **Tipo:** Evento (Change de ComboBox embutido na planilha)
4. **Escopo:** Private
5. **Objetivo:** Funcionar como um "menu de comandos rápidos" embutido na planilha: ao selecionar um item do combo, dispara a rotina de manutenção correspondente (limpar nomes definidos ou atualizar combinações de empresas).
6. **Quem chama / gatilho:** Gatilho por interação do usuário (seleção de um item no `ComboBox1` da planilha `Sheet8`).
7. **Procedimentos chamados:** `CLEAR_Defined_Names` (TK_Functions.bas:342) — se o texto contém "Defined Names"; `UPDATE_Combinacoes_Empresas` (TK_Functions.bas:115) — se o texto contém "Combinações".
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `ComboBox1` (ActiveX embutido na planilha, não um UserForm — evidência: é membro do `Sheet8.cls`, módulo de código de planilha, e não de um `.frm`).
11. **Abas/intervalos/arquivos acessados:** Indireto, via as rotinas chamadas (`CLEAR_Defined_Names` opera sobre `Names` do workbook; `UPDATE_Combinacoes_Empresas` opera sobre `Sheet3`/base — [NÃO APROFUNDADO], fora do escopo deste cluster).
12. **Pré-condições:** `ComboBox1` deve estar populado com os itens de comando (a rotina de população, `Worksheet_Activate`, está **comentada/desativada** — linhas 20-33 do Sheet8.cls — portanto o preenchimento atual do combo é [NÃO IDENTIFICADO]/pode depender de configuração manual do controle no design da planilha).
13. **Passos principais:** Verifica via `InStr` (busca de substring, case-sensitive quanto ao texto literal buscado) se o valor do combo contém "Defined Names" ou "Combinações"; chama a rotina correspondente.
14. **Pós-condições:** Dependem da rotina chamada (fora do escopo detalhado deste cluster — pertence a `TK_Functions.bas`).
15. **Efeitos colaterais / erro / mensagens / risco:** **RN-111 / Risco médio**: matching por substring de texto (`InStr`) em vez de valor/ID fixo — se o texto do item do combo for renomeado (ex.: erro de digitação ao editar a lista de itens do combo, hoje feita manualmente já que a população automática está comentada), o gatilho correspondente deixa de funcionar silenciosamente (nenhum `Else`/mensagem de "comando não reconhecido"). O bloco `Worksheet_Activate` comentado (que populava o combo com "Escolha um comando...", "LIMPAR | 'Defined Names'", "ATUALIZAR | Combinações 'Base + DropComb'") sugere que a população automática foi desativada propositalmente, deixando a manutenção da lista de opções do combo dependente de configuração manual do objeto na planilha — **[NÃO ACESSÍVEL]** confirmar o estado atual dos itens sem o arquivo binário.

### 2. `Worksheet_BeforeDoubleClick` — Sheet8.cls:35
1. **Nome completo:** `Worksheet_BeforeDoubleClick`
2. **Módulo:** Sheet8.cls
3. **Tipo:** Evento (BeforeDoubleClick de Worksheet)
4. **Escopo:** Private
5. **Objetivo:** Atalho de produtividade: permitir que o usuário abra diretamente um arquivo (ou a pasta que o contém) referenciado em células da planilha, com um simples duplo-clique, sem precisar copiar/colar o caminho em outro lugar.
6. **Quem chama / gatilho:** Gatilho automático do Excel a cada duplo-clique em qualquer célula de `Sheet8` (evento nativo de Worksheet — interação do usuário).
7. **Procedimentos chamados:** Nenhum procedimento VBA customizado; usa diretamente `Workbooks.Open` e `Shell "C:\WINDOWS\explorer.exe ..."` (funções nativas do VBA/Windows).
8. **Parâmetros:** `ByVal Target As Range` (célula duplo-clicada), `Cancel As Boolean` (usado para suprimir o comportamento padrão de edição de célula quando a ação de abrir arquivo/pasta é disparada).
9. **Retorno:** Nenhum (Sub); efeito via `Cancel` (este sim é um parâmetro válido de `BeforeDoubleClick`, ao contrário do problema identificado em `TB_Arq_KeyPress`/`TB_Diret_KeyPress`).
10. **Controles/variáveis relevantes:** Nenhum controle de UserForm — opera sobre células da própria planilha. Variáveis locais `celula`, `Diret`, `Arq`.
11. **Abas/intervalos/arquivos acessados:** A célula duplo-clicada e a célula imediatamente à esquerda (`Offset(0, -1)`), usada para identificar se o rótulo é "Diretório:" ou "Arquivo:"; se "Arquivo:", também lê a célula acima (`Offset(-1, 0)`) para obter o diretório correspondente. Abre o arquivo/pasta apontado.
12. **Pré-condições:** A célula à esquerda da célula duplo-clicada deve conter literalmente o texto "Diretório:" ou "Arquivo:" (comparação `UCase`, exata, não `InStr`).
13. **Passos principais:** (i) identifica se a célula-rótulo à esquerda é "Diretório:" ou "Arquivo:"; (ii) monta `Diret`/`Arq` conforme o caso; (iii) se ambos preenchidos, abre o arquivo (`Workbooks.Open`, `ReadOnly:=True`, com `DisplayAlerts` temporariamente desligado) e cancela o comportamento padrão do duplo-clique (`Cancel = True`); (iv) se só `Diret` preenchido, abre o Windows Explorer na pasta (`Shell`) e cancela o comportamento padrão.
14. **Pós-condições:** Arquivo aberto em modo leitura no Excel, ou janela do Explorer aberta na pasta referenciada; célula não entra em modo de edição (graças ao `Cancel = True`).
15. **Efeitos colaterais / erro / mensagens / risco:** Nenhuma validação de existência do arquivo/pasta antes de tentar abrir — se o caminho estiver desatualizado ou o arquivo tiver sido movido/renomeado, o Excel exibirá seu próprio erro nativo (não tratado/capturado pelo código). **RN-112**: atalho de UX não documentado na interface (nenhuma dica visual de que a célula é clicável) — usuário só descobre por tentativa ou treinamento. Path do Explorer hardcoded como `"C:\WINDOWS\explorer.exe"` (linha 57) — assume localização padrão do Windows; falharia em uma instalação não padrão (baixo risco prático, mas rígido).

---

## A.5 — Sheet3.cls (3 procedimentos)

### 1. `OptionButton1_Click` — Sheet3.cls:9
1. **Nome completo:** `OptionButton1_Click`
2. **Módulo:** Sheet3.cls (planilha "Base" consolidada)
3. **Tipo:** Evento (Click de OptionButton embutido na planilha)
4. **Escopo:** Private
5. **Objetivo:** Selecionar a primeira de três opções mutuamente exclusivas de um seletor de modo/visão na planilha Base, gravando a legenda escolhida em outra aba (`Sheet11!AV6`) para consumo por fórmulas/outras rotinas.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no `OptionButton1` embutido em `Sheet3`).
7. **Procedimentos chamados:** Nenhum (apenas leitura/escrita direta de propriedades de `Range`/`Font`).
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Sheet3.OptionButton1/2/3` (ActiveX embutidos, par mutuamente exclusivo — evidência: cada handler força negrito só no próprio e remove dos outros dois, linhas 11-13).
11. **Abas/intervalos/arquivos acessados:** `Sheet11!AV6` (escrita da legenda escolhida).
12. **Pré-condições:** Nenhuma.
13. **Passos principais:** Grava `Sheet3.OptionButton1.Caption` em `Sheet11.Range("AV6")`; aplica negrito à legenda do próprio botão e remove dos outros dois (simulação visual de estado "selecionado" já que `OptionButton`s de planilha não têm indicador visual tão claro quanto em UserForm).
14. **Pós-condições:** `Sheet11!AV6` contém o texto da opção 1; aparência visual (negrito) reflete a opção ativa.
15. **Efeitos colaterais / erro / mensagens / risco:** **RN-113 / Risco médio**: acoplamento direto entre `Sheet3` e uma célula fixa de `Sheet11` (`AV6`) sem camada de abstração (ex.: sem um nome definido/constante) — qualquer reorganização futura de `Sheet11` quebra silenciosamente esta gravação. Nenhuma validação de que `Sheet11` existe/está no estado esperado.

### 2. `OptionButton2_Click` — Sheet3.cls:16
1. **Nome completo:** `OptionButton2_Click`
2. **Módulo:** Sheet3.cls
3. **Tipo:** Evento (Click de OptionButton)
4. **Escopo:** Private
5. **Objetivo:** Idêntico ao anterior, para a segunda opção.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no `OptionButton2`).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Sheet3.OptionButton1/2/3`.
11. **Abas/intervalos/arquivos acessados:** `Sheet11!AV6`.
12. **Pré-condições:** Nenhuma.
13. **Passos principais:** Grava `Sheet3.OptionButton2.Caption` em `Sheet11.Range("AV6")`; aplica negrito ao próprio, remove dos outros dois.
14. **Pós-condições:** `Sheet11!AV6` contém o texto da opção 2.
15. **Efeitos colaterais / erro / mensagens / risco:** Mesma observação de RN-113.

### 3. `OptionButton3_Click` — Sheet3.cls:23
1. **Nome completo:** `OptionButton3_Click`
2. **Módulo:** Sheet3.cls
3. **Tipo:** Evento (Click de OptionButton)
4. **Escopo:** Private
5. **Objetivo:** Idêntico aos anteriores, para a terceira opção.
6. **Quem chama / gatilho:** Gatilho por interação do usuário (clique no `OptionButton3`).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** `Sheet3.OptionButton1/2/3`.
11. **Abas/intervalos/arquivos acessados:** `Sheet11!AV6`.
12. **Pré-condições:** Nenhuma.
13. **Passos principais:** Grava `Sheet3.OptionButton3.Caption` em `Sheet11.Range("AV6")`; aplica negrito ao próprio, remove dos outros dois.
14. **Pós-condições:** `Sheet11!AV6` contém o texto da opção 3.
15. **Efeitos colaterais / erro / mensagens / risco:** Mesma observação de RN-113. Nota adicional: o conteúdo textual (`Caption`) exato dos três `OptionButton`s — ou seja, quais são as 3 opções de negócio oferecidas — é **[NÃO ACESSÍVEL]** a partir deste dump de texto (propriedade `Caption` só está disponível no `.frx`/definição binária do controle).

---

## A.6 — ThisWorkbook.cls (1 procedimento)

### 1. `Workbook_Open` — ThisWorkbook.cls:9
1. **Nome completo:** `Workbook_Open`
2. **Módulo:** ThisWorkbook.cls
3. **Tipo:** Evento (Open do Workbook)
4. **Escopo:** Private
5. **Objetivo:** [NÃO IMPLEMENTADO] — nominalmente seria o ponto de inicialização automática do sistema ao abrir o arquivo, mas o corpo está vazio.
6. **Quem chama / gatilho:** Disparado automaticamente pelo Excel sempre que o workbook `.xlsb` é aberto (evento de ciclo de vida do host, não uma chamada de código).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum.
11. **Abas/intervalos/arquivos acessados:** Nenhum.
12. **Pré-condições:** N/A.
13. **Passos principais:** Nenhum — corpo do procedimento vazio (evidência: ThisWorkbook.cls, linhas 9-12, apenas `Private Sub Workbook_Open()` seguido de `End Sub`).
14. **Pós-condições:** Nenhuma alteração de estado provocada por este evento.
15. **Efeitos colaterais / erro / mensagens / risco:** **Fato confirmado, relevante para a reescrita**: não há nenhuma rotina de inicialização automática (sem splash screen, sem checagem de versão automática, sem carregamento de menu, sem verificação de ambiente/rede) disparada na abertura do arquivo. A checagem de versão (`Verifica_Versao`, Auxiliar.bas:82) é feita **manualmente**, ponto a ponto, antes de operações específicas (ex.: antes de `Form_Importacao.Show` em `Importar_Fronts`, Auxiliar.bas:58; antes de `Extrair_Todas_as_Bases`, Auxiliar.bas:8) — não há garantia de que a versão seja checada em todo fluxo possível de entrada do usuário no sistema.

---

## A.7 — Module2.bas (4 procedimentos)

### 1. `inverter_valores` — Module2.bas:2
1. **Nome completo:** `inverter_valores`
2. **Módulo:** Module2.bas
3. **Tipo:** Sub (macro gravada — evidência: comentário padrão "' inverter_valores Macro" e atributo `VB_ProcData.VB_Invoke_Func`, linhas 3-6, gerados automaticamente pelo gravador de macros do Excel)
4. **Escopo:** Public (implícito — sem `Private`)
5. **Objetivo:** Utilitário manual para multiplicar em massa um intervalo de valores por -1 (padrão comum de "inverter sinal" de uma linha de dados financeiros), a partir da célula `AJ1` até o fim da linha.
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump** (evidência: `grep -rn "inverter_valores\b" vba_dump_tmp/` só retorna a própria declaração/atributos/comentário). Execução manual pelo usuário avançado, provavelmente via `Application.Run` fora do fluxo de UI documentado, ou diretamente pelo Editor VBA/menu de macros do Excel (Alt+F8) — não está ligado a nenhum botão de planilha ou form identificado no código.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum controle de UI — opera sobre `Selection`/`Range("AJ1")` da planilha ativa no momento da execução (comportamento típico de macro gravada, dependente do estado de seleção do usuário).
11. **Abas/intervalos/arquivos acessados:** Planilha ativa no momento (não especifica qual — `ActiveSheet` implícito), intervalo a partir de `AJ1` até o fim da linha (`.End(xlToRight)`).
12. **Pré-condições:** Uma planilha estar ativa com dados a partir de `AJ1` na linha 1.
13. **Passos principais:** Seleciona de `AJ1` até o fim da linha para a direita; copia; cola especial multiplicando pelos valores atuais (`Operation:=xlMultiply`) — o efeito depende do que já está na área de transferência antes da execução (a macro copia a própria seleção e cola sobre si mesma multiplicando, o que é peculiar — ver risco).
14. **Pós-condições:** Valores do intervalo alterados conforme a operação de colagem especial.
15. **Efeitos colaterais / erro / mensagens / risco: ** **Risco baixo, mas comportamento pouco claro**: a macro copia o próprio intervalo (`Selection.Copy`) e imediatamente cola sobre ele mesmo com `Operation:=xlMultiply` — matematicamente isso eleva cada célula ao quadrado (`valor × valor`), o que não corresponde à descrição usual de "inverter valores" (que se esperaria ser multiplicação por -1). **Esta é uma leitura literal do código-fonte**; o efeito real só ficaria claro executando-a — classificar como [VALIDAR COM O NEGÓCIO] antes de decidir se replicar esta lógica ou se é um bug histórico nunca notado por falta de uso. Nenhuma confirmação, nenhuma mensagem, nenhum tratamento de erro — típico de macro gravada para uso pontual de um analista, não de uma funcionalidade de produto.

### 2. `inverte_valores_2` — Module2.bas:15
1. **Nome completo:** `inverte_valores_2`
2. **Módulo:** Module2.bas
3. **Tipo:** Sub (macro gravada)
4. **Escopo:** Public (implícito)
5. **Objetivo:** Variante da anterior — soma o intervalo `AJ1:<fim da linha>` a ele mesmo (`Operation:=xlAdd`), efetivamente dobrando os valores.
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump** (mesma evidência de busca). Execução manual.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum — mesmo padrão de `Selection`/`Range` da anterior.
11. **Abas/intervalos/arquivos acessados:** Planilha ativa, intervalo a partir de `AJ1`.
12. **Pré-condições:** Mesmas da anterior.
13. **Passos principais:** Seleciona `AJ1` até o fim da linha; copia; cola especial somando (`Operation:=xlAdd`) sobre si mesmo — efetivamente dobra os valores (`valor + valor = 2×valor`).
14. **Pós-condições:** Valores do intervalo dobrados.
15. **Efeitos colaterais / erro / mensagens / risco:** Mesmo padrão de risco/observação do item anterior — nome sugere "inverter" mas o efeito matemático real é "dobrar". [VALIDAR COM O NEGÓCIO] quanto ao uso real/intenção histórica antes de portar.

### 3. `inverter_valores_3` — Module2.bas:28
1. **Nome completo:** `inverter_valores_3`
2. **Módulo:** Module2.bas
3. **Tipo:** Sub (macro gravada)
4. **Escopo:** Public (implícito)
5. **Objetivo:** Prepara uma seleção auxiliar (`AJ1:AV283`) após escrever `-1` em `AW1` — o efeito de multiplicação em si **não é executado** neste procedimento (ele só seleciona, não cola); aparenta ser uma macro gravada interrompida antes da etapa final de "Colar Especial".
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump**. Execução manual (provavelmente nunca usada como está — ver risco).
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum.
11. **Abas/intervalos/arquivos acessados:** Planilha ativa; escreve fórmula `-1` em `AW1`; seleciona faixa fixa `AJ1:AV283`.
12. **Pré-condições:** Nenhuma verificada no código.
13. **Passos principais:** Escreve `-1` (como `FormulaR1C1`) em `AW1`; copia `AW1`; seleciona de `AJ1` para baixo e para a direita; por fim redefine a seleção explicitamente para o intervalo fixo `AJ1:AV283` (linha 42) — **e o procedimento termina aí**, sem colar (`PasteSpecial`) em lugar nenhum.
14. **Pós-condições:** `AW1` contém a fórmula `-1`; área de transferência contém a cópia de `AW1`; seleção ativa é `AJ1:AV283` — nenhuma alteração de dados de fato ocorre nas células de `AJ1:AV283`.
15. **Efeitos colaterais / erro / mensagens / risco:** **Risco baixo/confirma suspeita de macro incompleta**: comparado às outras 3 macros do módulo (que terminam com `.PasteSpecial ... Operation:=xlMultiply/xlAdd`), esta **não tem a linha de colagem final** — é evidência forte de uma gravação de macro interrompida/incompleta (o usuário provavelmente parou a gravação antes de executar o "Colar Especial" final, ou editou o código manualmente e esqueceu a última linha). Não deve ser replicada como está — comportamento incompleto. [VALIDAR COM O NEGÓCIO] se esta macro tinha uma versão funcional em algum outro lugar/versão do arquivo.

### 4. `inverter_valores_4` — Module2.bas:44
1. **Nome completo:** `inverter_valores_4`
2. **Módulo:** Module2.bas
3. **Tipo:** Sub (macro gravada)
4. **Escopo:** Public (implícito)
5. **Objetivo:** Multiplica um intervalo de dados (`AJ1` até o fim da linha e para baixo) por `-1`, usando `AX1` como célula auxiliar contendo o fator -1 — esta é a única das 4 macros do módulo que efetivamente completa a operação de "inverter sinal" (multiplicação por -1) de forma consistente com o nome do módulo.
6. **Quem chama / gatilho:** **Não referenciado em nenhum outro lugar do dump**. Execução manual.
7. **Procedimentos chamados:** Nenhum.
8. **Parâmetros:** Nenhum.
9. **Retorno:** Nenhum (Sub).
10. **Controles/variáveis relevantes:** Nenhum.
11. **Abas/intervalos/arquivos acessados:** Planilha ativa; escreve `-1` em `AX1`; aplica sobre a faixa a partir de `AJ1` (expandida para a direita e para baixo, `.End(xlToRight)` seguido de `.End(xlDown)`).
12. **Pré-condições:** Nenhuma verificada no código.
13. **Passos principais:** Escreve `-1` em `AX1`; copia `AX1`; seleciona a partir de `AJ1`, expande para a direita e depois para baixo (área retangular de dados); cola especial multiplicando (`Operation:=xlMultiply`, `Paste:=xlPasteAll`) — inverte efetivamente o sinal de toda a área de dados numéricos encontrada.
14. **Pós-condições:** Todos os valores numéricos da área a partir de `AJ1` (expandida) multiplicados por -1 (sinal invertido); `AX1` permanece com o valor `-1`.
15. **Efeitos colaterais / erro / mensagens / risco:** Risco baixo (não usada em nenhum fluxo automatizado, mas é a única das 4 macros que "faz o que o nome promete"). `Paste:=xlPasteAll` (em vez de `xlPasteValues`) também copia formatação/fórmulas da célula `AX1` para toda a área de destino — pode alterar inadvertidamente a formatação numérica/de célula da área de dados além do valor em si. Nenhuma confirmação/mensagem — mesma classe de macro "utilitário manual de analista" das demais deste módulo, não parte de fluxo de UI formal. **Nenhuma das 4 macros de Module2.bas está ligada a qualquer botão, menu ou evento identificado no restante do código** — recomenda-se [VALIDAR COM O NEGÓCIO] se essas macros ainda são usadas manualmente pela equipe antes de decidir se algo equivalente precisa existir na nova interface.

---

## A.8 — Sheet*.cls sem código associado (30 módulos, 0 procedimentos cada)

Confirmado por leitura integral de cada arquivo e por busca global (`grep -rn "Private Sub|Public Sub|Function" vba_dump_tmp/Sheet*.cls`, que retornou correspondências **apenas** em `Sheet3.cls` e `Sheet8.cls` — evidência de que todos os demais são estruturalmente idênticos: contêm somente os atributos padrão gerados automaticamente pelo VBA (`VB_Name`, `VB_Base`, `VB_GlobalNameSpace`, `VB_Creatable`, `VB_PredeclaredId`, `VB_Exposed`, `VB_TemplateDerived`, `VB_Customizable`), sem nenhum evento de planilha implementado (`Worksheet_Change`, `Worksheet_Activate`, `Worksheet_BeforeDoubleClick`, etc.) e sem nenhuma Sub/Function própria.

| Módulo | Linhas no arquivo | Observação |
|---|---|---|
| Sheet1.cls | 9 | Módulo vazio, sem código associado. |
| Sheet2.cls | 9 | Módulo vazio, sem código associado. |
| Sheet4.cls | 9 | Módulo vazio, sem código associado. |
| Sheet5.cls | 9 | Módulo vazio, sem código associado. |
| Sheet6.cls | 9 | Módulo vazio, sem código associado. |
| Sheet7.cls | 9 | Módulo vazio, sem código associado. |
| Sheet9.cls | 9 | Módulo vazio, sem código associado. |
| Sheet10.cls | 9 | Módulo vazio, sem código associado. |
| Sheet11.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet12.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet13.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet14.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet15.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet16.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet17.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet18.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet19.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet20.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet21.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet22.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet23.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet24.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet25.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet26.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet27.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet28.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet29.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet30.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet31.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |
| Sheet33.cls | [NÃO RELIDO NESTA RODADA] | Módulo vazio, sem código associado (confirmado via grep consolidado). |

Nota: Sheet1, Sheet2, Sheet4–Sheet7, Sheet9 e Sheet10 foram relidos integralmente nesta sessão (confirmação direta, 9 linhas cada, só atributos). Sheet11–Sheet31 e Sheet33 haviam sido confirmados como vazios em análise anterior desta mesma sessão de trabalho, por leitura integral prévia, e reconfirmados agora pela busca consolidada `grep -rn "Private Sub|Public Sub|Function" vba_dump_tmp/Sheet*.cls`, que não retornou nenhuma ocorrência para esses arquivos — evidência indireta, mas de alta confiança (busca de padrão cobre 100% do conteúdo textual de cada arquivo). Não há função de negócio nem interatividade codificada nessas 30 abas — qualquer comportamento que essas planilhas exibam vem de fórmulas, formatação condicional ou controles ActiveX/Forms ligados a macros de outros módulos (fora do escopo deste cluster).

---

