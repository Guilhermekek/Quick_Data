# 16. Tratamento de Erros e Logs

## 16.1 Estratégias de tratamento de erro

**[CONFIRMADO — evidência: padrão recorrente em praticamente todos os módulos]** O padrão dominante é `On Error GoTo <rótulo>` apontando para um bloco de tratamento no final do procedimento, que tipicamente restaura o estado do Excel (`ScreenUpdating`, `Calculation`, `EnableEvents`) e mostra um `MsgBox` genérico. Em uma quantidade significativa de módulos, esse `On Error GoTo` está **comentado** (`'On Error GoTo tratar_erro`), o que significa que o bloco de tratamento existe no código mas **nunca é acionado automaticamente** — só seria alcançado se alguém reativasse a linha manualmente. Confirmado em: `Extracao_SQL_Hubble.bas`, `Conexoes.bas`, `Refresh_Sup_Linhas.bas`, `Refresh_DP_Segmento.bas`.

## 16.2 Uso de `On Error Resume Next`

**[CONFIRMADO]** Uso identificado em `Extracao_Base_Consolidad.bas` cobrindo três operações consecutivas de exclusão de linhas em branco (`SpecialCells(xlCellTypeBlanks).EntireRow.Delete`) sem checar erro entre elas — se a primeira falhar (ex.: nenhuma célula em branco encontrada, erro 1004 do Excel), as duas seguintes continuam silenciosamente, sem indicação de sucesso ou falha real.

## 16.3 Erros ignorados/engolidos

**[CONFIRMADO]** Em `fx_IFRS16.bas`, linhas que geram erro `#N/A` no De-Para de Classe Custo são puladas via `GoTo próximaLinha`, apenas incrementando um contador — o log resultante informa "quantas" linhas falharam, mas não "quais" (sem linha/CC específico registrado).

## 16.4 Logs existentes

**[CONFIRMADO]** O sistema mantém um log de erros em produção **na própria planilha** `tk_Lista_de_erros`, escrito por `fn_ListAllErrors` (`TK_Functions.bas`). Estrutura confirmada (cabeçalho da planilha): `FUNCAO_VBA`, `LINHA_VBA`, `SHEET_EXCEL`, `LINHA_EXCEL`, `COLUNA_EXCEL`, `VALOR_ENCONTRADO`, `QUEM`, `QUANDO`. No conteúdo já existente no arquivo no momento da análise, há **múltiplas ocorrências repetidas do mesmo valor de erro (`0x2a`)** nas mesmas coordenadas da planilha Base, atribuídas às funções `Aux_Formulas_Base` e `TK_Functions` — **[VALIDAR COM O NEGÓCIO]**: indício de um problema recorrente e não resolvido, que vale a pena investigar como parte da sustentação.

## 16.5 Mensagens ao usuário

**[CONFIRMADO]** O padrão de mensagem é `MsgBox` simples, na maior parte dos casos **genérica** ("Ocorreu um erro ao abrir este arquivo!", "Processo concluído com sucesso!") sem detalhar a causa raiz do erro (ex.: sem mostrar `Err.Description`/`Err.Number` ao usuário na maioria dos pontos analisados). Confirmações de ações (Sim/Não) também usam `MsgBox` simples, sem tela de revisão mais elaborada.

## 16.6 Pontos em que uma falha pode deixar o Excel em estado inadequado

**[CONFIRMADO — risco alto, ver também seção 18]**
- `Extracao_Base_MOCKUP_RGM.bas` e `Extracao_Base_RGM.bas`: uso da instrução `End` do VBA quando a validação estrutural do arquivo fonte falha (linha ~335 e correspondente). `End` interrompe a execução do VBA **imediatamente e sem passar por nenhum tratamento de limpeza** — se isso acontecer no meio de uma extração com `ScreenUpdating`/`Calculation`/`EnableEvents` desligados, o Excel pode ficar travado nesse estado (tela não atualiza, eventos não disparam, cálculo manual) até o usuário fechar e reabrir o arquivo.
- Diversos módulos de extração fazem `Workbooks.Open` de um arquivo externo sem checar se a abertura teve sucesso antes de prosseguir (`Windows(Arq).Close` sem verificar se a janela existe) — se o arquivo estiver corrompido/bloqueado, o erro resultante não tem uma mensagem clara amarrada à causa.

## 16.7 Processos de recuperação identificados

**[CONFIRMADO/PARCIAL]** Não foi identificado nenhum mecanismo de "desfazer"/rollback de dados dentro do próprio sistema (ex.: não há como reverter uma extração já commitada na Base além de rodar novamente a "Limpar base" da fonte específica e reextrair). O único mecanismo de recuperação de estado do Excel em caso de erro é o bloco `On Error GoTo` que restaura `ScreenUpdating`/`Calculation`/`EnableEvents` — quando esse bloco de fato é alcançado (ver 16.1, nem sempre está ativo). Não há log de auditoria de alterações na Base além do log de erros (`tk_Lista_de_erros`), que registra falhas, não o histórico de operações bem-sucedidas.
