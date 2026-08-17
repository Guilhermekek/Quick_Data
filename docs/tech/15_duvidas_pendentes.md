# 23. Dúvidas e Pontos Pendentes

> Itens que exigem confirmação de alguém com acesso ao Excel/VBE ou conhecimento de negócio — a análise de código, sozinha, não conseguiu resolvê-los. Cada item indica por que é uma dúvida e não uma conclusão.

## 23.1 Requerem acesso interativo ao arquivo (VBE/Excel aberto)

1. **Proteção de planilha e de workbook** — não acessível pelas ferramentas usadas (seção 4.4). Verificar em cada aba (Revisão → Proteger Planilha) e no workbook (Revisão → Proteger Pasta de Trabalho).
2. **Mapeamento completo CodeName ↔ nome de aba** — vários CodeNames permanecem [NÃO IDENTIFICADO] na seção 7. Confirmação rápida: abrir o VBE (Alt+F11) e olhar a janela "Project Explorer", que mostra `SheetN (NomeDaAba)` lado a lado.
3. **Senha do projeto VBA** — o projeto está protegido (seção 17.3). Confirmar se a senha é conhecida/documentada pela equipe de sustentação atual; se não for, isso trava qualquer edição direta do código pelo Excel.
4. **Ligação exata entre cada shape/botão e a macro que ele dispara** — os rótulos dos botões foram confirmados (seção 7, ex.: "EXTRAIR BASE HUBBLE"), mas a associação shape→macro (`OnAction`) fica em uma propriedade do objeto de desenho não incluída no dump de texto do VBA. Confirmar clicando com o botão direito em cada shape → "Atribuir Macro".
5. **Formatação condicional relevante ao processo** — não foi extraída sistematicamente (seção 4.4).
6. **Arquitetura do Office (32/64 bits)** — inferida como provável 64 bits pelo caminho de referência (seção 17.1), não confirmada.
7. **Confirmação visual dos controles dos UserForms** — tipos/posições foram inferidos pelo uso no código (seção 9.4), não lidos diretamente do `.frx`.

## 23.2 Requerem confirmação de negócio

8. **Divergência de número de versão**: o arquivo se chama "3.23", mas o gate de versão em `Verifica_Versao` (`Auxiliar.bas`) compara contra a string `"3.0"`. Esses dois números têm significados diferentes (ex.: um é "versão de arquivo", outro "versão de compatibilidade de motor") ou o gate está desatualizado?
9. **Status real de `BackupCodigo_MainResults`** (validação automática de hierarquia de dropdowns no Main Results): a estrutura do código (evento `Worksheet_Change` dentro de um módulo `.bas` comum, não no code-behind da planilha) sugere fortemente que está inativo — mas isso precisa ser confirmado interativamente (ver CT-17) ou perguntando a quem usa a planilha Main Results hoje.
10. **Quem mantém o De-Para de IFRS16** (colunas 71, 73, 15, 22 e `BX1` de `Sup_Linhas`) — não é alimentado por nenhuma macro de refresh identificada. Precisa ser localizado antes de qualquer migração, sob risco de perda de conhecimento tácito.
11. **Significado exato de "Front"** — inferido como "aba de relatório/layout" pelo uso em `Form_Importacao`/`Form_Exportacao`, nunca definido explicitamente no código (seção 22).
12. **`LISTA_ARQ_AUX` é documentação viva ou morta?** — a aba lista caminhos de rede e nomes de tabela, mas nenhuma macro lida a usa diretamente (os caminhos reais de configuração ficam em `Sheet24`). Confirmar se alguém atualiza `LISTA_ARQ_AUX` manualmente como referência ao mudar `Sheet24`, ou se ela está desatualizada.
13. **Fallback do rateio variável sem percentual cadastrado** (CT-16) — comportamento não determinável por leitura de código; precisa de teste dirigido ou confirmação de quem mantém `Sh_Rateio`.
14. **Inconsistência de nome de função em `Extracao_Sheet_Ajustes.bas`** (chama `Atualizar_Lista_KPI_Versao`, os outros 10 módulos de extração chamam `Atualizar_Lista_KPI_Versao_Interna`) — bug (funcionalidade errada por engano) ou distinção proposital não documentada?
15. **Eficácia real do bloqueio de digitação em `Form_Importacao`** (RN-101): o código tenta `Cancel = True` num evento `KeyPress` de TextBox, o que não é um parâmetro válido para esse tipo de controle em VBA — o bloqueio pode não funcionar de fato. Precisa de teste funcional no arquivo real.
16. **Referência VBA quebrada** (GUID zerado, seção 17.2) — a que biblioteca ela apontava originalmente? Pode ter sido uma biblioteca removida do Windows/Office da máquina onde o projeto foi salvo por último; não afeta a execução atual claramente, mas pode gerar erro "não foi possível encontrar referência" em outra máquina.
17. **`Registrar_Sheet` referenciado em `Form_Exportacao.frm` (`ListBox1_DblClick`) sem definição encontrada no restante do código lido** — procedimento removido, renomeado, ou existe em um módulo não coberto por esta análise? Se o controle `ListBox1` ainda existir na interface real, clicar duas vezes nele pode gerar erro.
18. **Ordem de execução esperada não é imposta pelo sistema**: o fluxo "primeiro atualizar bases auxiliares, depois extrair" é uma convenção observada na estrutura dos botões, mas **nada no código impede** o usuário de extrair antes de atualizar as tabelas mestre, potencialmente usando regras de classificação desatualizadas sem aviso. Confirmar se isso já causou problema conhecido.
19. **Origem exata do sistema de onde vem a "Base 1009"** e demais nomes de fonte (RGM, Fixed Revenues) — os nomes são conhecidos, mas o sistema/processo de origem de cada um não está documentado no código (fora do escopo de uma análise de VBA).
20. **Perfis de usuário** — o sistema não distingue usuários tecnicamente (seção 5.5); confirmar se isso é aceitável para o processo de negócio atual ou se há controle de acesso fora do arquivo (ex.: permissão na pasta de rede).

## 23.3 Erros recorrentes já observados, sem causa raiz confirmada

21. **Erro `0x2a` recorrente** em `tk_Lista_de_erros`, atribuído a `Aux_Formulas_Base`/`TK_Functions`, sempre nas mesmas coordenadas da planilha Base — vale investigar como item de sustentação prioritário, já que é um problema que o próprio sistema já registrou repetidamente sem resolução aparente.
