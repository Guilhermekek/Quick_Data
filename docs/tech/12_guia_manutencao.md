# 19. Guia de Manutenção

> Esta seção é **recomendação**, não fato extraído do arquivo — está identificada como tal conforme exigido no escopo desta análise.

## 19.1 Checklist antes de modificar

1. **Nunca edite o arquivo produtivo diretamente.** Trabalhe sempre em uma cópia local, fora de qualquer pasta de rede sincronizada/compartilhada, até a alteração estar validada.
2. Identifique, no Catálogo de Procedimentos (seção 10), todos os procedimentos que **chamam** e são **chamados** pelo procedimento que você vai alterar (campos 6 e 7 do catálogo).
3. Verifique se o procedimento tem uma versão duplicada com nome parecido (padrão confirmado neste sistema: sufixo `_TK`, `_OLD`) — alterar só uma cópia sem verificar a outra é a causa mais provável de "correção que não funciona" ou de divergência silenciosa entre módulos.
4. Confira, na seção 18 (Riscos), se o componente que você vai tocar já está listado — leia os "cuidados antes de alterar" registrados lá.
5. Se a alteração tocar em qualquer procedimento ligado a uma Regra de Negócio (seção 12), confirme com o time de negócio se a regra ainda é válida **antes** de alterar o código — várias regras deste sistema estão marcadas `[VALIDAR COM O NEGÓCIO]` porque a análise de código sozinha não conseguiu confirmar se ainda estão em uso.
6. Faça backup do arquivo (cópia com timestamp) antes de rodar qualquer macro de limpeza/saneamento (ex.: `CLEAR_Defined_Names`) — são operações destrutivas sem undo.

## 19.2 Como avaliar impacto

- Use a coluna "Abas/intervalos/arquivos acessados" do catálogo de procedimentos (seção 10) para levantar todas as planilhas tocadas, direta ou indiretamente (via procedimentos chamados).
- Use o Mapa de Dependências (seção 8) para verificar se a aba/tabela que você vai alterar alimenta outras partes do sistema — a planilha "Base" e as tabelas mestre (Sup_Linhas, DP_Segmento, Ref_Cruzada) têm o maior raio de impacto por serem lidas por praticamente todo o sistema.
- Se a alteração for em uma fórmula gerada dinamicamente por VBA (padrão `FormulaR1C1` + bake-in em valor, ver seção 15), lembre que **a fórmula não fica visível na planilha depois de rodar** — para conferir o efeito, é preciso rodar a macro de novo e inspecionar o resultado, não basta olhar a célula.

## 19.3 Componentes que devem ser testados juntos

Com base nas dependências de ordem de execução confirmadas (seção 8):

| Se alterar... | Teste também... |
|---|---|
| `Sup_Linhas`, `DP_Segmento`, `Ref_Cruzada_1/2`, `DropComb` (tabelas mestre) | Todo o pipeline `Refresh_Base_Aux` (ordem fixa, F-01) + qualquer extração que rode depois (F-02/F-03) + `Limpeza_Base_Ajustes`/`Lista_Validacao_Ajustes` (dependem dessas tabelas) |
| Qualquer módulo `Extracao_*` | O motor de enriquecimento comum (`Aux_Formulas_Base.bas`) — todas as fontes passam pelas mesmas funções `Form_*` |
| `Form_Segmentos` (motor de rateio) | O dashboard `DP_Rateio` (tabelas dinâmicas/slicers) e qualquer relatório que dependa da coluna Segmento/Abertura_1 |
| Layout da aba "Extração" (Sheet8) | Todos os módulos `Extracao_*` e `Aux_Leitura_Nome_Arqs.bas` — dependem de `Cells.Find` sobre texto/posição nessa aba (ver riscos, seção 18) |
| `Auxiliar.bas` (reclassificações) | `TK_Functions.bas` (versões `_TK`) — verificar se a mudança precisa ser replicada nas duas cópias |
| `fx_IFRS16.bas` | Colunas 71/73/15/22 e `BX1` de `Sup_Linhas` — confirme que o De-Para manual está atualizado antes de testar |

## 19.4 Checklist posterior à alteração

1. Rode o fluxo completo afetado (não só o procedimento alterado) em uma cópia de teste, com dados reais de um período fechado (para poder comparar contra um resultado conhecido).
2. Compare o resultado da planilha "Base" (contagem de linhas, soma de FY por Fonte/KPI) antes e depois da alteração.
3. Verifique a planilha `tk_Lista_de_erros` depois da execução — novas entradas indicam problema introduzido.
4. Se a alteração tocou em regra de rateio ou IFRS16, valide manualmente uma amostra de linhas contra o cálculo esperado (essas são as regras de maior risco financeiro).

## 19.5 Estratégia de controle de versão (recomendação)

O arquivo hoje não tem nenhum controle de versão formal (é um único `.xlsb` binário). Recomenda-se, no mínimo: manter cópias datadas antes de cada alteração relevante (`Quick Data 3.23 - AAAA-MM-DD - antes de <descrição>.xlsb`); documentar a alteração na tabela de Controle de Versões deste documento; considerar exportar o código VBA para arquivos `.bas`/`.cls`/`.frm` de texto (via `oletools` ou o próprio VBE com um add-in de exportação) e versionar esses arquivos de texto em um sistema como Git, mesmo enquanto o sistema continuar sendo um `.xlsb` — isso permite comparar (`diff`) mudanças de código entre versões, o que é impossível de fazer no binário diretamente.

## 19.6 Procedimento de rollback (recomendação)

Como não há mecanismo de undo dentro do sistema (seção 16.7), o rollback é sempre "restaurar a cópia de backup anterior à alteração". Antes de aplicar qualquer alteração em produção: garanta que existe uma cópia de backup íntegra e testável fora da pasta de trabalho ativa; comunique a janela de mudança para os usuários (o arquivo não suporta edição concorrente seura — é um único arquivo Excel).

## 19.7 Cuidados para não trabalhar no arquivo produtivo

- O arquivo de produção deve ser tratado como "somente leitura para experimentação" — qualquer teste de macro deve rodar em cópia local.
- Cuidado especial com as macros de saneamento (`CLEAR_Defined_Names`, `RUN_Apagar_defined_names_definitivamente`, `Limpar_Todas_as_Bases`) — são destrutivas e não têm confirmação em todos os pontos.
- O projeto VBA está protegido por senha (seção 17.3) — confirme que a senha é conhecida pela equipe antes de depender de poder editar o código diretamente no Excel; se não for, o fluxo de manutenção precisa passar por ferramentas externas de parsing/edição do binário, o que é mais arriscado e deve ser evitado como prática recorrente (usar apenas para recuperação emergencial).
