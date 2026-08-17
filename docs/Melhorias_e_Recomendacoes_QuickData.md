# Quick Data 3.23 — Melhorias e Recomendações

**Sistema:** Quick Data 3.23
**Base desta análise:** [Documentação Técnica e Funcional do Quick Data 3.23](Documentacao_Tecnica_Funcional_QuickData.docx) — 192 procedimentos catalogados, 112 regras de negócio (RN-001 a RN-115) e **46 riscos classificados por severidade** (9 críticos, 17 altos, 14 médios, 6 baixos)
**Data:** 2026-08-14
**Status:** Substitui a versão anterior deste documento (`02_melhorias_e_recomendacoes.md`), que foi escrita antes do catálogo completo de procedimentos existir — os achados abaixo são todos rastreáveis a um ID de risco (D-xx/R-xx) e/ou regra de negócio (RN-xxx) específico da documentação técnica, não a impressões gerais.

---

## 1. Sumário executivo

Dos 46 riscos catalogados, **9 são críticos** e **17 são altos** — 26 dos 46 (57%) exigem atenção antes ou durante qualquer trabalho de manutenção no sistema atual. A natureza desses riscos é overwhelmingly **estrutural**: duplicação de lógica de negócio em implementações paralelas divergentes, ausência de tratamento de erro que causa parada abrupta do Excel, conhecimento de negócio mantido fora de qualquer automação, e um achado de segurança que deveria ter sido corrigido independentemente de qualquer decisão de reescrita.

Isso confirma o diagnóstico da primeira versão deste documento: **os problemas não são pontuais, são de arquitetura** — a decisão de reconstruir o sistema como executável standalone continua sendo a resposta correta. O que muda nesta versão é a precisão: cada recomendação abaixo aponta para o procedimento, linha e regra de negócio exatos, não para uma categoria genérica de problema.

---

## 2. Ação imediata — antes de qualquer outra decisão

### 🔴 Credenciais de banco de dados em texto plano (D-01 / R-05 — CRÍTICO, achado transversal)

Confirmado de forma idêntica por 2 das 4 frentes de análise (o achado aparece em `Conexoes.bas`, chamado por módulos de 3 dos 4 clusters — ETL, Core e Refresh): usuário e senha do SQL Server (`AdminBPAM`/catálogo `BPAM`, fallback `Report2BPAM`.`URELATIG`/catálogo `InfoGER`, com a **mesma senha reutilizada entre os dois pares de login**) estão hardcoded em texto plano no VBA (`Conexoes.bas:26-27, 36-37`). Qualquer pessoa com acesso ao arquivo `.xlsb` tem acesso de leitura/escrita ao banco corporativo.

- **Custo de corrigir:** baixo — rotacionar a senha no SQL Server, revisar quem tem o arquivo.
- **Benefício:** altíssimo — elimina exposição de credencial institucional a um universo indeterminado de pessoas.
- **Ação:** rotacionar a senha agora, independentemente do cronograma de qualquer reescrita. Confirmar com o usuário se isso já foi feito desde a análise técnica.

### 🟠 SQL não parametrizado (D-02 — CRÍTICO)

Todo o SQL dinâmico da extração Hubble é montado por concatenação direta de valores de planilha (`Extracao_SQL_Hubble.bas:399,485-498`), sem uso de parâmetros ADODB. O risco prático de injeção maliciosa é menor (os valores vêm de configuração interna, não de input externo controlado), mas o risco de **quebra operacional real** é alto: qualquer apóstrofo em um nome de KPI/Empresa quebra a query.

- **Ação:** ao portar a extração Hubble para a reescrita, usar SQL parametrizado desde o primeiro dia (não é um "nice to have" — RN-018 mostra que a lógica de filtro já é complexa o suficiente para se beneficiar de parâmetros nomeados).

### 🟡 `Registrar_Sheet` — chamada a procedimento inexistente (UI-1 — CRÍTICO, requer confirmação urgente)

`Form_Exportacao.ListBox1_DblClick` (linha 469) chama um procedimento `Registrar_Sheet` que **não existe em nenhum lugar do código extraído**. Em VBA, isso normalmente impede a compilação do projeto inteiro assim que qualquer macro roda. Duas hipóteses, nenhuma confirmada: (a) o projeto de produção realmente tem esse erro de compilação latente, potencialmente quebrando **todo** o VBA na primeira execução que tocar esse caminho; ou (b) o `ListBox1` está órfão/inatingível na interface real e esse código nunca executa.

- **Ação imediata e barata:** abrir o arquivo original no VBE e rodar `Debug → Compile VBAProject`. Se der erro, é um problema ativo em produção que precisa ser corrigido (ou o `ListBox1` removido) antes de qualquer outra manutenção no Form_Exportacao. Se compilar limpo, o achado vira apenas uma nota de código morto — mas isso só se sabe testando.

---

## 3. Diagnóstico consolidado por categoria

Cada item cita o ID exato no catálogo de riscos (seção 18 do documento técnico) para rastreabilidade.

### 3.1 Segurança

| Achado | ID | Severidade |
|---|---|---|
| Credenciais SQL em texto plano, senha reutilizada entre 2 logins | D-01 / R-05 | Crítico |
| SQL construído por concatenação de string | D-02 | Crítico |
| Nenhum controle de quem pode editar as tabelas mestre (Sup_Linhas, De-Para) — qualquer usuário com o arquivo altera regra de negócio sem trilha de auditoria além do log de erros | — (achado transversal, seção 17.9 do doc técnico) | Alto (não classificado individualmente no catálogo, mas decorre de D-01 combinado à ausência de controle de acesso) |

### 3.2 Robustez / continuidade operacional

| Achado | ID | Severidade |
|---|---|---|
| `End` abrupto do VBA em divergência estrutural (RGM/MOCKUP_RGM) — mata o Excel sem limpar estado | D-03 | Alto |
| `MsgBox` bloqueantes no meio do pipeline — impede execução desatendida/agendada | D-04 | Alto |
| Tabela de staging SQL (Hubble) pode ficar "presa" se a execução falhar no meio | D-05 (ETL) | Alto |
| De-Para de IFRS16 mantido 100% fora de qualquer automação rastreável | R-01 | Crítico |
| Status de ativação da validação de hierarquia do Main Results é incerto | R-02 | Crítico |
| `Atualizar_Validacao_Linhas_Geral` está incompleta mas reporta sucesso ao usuário | R-08 | Médio |
| Ausência de limpeza prévia em 4 fontes (MOCKUP_RGM, RGM, Fixed_Revenues, Quick_Data) — reextrair duplica dados | D-06 | Médio |
| Erros de De-Para descartados sem log individual (IFRS16, Ref Cruzada) | R-04 | Alto |
| Dependência de ordem de execução (Refresh de tabelas mestre → limpeza de Ajustes) não garantida por código | R-09 | Médio |
| Referência VBA quebrada (GUID zerado) no projeto | — (seção 17.2 do doc técnico) | Baixo |

### 3.3 Duplicação e divergência de regra de negócio

Esta é a categoria com o achado individualmente mais grave do sistema, pelo raciocínio de impacto financeiro direto:

| Achado | ID | Severidade |
|---|---|---|
| Reclassificação de Empresa/IFRS/Proforma implementada 2x (`Auxiliar.bas` com busca dinâmica vs. `TK_Functions.bas` `_TK` com linha 6 fixa), com **3 fórmulas diferentes entre si** para calcular fim de range dentro da própria família `_TK` | D-1 (Core) | Crítico |
| `Ordenar_Lista` existe em 2 escopos diferentes (Public em Auxiliar.bas, Private em BackupCodigo_MainResults.bas) — qual versão roda depende de quem chama | D-4 (Core) | Alto |
| `Processo_Extracao_Sheet_Base` triplicada (RGM/MOCKUP_RGM/Fixed_Revenues) com pequenas variações não documentadas | D-07 (ETL) | Médio |
| Regra "N203073156+NT→382" copiada 3x sem comentário do motivo de negócio | D-6 (Core) | Alto |
| Regra de negócio (cascata Ref Cruzada, De-Para) expressa como string de fórmula Excel dentro do VBA, não como função testável | R-03 | Alto |

**Por que isso é o achado mais caro do sistema:** o risco não é "o código é feio" — é que **uma correção aplicada em uma cópia da regra de reclassificação Empresa/IFRS/Proforma não é automaticamente aplicada na outra**, e os dois caminhos são acionados por gatilhos de negócio reais e diferentes (extração completa vs. fluxo pós-IFRS16). Isso significa que, hoje, é plausível que a classificação de uma mesma linha varie dependendo de qual botão gerou aquele dado — um problema de **qualidade do dado financeiro reportado**, não só de código.

### 3.4 Performance

| Achado | ID | Severidade |
|---|---|---|
| `EntireRow.Copy`/`EntireRow.Delete` em laços sobre a Base inteira (O(n²)) | D-5 (Core) | Alto |
| Manipulação de `SlicerCaches` item a item dentro do motor de rateio | D-8 (Core) | Alto |
| `INDIRECT()` volátil aplicado em massa | D-11 (Core) | Médio |
| Array fixo `Campos_Chaves(10000)` sem verificação de limite | D-08 (ETL) | Médio |
| Posicionamento de coluna por deslocamento aritmético em vez de nome (frágil a reordenação, não só lento) | R-06 | Médio |

O padrão "escrever fórmula em massa → calcular → copiar → colar valor", repetido dezenas de vezes por execução completa, é a causa raiz mais provável da lentidão reportada pelos usuários — mas note que **nenhum item de performance está classificado como crítico**: os riscos críticos do sistema são todos de corretude/segurança, não de velocidade. Isso é relevante para priorização (seção 4).

### 3.5 Manutenibilidade / código morto

| Achado | ID | Severidade |
|---|---|---|
| Ao menos 13 procedimentos do cluster Core são código morto/órfão confirmado (`Form_Segmentos_OLD`, `Testeee`, `Teste` com chamada quebrada etc.) | D-10 (Core) | Médio |
| 9 módulos inteiros (`Module1`–`Module10`, exceto Module2) são resíduo de gravação de macro/depuração | — (seção 9.3 do doc técnico) | Baixo |
| Falta `Option Explicit`/tipagem em todo o projeto | D-12 (Core) | Médio |
| Blocos extensos de código comentado ao lado do código ativo | D-14 (Core), D-10 parcial (Refresh) | Baixo |
| 166 nomes definidos, 42 já sinalizados pelo próprio arquivo como quebrados (links de 1998-2014) | — (seção 24.3 do doc técnico) | Médio (causa clássica de lentidão/corrupção de `.xlsb`) |
| Botões de UI (`Shape.OnAction`) não rastreáveis por análise textual — pelo menos 12 procedimentos "de botão" sem chamador encontrado no código | D-7 (Core) | Alto (risco de mapear incorretamente UI→lógica na reescrita) |

### 3.6 Experiência do usuário / operação

| Achado | ID | Severidade |
|---|---|---|
| Processamento longo sem barra de progresso real nem cancelamento (Importação/Exportação) | UI-3 | Alto |
| Importação acoplada a rótulos textuais fixos, sem validação de schema | UI-4 | Alto |
| Sobrescrita silenciosa de arquivo exportado (`Kill` sem aviso), risco real se 2 exportações caírem no mesmo minuto | UI-2 | Crítico |
| Inconsistência clique-simples vs. duplo-clique entre Form_Importacao e Form_Exportacao | UI-5 | Médio |
| Corte silencioso de listas de validação acima de 747 itens | R-07 | Médio |
| Filtro de seleção de arquivo sem restrição de tipo (`*.*`) | D-12 (ETL) | Baixo |

---

## 4. Avaliação de melhorias — vale a pena?

| # | Iniciativa | Risco(s) endereçado(s) | Custo | Benefício | Vale a pena? |
|---|---|---|---|---|---|
| 1 | Rotacionar credenciais SQL, sair do padrão hardcoded | D-01/R-05 | Baixo | Altíssimo | ✅ Já, fora do cronograma da reescrita |
| 2 | Testar compilação do VBAProject e resolver `Registrar_Sheet` | UI-1 | Baixo (1 clique no VBE) | Alto (evita presumir estabilidade que pode não existir) | ✅ Já — é praticamente gratuito |
| 3 | Reescrita completa como executável Python (decisão já tomada) | Resolve estruturalmente D-1/D-3/D-4/D-5/D-7/D-8/D-11/D-12 (Core), D-03/D-04/D-07/D-08 (ETL), R-03/R-06 (Refresh) e toda a seção 3.6 | Alto | Alto — é a única forma de eliminar a *causa*, não só o sintoma, da maioria dos 46 riscos | ✅ Confirmado pelo catálogo: 26 dos 46 riscos são estruturais, não corrigíveis com patch pontual |
| 4 | Consolidar as 2 famílias de reclassificação (`Auxiliar` vs. `_TK`) numa única implementação | D-1 (crítico) | Médio | Crítico — elimina o risco de classificação financeira divergente | ✅ Prioridade #1 técnica dentro da reescrita — não fazer isso é herdar o pior bug do legado |
| 5 | Decidir explicitamente sobre `Registrar_Sheet`/ListBox1 antes de portar Form_Exportacao | UI-1 | Baixo | Alto | ✅ Pré-requisito de UI-3/UI-4 abaixo |
| 6 | Externalizar regras hardcoded (De-Para de CC, exceção 5G/METIS_CZ, código 382) para tabela de configuração | D-6, RN-039, RN-066 e outras | Médio | Alto — auditável e editável sem programador | ✅ Princípio de design da versão nova |
| 7 | Progresso real + cancelamento + validação de schema na importação/exportação | UI-2, UI-3, UI-4 | Médio | Alto (adoção, confiança, menos risco de arquivo corrompido por "Ctrl+Alt+Del") | ✅ Natural ao reescrever a UI |
| 8 | Validar com o negócio: De-Para IFRS16 (R-01), status do BackupCodigo_MainResults (R-02), `Atualizar_Validacao_Linhas_Geral` (R-08), macros de Module2 ainda usadas manualmente (UI-8) | R-01, R-02, R-08 | Baixo (é uma rodada de entrevistas) | Crítico — sem isso a reescrita corre risco de reproduzir números errados (IFRS16) ou reimplementar telas que ninguém usa | ✅ Fazer **antes** de portar a lógica correspondente, não depois |
| 9 | Log estruturado de erro de De-Para (linha + chave + motivo) em vez de contador agregado | R-04 | Baixo-médio | Alto (auditoria/reconciliação) | ✅ Baixo custo incremental dentro da reescrita |
| 10 | Limpar os 42 nomes definidos quebrados e os 9 módulos de macro gravada residual | seção 3.5 | Baixo | Baixo-médio, só temporário | ⚠️ Só vale a pena se o Excel legado precisar sobreviver mais alguns meses em paralelo — descartável se a migração for rápida |
| 11 | Corrigir isoladamente os padrões O(n²)/Slicers/INDIRECT no VBA atual | seção 3.4 | Médio-alto | Médio, ganho temporário | ❌ Não vale a pena isoladamente — o ganho desaparece quando o Python substituir a Base |
| 12 | Ativar `Option Explicit` retroativamente no legado | D-12 (Core) | Alto (exige testar cada módulo) | Baixo no legado | ❌ Não fazer no legado sem suíte de testes; aplicar tipagem forte (type hints) só na reescrita |

**Conclusão:** nenhum item novo muda a conclusão da versão anterior deste documento — a reescrita continua sendo a decisão certa. O que o catálogo completo adiciona é uma **ordem de prioridade dentro da reescrita**: o item 4 (consolidar as famílias de reclassificação duplicadas) é, pela primeira vez, identificável como o risco técnico individualmente mais caro do sistema — mais do que a lentidão, que é o sintoma mais visível mas não o mais perigoso.

---

## 5. Pontos que exigem validação de negócio antes de portar qualquer lógica

Lista mínima — cada um está detalhado com a evidência completa na seção 23 do documento técnico:

1. **De-Para de IFRS16** (colunas 71/73/15/22 de Sup_Linhas) — quem mantém, com que fonte de verdade, com que frequência.
2. **`BackupCodigo_MainResults`** — a validação automática de hierarquia de dropdowns está realmente ativa hoje?
3. **`Atualizar_Validacao_Linhas_Geral`** — o botão correspondente ainda está exposto na UI? Alguém presume que ele funciona?
4. **Macros de `Module2.bas`** (`inverter_valores*`) e demais utilitários órfãos — algum analista ainda os aciona manualmente via Alt+F8?
5. **`Form_Tratamento_Opcoes`** — tela usada por alguém hoje, ou pode ser descartada na reescrita?
6. **Divergência de versão** (arquivo "3.23" vs. gate de código "3.0") — os dois números significam coisas diferentes ou o gate está desatualizado?
7. **Ausência de limpeza prévia em 4 fontes** (D-06) — existe um processo manual de deduplicação hoje que compensa isso?

---

## 6. Roadmap sugerido (atualizado)

1. **Fase 0 — Mitigação imediata (dias, não semanas):**
   - Rotacionar credenciais SQL (seção 2).
   - Testar `Debug → Compile VBAProject` no arquivo original e resolver o achado `Registrar_Sheet` se confirmado.
   - Rodar a rodada de entrevistas da seção 5 — idealmente com as mesmas pessoas que ainda sustentam o sistema hoje, antes que esse conhecimento se perca.

2. **Fase 1 — Motor de dados, na ordem de risco, não na ordem de módulo:**
   - Portar primeiro a lógica de reclassificação Empresa/IFRS/Proforma como **uma única implementação testável**, resolvendo D-1 na origem em vez de portar as duas versões.
   - Em seguida, o motor de rateio (`Form_Segmentos`) e o tratamento IFRS16 — as duas peças de maior risco financeiro direto se saírem erradas.
   - Validar cada regra portada linha a linha contra a saída do Excel atual, usando os RN-xxx do catálogo como checklist de aceite (112 regras = 112 casos de teste candidatos, não é necessário testar todos com o mesmo rigor, mas as marcadas Crítico/Alto na seção 18 do documento técnico devem ter teste de paridade explícito).

3. **Fase 2 — Interface:**
   - Construir a UI (ver [prompt de design](03_prompt_claude_design.md)) já com progresso real, cancelamento e validação de schema desde o primeiro protótipo — são baratos de incluir agora e caros de adicionar depois.

4. **Fase 3 — Transição:**
   - Rodar em paralelo por 1-2 fechamentos, comparando resultado linha a linha.
   - Só então os itens de limpeza do Excel legado (nomes definidos, módulos de macro gravada) deixam de ser relevantes.
