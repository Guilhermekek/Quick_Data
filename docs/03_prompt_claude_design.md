# Prompt para claude.ai/design — Novo Quick Data

> Copie o texto do bloco abaixo e cole no claude.ai/design. Cores e stack já atualizadas
> com base no projeto FIT (`Tim_Trabalho/pipeline_runner`) — o Quick Data novo será
> construído do mesmo jeito (executável Python + React via pywebview), então o design
> precisa ser implementável com esses mesmos componentes/tokens, não só bonito.

---

## Prompt

```
Contexto do projeto:
Estou reconstruindo uma ferramenta financeira interna corporativa chamada "Quick Data",
usada pelo time de Planejamento & Controle (P&C) da TIM Brasil para consolidar e reportar
dados financeiros/contábeis. A versão atual é uma planilha Excel com macros VBA, lenta e
com interface datada. A nova versão será um EXECUTÁVEL DESKTOP STANDALONE (Python), com
motor de dados totalmente novo por trás — mas a interface precisa ser visualmente
RECONHECÍVEL para quem já usa a ferramenta atual, para reduzir a curva de aprendizado de
um time que já está acostumado com o layout existente. Ou seja: modernizar sem alienar.

Quero mockups/wireframes de alta fidelidade das telas abaixo, com um guia de estilo
(cores, tipografia, componentes) consistente entre elas.

STACK TÉCNICA (o design precisa ser implementável nisso, não genérico)
- Frontend em React 18 + Vite, rodando dentro de uma janela desktop nativa via
  pywebview (WebView2 do Windows) — mesmo padrão do nosso outro app interno (TIM
  Pipeline Runner / projeto FIT). Ou seja: pense em componentes web (cards, tabelas,
  modais, sidebar) com CSS via variáveis/tokens, não em elementos nativos de SO.
- Já existe um design system parcialmente pronto desse app irmão (sidebar fixa à
  esquerda em cor de marca, footer fino, tema claro/escuro via toggle, densidade
  compacta/regular/confortável). Aproveite a MESMA linguagem de tokens para os dois
  apps ficarem consistentes entre si — mas o tema padrão do Quick Data deve ser
  **CLARO** (parecido com planilha/relatório financeiro), diferente do Pipeline Runner
  que é escuro por padrão.

IDENTIDADE VISUAL
- Empresa: TIM Brasil (telecom). Cores de marca já definidas no projeto irmão:
  - Azul primário: `#004691` (hover `#0057b3`, ativo/pressed `#003a78`) — usar em
    sidebar, cabeçalhos, botões primários de navegação.
  - Vermelho de destaque: `#d10a11` — usar com moderação, só para o emblema/wordmark
    da marca e talvez ações de alerta, não como cor de botão genérica.
  - Cores de status (mesma paleta do app irmão, reaproveitar): sucesso `#16a34a`,
    atenção `#d97706`, erro `#dc2626`, neutro `#64748b`, informativo `#2563eb`.
  - Tema claro: fundo `#f4f6fa`, superfícies `#ffffff`/`#f8fafc`, texto `#0f172a`,
    bordas `#e2e8f0`.
- Fonte: Segoe UI (nativa do Windows, mesma do app irmão) — boa legibilidade em
  tabelas densas de números.
- Estética geral: ferramenta financeira corporativa séria, mas com visual "app de
  desktop moderno" (pense em algo entre Power BI Desktop e um ERP moderno) — não deve
  parecer uma planilha Excel crua, mas deve manter a MESMA LÓGICA DE NAVEGAÇÃO E OS
  MESMOS NOMES DE BOTÕES da versão atual, listados abaixo.
- Botões de ação: hoje são retângulos de cantos arredondados, coloridos, com texto
  branco em negrito, organizados em "cartões" por fonte de dados. Mantenha esse padrão
  visual (cartão colorido com ação), mas com acabamento mais moderno (sombra sutil,
  hover state, ícones) — usar os tokens de raio (`--radius: 8px`, cards `--radius-lg:
  12px`) e sombras sutis (`--shadow-1`, `--shadow-2`) do app irmão.

TELA 1 — Painel Principal / Home
Uma tela inicial estilo dashboard com cards de navegação para as áreas principais:
Extração de Dados, Ajustes Manuais, Relatórios/Main Results, Configurações. Deve
funcionar como "para onde eu sempre volto" — o app atual sempre retorna a esta tela
depois de qualquer operação.

TELA 2 — Painel de Extração (a tela mais usada do sistema)
Um grid/lista de "cartões", um por fonte de dados, cada cartão com:
- Nome da fonte (ex.: HUBBLE, BASE 1009, RGM, MOCKUP RGM, FIXED REVENUES, OTHER INCOME,
  BASE CONSOLIDADA, QUICK DATA, AJUSTES)
- Um indicador de status (ex.: "última extração: há 2h" / nunca extraído / com erro)
- Três botões por cartão: "Extrair/Importar", "Informações" (abre detalhes da config
  daquela fonte) e "Limpar base" (remove os dados já extraídos dessa fonte)
Acima do grid, uma barra de ações globais com os botões:
"Extrair todas as bases", "Atualizar bases auxiliares", "Limpar todas as bases
extraídas", "Exportar BASE e/ou Front(s)", "Importar BASE e/ou Front(s)", "Gerar base
pré-closing", "Aplicar CDC por referência".
Ao clicar em extrair, mostrar uma BARRA DE PROGRESSO REAL (não indeterminada) com
percentual, etapa atual (ex.: "Conectando ao SQL Server...", "Processando linha 4.200
de 12.000...") e um botão CANCELAR — isso é uma melhoria importante em relação à
versão atual, que não tem progresso real nem cancelamento.

TELA 3 — Importar/Exportar Front ou Base
Modal ou painel com duas listas lado a lado ("Disponível" à esquerda, "Selecionado" à
direita) e setas de mover itens entre elas (padrão dual-listbox clássico) — mantenha
esse padrão pois os usuários já conhecem, mas modernize o visual (cards em vez de
lista simples, drag-and-drop opcional). Abaixo, opção de exportar como "arquivo único"
ou "um arquivo por item selecionado". Botão de confirmação com resumo antes de
executar (quantos itens, para onde).

TELA 4 — Grade de Ajustes Manuais
Uma tabela editável estilo planilha (grid com linhas/colunas, mas com validação de
dropdown nas colunas de dimensão: Empresa, Classe, Centro de Custo, Diretoria N1-N3,
Abertura) e colunas de meses (Jan-Dez) editáveis numericamente. Deve parecer uma
tabela de dados profissional (tipo AG Grid / Excel-like grid), não um formulário.

TELA 5 — Relatório / Main Results
Uma visão de relatório com filtros no topo (chips ou dropdowns para Empresa, Abertura,
Linha_BD, IFRS Contábil — equivalente aos slicers da versão atual) e uma tabela
dinâmica/pivot abaixo, com totais e subtotais destacados visualmente.

REQUISITOS DE INTERAÇÃO A MELHORAR (vs. versão atual)
- Toda operação longa precisa de barra de progresso real + botão cancelar (hoje só
  tem texto na barra de status, sem cancelamento).
- Mensagens de erro específicas e acionáveis, nunca genéricas tipo "ocorreu um erro".
- Confirmações de ação destrutiva (ex.: sobrescrever arquivo, limpar base) com resumo
  claro do que será afetado antes de confirmar.
- Consistência de interação entre telas semelhantes (ex.: mesmo padrão de
  clique/duplo-clique em todas as listas do sistema).

ENTREGÁVEIS
Gere wireframes/mockups de alta fidelidade para as 5 telas acima, mais um guia de
estilo resumido (paleta de cores com hex, tipografia, especificação dos componentes
de botão/card/tabela/barra de progresso) para garantir consistência visual entre
todas as telas do aplicativo.
```

---

### Notas para você (fora do prompt)

- Os nomes de botões e o layout de cartões da Tela 2 vieram diretamente dos shapes reais extraídos do arquivo `.xlsb` (não são suposição) — isso deve garantir que quem já usa a ferramenta reconheça a tela na hora.
- Se o claude.ai/design pedir upload de imagem de referência, um print da aba "Extração" do arquivo atual (aberto no Excel) ajudaria bastante — é a tela com mais identidade visual a preservar. Um print do Pipeline Runner (FIT) também ajuda, para calibrar o quanto o resultado deve puxar para o "app irmão" vs. para a planilha original.
- Cores confirmadas a partir de `frontend/src/styles.css` do projeto `Tim_Trabalho/pipeline_runner` (não são mais suposição).

### Como isso vira executável (mesmo caminho do FIT)

Quando o design estiver pronto, a implementação segue o padrão já validado no `pipeline_runner`:
- **Frontend:** React 18 + Vite, build estático (`npm run build`), sem CDN — offline-first.
- **Shell desktop:** `pywebview`, janela nativa via WebView2, carregando o bundle estático em produção e o dev server do Vite (hot reload) durante o desenvolvimento.
- **Empacotamento:** PyInstaller `--onedir`.
- **Restrição de ambiente:** Python 3.11 ou 3.12 — `pythonnet` (dependência transitiva do pywebview no Windows) ainda não tem wheels para 3.13/3.14.
- Estrutura de pastas espelhando `pipeline_runner/` (`app/` com o launcher pywebview, `frontend/` com o React, `scripts/build.bat` e `scripts/dev.bat`) — reduz a curva de manutenção porque quem mexe em um app já sabe mexer no outro.
