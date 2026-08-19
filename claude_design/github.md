repo: Guilhermekek/Quick_Data
branch: main

## Last sync
date: 2026-08-19T14:02:18Z

### Updated in this project
- Quick Data reconstruído como app desktop navegável: 11 telas em `Quick Data.dc.html`, tema claro por padrão.
- Login por matrícula com perfis Administrador e Usuário; ações restritas aparecem com cadeado.
- Configurações em abas: Geral, Conexões, Arquivos auxiliares, Notificações (ADM) e Usuários e permissões (ADM).
- Pop-ups de conclusão e de falha com som, progresso real com cancelamento, Manual em vídeo e tela de Atualizações.

## Pending commit
message: claude design primeira versao
files:
- Quick Data.dc.html
- support.js
- github.md
note: exportado do projeto; o commit precisa ser feito manualmente (acesso somente leitura ao repo)

## Screen map
| Tela no projeto | Origem |
| --- | --- |
| Shell (sidebar, footer, tokens, tema, densidade) | Tim_Trabalho: pipeline_runner/frontend/src/styles.css, App.jsx |
| Botões, badges, modais, progresso, toasts | Tim_Trabalho: pipeline_runner/frontend/src/components/* |
| Configurações (abas, settings-group, settings-row) | Tim_Trabalho: pipeline_runner/frontend/src/screens/Settings.jsx |
| Notificações (canal do Teams, destinatários) | layout de Settings.jsx; canal e escolha por pessoa são novos |
| Relatar bug ou melhoria | Tim_Trabalho: pipeline_runner/frontend/src/screens/BugReport.jsx |
| Manual (vídeo, capítulos) e Atualizações | Tim_Trabalho: pipeline_runner/frontend/src/screens/Manual.jsx |
| Login e perfis de acesso | novo — sem equivalente no Pipeline Runner |
| Painel Principal, Extração, Import/Export, Ajustes, Main Results | prints da planilha Quick Data 3.23 (uploads/) |
| Configurações · Arquivos auxiliares | print da aba LISTA_ARQ_AUX |

## Sync history
- 2026-08-14T18:24:31Z — leitura de Settings.jsx e BugReport.jsx para Notificações e melhorias.
- 2026-08-14T13:54:02Z — Manual em vídeo, Atualizações e fila de melhorias.
- 2026-08-14T13:25:44Z — abas de Configurações e permissões.
- 2026-08-14T00:00:00Z — leitura inicial do design system do Pipeline Runner.
