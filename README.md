# Quick Data

Projeto de trabalho com foco em replicar a ferramenta da empresa TIM, Quick Data — hoje uma planilha Excel/VBA (Quick Data 3.23 versao de referencia usada ) usada pelo time de Planning & Control para extração, consolidação e reporte financeiro. O objetivo é substituir essa planilha por um app desktop, mantendo o fluxo de trabalho que o time já conhece.

## Stack

Construído usando Python + React.

- **Back**: Python, com [pywebview](https://pywebview.flowrl.com/) como shell desktop (a janela do app é um WebView nativo, sem depender de instalar navegador ou runtime externo).
- **Front**: interface web (HTML/CSS/JS, com componentes em React) renderizada dentro do pywebview.

## Design

Foi usado Claude Design com o foco em ajudar na criação de um design novo que correspondesse ao Quick Data atual — o layout, as telas e os componentes visuais foram desenhados a partir de prints da planilha original, mantendo a identidade visual da TIM.

## Estrutura do projeto

```
app/
  ui/          → Front (interface)
  api.py       → ponte Python ↔ JavaScript (pywebview.api)
  fronts.py    → lógica de detecção/importação de Fronts e leitura de abas
  Back.py      → backend real (em andamento)
  main.py      → inicialização do app (janela pywebview)
```

Só o conteúdo de `app/` (Front + Back) fica versionado no GitHub. Documentação, material de referência e arquivos do projeto original ficam fora do repositório.

## Como rodar

Requer Python 3.11 ou 3.12.

```bash
py -3.12 -m venv .venv312
.venv312/Scripts/pip install -r requirements.txt
.venv312/Scripts/python.exe -m app --debug
```

## Status

Em desenvolvimento. Já funcionando:

- Importação de Fronts (abas de relatório) de arquivos Excel externos, com visualização do conteúdo dentro do próprio app.
- Criação de abas customizadas no menu lateral, com grade editável (colunas e linhas).
- Modal de parâmetros de extração por fonte de dados (KPI, mês, arquivo/pasta), validado contra o comportamento da planilha original.
- Tela de Main Results com filtros e grade editáveis.

Em andamento: backend real de extração (`app/Back.py`).
