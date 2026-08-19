"""
Quick Data (novo) — launcher pywebview.

Dois modos:
  - dev:  carrega o servidor do Vite em http://localhost:5173 (hot reload).
          `python -m app --dev` (com `cd frontend && npm run dev` rodando
          em outro terminal).
  - prod: carrega o bundle estático de `app/ui/index.html`.
          `python -m app` (rodar `scripts\\build.bat` antes).
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import webview

from app.api import Api

APP_TITLE = "Quick Data"
APP_VERSION = "0.1.0"

DEV_URL = "http://localhost:5173"


def _ui_index() -> Path:
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)  # type: ignore[attr-defined]
    else:
        base = Path(__file__).resolve().parent
    return base / "ui" / "index.html"


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="quick-data")
    p.add_argument("--dev", action="store_true", help="Carrega o servidor de dev do Vite.")
    p.add_argument("--debug", action="store_true", help="Abre com o devtools do WebView habilitado.")
    return p.parse_args()


def run() -> None:
    args = _parse_args()

    if args.dev:
        url = DEV_URL
        print(f"[quick-data] modo dev — carregando {url}")
    else:
        index = _ui_index()
        if not index.exists():
            print(
                f"[quick-data] bundle da UI não encontrado em {index}.\n"
                f"  Rode primeiro: scripts\\build.bat\n"
                f"  Ou use o modo dev: python -m app --dev",
                file=sys.stderr,
            )
            sys.exit(1)
        url = index.as_uri()
        print(f"[quick-data] modo prod — carregando {url}")

    api = Api()

    webview.create_window(
        title=APP_TITLE,
        url=url,
        js_api=api,
        width=1360,
        height=840,
        min_size=(1100, 680),
        background_color="#F4F6FA",
        text_select=True,
    )
    webview.start(debug=args.debug)


if __name__ == "__main__":
    run()
