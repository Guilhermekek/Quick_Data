"""Ponte Python <-> JavaScript exposta ao frontend via pywebview.api.*"""

from __future__ import annotations

import json
from dataclasses import asdict

import webview

from app.fronts import SheetInfo, build_output, list_sheets


class Api:
    def pick_file(self) -> str | None:
        """Abre o diálogo nativo de seleção de arquivo. Devolve o caminho
        escolhido ou None se o usuário cancelou."""
        result = webview.windows[0].create_file_dialog(
            webview.OPEN_DIALOG,
            file_types=("Planilhas Excel (*.xlsx;*.xlsm)", "Todos os arquivos (*.*)"),
        )
        if not result:
            return None
        return result[0]

    def list_sheets(self, path: str) -> str:
        """Devolve, em JSON, as abas do arquivo e quais parecem ser Fronts."""
        sheets: list[SheetInfo] = list_sheets(path)
        return json.dumps([asdict(s) for s in sheets])

    def build_output(self, source_path: str, items: list[dict], dest_path: str) -> str:
        """Monta dest_path a partir da lista ordenada `items` (Fronts
        copiados de source_path e/ou abas novas criadas na tela — ver
        app.fronts.build_output). Devolve JSON com os nomes finais das
        abas criadas (pode diferir do original em caso de colisão)."""
        created = build_output(source_path, items, dest_path)
        return json.dumps({"created": created})

    def pick_dest_file(self) -> str | None:
        """Diálogo de salvar — onde gravar/acrescentar os Fronts importados."""
        result = webview.windows[0].create_file_dialog(
            webview.SAVE_DIALOG,
            save_filename="BASE_QUICK_DATA.xlsx",
            file_types=("Planilha Excel (*.xlsx)",),
        )
        if not result:
            return None
        path = result if isinstance(result, str) else result[0]
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return path
