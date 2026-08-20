"""Ponte Python <-> JavaScript exposta ao frontend via pywebview.api.*"""

from __future__ import annotations

import json
from dataclasses import asdict

import webview

from app.fronts import (
    SheetInfo,
    apply_sheet_edits,
    build_output,
    delete_sheet_column,
    get_sheet_data,
    insert_sheet_column,
    list_sheets,
)


class Api:
    def pick_file(self) -> str | None:
        """Abre o diálogo nativo de seleção de arquivo. Devolve o caminho
        escolhido ou None se o usuário cancelou."""
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.OPEN,
            file_types=("Planilhas Excel (*.xlsx;*.xlsm)", "Todos os arquivos (*.*)"),
        )
        if not result:
            return None
        return result[0]

    def pick_folder(self) -> str | None:
        """Diálogo nativo de seleção de pasta — usado por fontes que leem
        vários arquivos de um diretório em vez de um arquivo único (ex.:
        Other Income, RN-030 na doc técnica: única fonte multi-arquivo do
        Quick Data original)."""
        result = webview.windows[0].create_file_dialog(webview.FileDialog.FOLDER)
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

    def get_sheet_data(self, path: str, sheet_name: str, start_row: int, max_rows: int) -> str:
        """Página de linhas de uma aba, para o visualizador de abas do app
        (menu lateral) mostrar o conteúdo real do Front importado, em vez
        de só ter gerado um arquivo no disco."""
        data = get_sheet_data(path, sheet_name, start_row, max_rows)
        return json.dumps(data)

    def apply_sheet_edits(
        self,
        path: str,
        sheet_name: str,
        cell_edits: list[dict],
        row_colors: list[dict],
        row_formats: list[dict] | None = None,
        column_widths: list[dict] | None = None,
        cell_formats: list[dict] | None = None,
    ) -> str:
        """Grava edições de célula, cor/fonte/altura de linha, largura de
        coluna e formatação por célula direto na aba real do Front
        (visualizador do menu lateral)."""
        apply_sheet_edits(path, sheet_name, cell_edits, row_colors, row_formats, column_widths, cell_formats)
        return json.dumps({"ok": True})

    def delete_sheet_column(self, path: str, sheet_name: str, col: int) -> str:
        """Remove uma coluna inteira da aba real do Front (todas as
        linhas da planilha, não só a página visível)."""
        delete_sheet_column(path, sheet_name, col)
        return json.dumps({"ok": True})

    def insert_sheet_column(self, path: str, sheet_name: str, col: int) -> str:
        """Insere uma coluna em branco na aba real do Front (todas as
        linhas da planilha)."""
        insert_sheet_column(path, sheet_name, col)
        return json.dumps({"ok": True})

    def pick_dest_file(self) -> str | None:
        """Diálogo de salvar — onde gravar/acrescentar os Fronts importados."""
        result = webview.windows[0].create_file_dialog(
            webview.FileDialog.SAVE,
            save_filename="BASE_QUICK_DATA.xlsx",
            file_types=("Planilha Excel (*.xlsx)",),
        )
        if not result:
            return None
        path = result if isinstance(result, str) else result[0]
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"
        return path
