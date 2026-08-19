"""
Detecção e importação de "Fronts" (abas de relatório/layout) de arquivos
Quick Data externos — equivalente ao Form_Importacao do VBA legado
(RN-101 a RN-104 na documentação técnica).

Abas de infraestrutura (Base, tabelas mestre, config, log) são excluídas
da lista por padrão; o restante é oferecido ao usuário para importar.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import openpyxl
from openpyxl.utils import get_column_letter

# Nomes de aba que fazem parte da infraestrutura do Quick Data, não são
# "Fronts" de relatório. Comparação sem diferenciar maiúscula/minúscula.
SUPPORT_SHEET_NAMES = {
    "front >>", "control panel >>", "support >>", "aux_extracoes >>",
    "extracao", "ajustes", "base",
    "ref_cruzada_1", "ref_cruzada_2", "sup_linhas", "aux_ifrs16",
    "valid_lin", "cc bd", "dropcomb", "visao_it", "aux",
    "tk_lista_de_erros", "rgm_1", "rgm_2", "fixed_1", "fixed_2",
    "mockup_rgm", "dp_segmento", "listdefinednames", "lista_arq_aux",
    "painel_dm", "dp_rateio",
}


@dataclass
class SheetInfo:
    name: str
    is_front: bool
    rows: int
    cols: int


def _is_front(sheet_name: str) -> bool:
    key = sheet_name.strip().lower()
    if key in SUPPORT_SHEET_NAMES:
        return False
    if key.endswith(">>"):
        return False
    return True


def list_sheets(path: str) -> list[SheetInfo]:
    """Abre o arquivo (somente leitura) e devolve todas as abas, marcando
    quais parecem ser Fronts (candidatas a importação) vs. infraestrutura."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        result = []
        for name in wb.sheetnames:
            ws = wb[name]
            result.append(
                SheetInfo(
                    name=name,
                    is_front=_is_front(name),
                    rows=ws.max_row or 0,
                    cols=ws.max_column or 0,
                )
            )
        return result
    finally:
        wb.close()


def _unique_sheet_name(wb: openpyxl.Workbook, desired: str) -> str:
    """Replica a regra do Form_Importacao original: nome duplicado ganha
    sufixo numérico incremental em vez de travar a importação (RN-103)."""
    if desired not in wb.sheetnames:
        return desired
    n = 2
    while f"{desired} ({n})" in wb.sheetnames:
        n += 1
    return f"{desired} ({n})"


def import_fronts(source_path: str, sheet_names: list[str], dest_path: str) -> list[str]:
    """Copia os Fronts selecionados de `source_path` para `dest_path`
    (criado do zero se não existir). Só valores + largura de coluna +
    formato numérico são preservados nesta primeira versão — fórmulas
    não são copiadas (mesma simplificação que a exportação de Front do
    VBA original já fazia: converte fórmula em valor antes de exportar).

    Retorna a lista de nomes finais das abas criadas no destino.
    """
    src = openpyxl.load_workbook(source_path, read_only=True, data_only=True)
    dest_file = Path(dest_path)
    if dest_file.exists():
        dest = openpyxl.load_workbook(dest_path)
    else:
        dest = openpyxl.Workbook()
        # remove a aba padrão em branco que o openpyxl cria sozinho
        default = dest.active
        dest.remove(default)

    created: list[str] = []
    try:
        for name in sheet_names:
            if name not in src.sheetnames:
                continue
            src_ws = src[name]
            final_name = _unique_sheet_name(dest, name)
            dst_ws = dest.create_sheet(title=final_name[:31])  # limite do Excel

            for row in src_ws.iter_rows():
                for cell in row:
                    if cell.value is None:
                        continue
                    dst_ws.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.number_format:
                        dst_ws.cell(row=cell.row, column=cell.column).number_format = cell.number_format

            for col_idx, dim in getattr(src_ws, "column_dimensions", {}).items():
                if dim.width:
                    dst_ws.column_dimensions[col_idx].width = dim.width

            created.append(final_name)

        dest.save(dest_path)
    finally:
        src.close()
        dest.close()

    return created
