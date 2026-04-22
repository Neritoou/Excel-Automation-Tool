"""
Capa de abstracción para operaciones comunes de lectura/escritura Excel.

Encapsula openpyxl para que las tareas no dependan directamente
de la librería y se puedan testear/mockear fácilmente.
"""

from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.utils import get_column_letter


class ExcelHandler:

    @staticmethod
    def load(path: str) -> Workbook:
        return load_workbook(path)

    @staticmethod
    def get_sheet_names(path: str) -> list[str]:
        wb = load_workbook(path, read_only=True)
        names = wb.sheetnames
        wb.close()
        return names

    @staticmethod
    def get_column_headers(path: str, sheet_name: str) -> list[str]:
        """Devuelve los encabezados (fila 1) de la hoja indicada."""
        wb = load_workbook(path, read_only=True)
        ws = wb[sheet_name]
        headers = []
        for cell in next(ws.iter_rows(min_row=1, max_row=1)):
            headers.append(str(cell.value) if cell.value is not None else "")
        wb.close()
        return headers

    @staticmethod
    def read_column_values(
        ws: Worksheet, col_idx: int, start_row: int = 2
    ) -> list[tuple[int, Any]]:
        """Lee los valores de una columna, devolviendo (fila, valor)."""
        max_row = ws.max_row if ws.max_row is not None else 0
        values: list[tuple[int, Any]] = []
        for row in range(start_row, max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value is not None:
                values.append((row, cell.value))
        return values

    @staticmethod
    def column_letter(idx: int) -> str:
        return get_column_letter(idx)
