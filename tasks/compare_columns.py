"""
Tarea: Comparar columnas entre dos archivos Excel.

Compara cada valor de la columna del Excel 1 contra todos los valores
de la columna del Excel 2 (y viceversa). El modo de comparación determina
cómo se normalizan los valores y qué colores se aplican.
"""

from dataclasses import dataclass
from os.path import abspath
from typing import Any

import xlwings as xw

from core.base_task import BaseTask, ParamType, TaskParam, TaskResult
from core.exceptions import ValidationTaskError
from core.excel_handler import ExcelHandler, CellStyle, CellValue


# ── Modos de comparación ─────────────────────────────────────────

@dataclass(frozen=True)
class CompareMode:
    """Modo base. Normaliza a minúsculas sin espacios, convierte floats enteros a int."""

    label: str
    match_style: CellStyle
    no_match_style: CellStyle

    @staticmethod
    def _to_clean_str(value: Any) -> str:
        """Convierte floats enteros (444.0) a int antes de hacer string."""
        if isinstance(value, float) and value.is_integer():
            return str(int(value))
        return str(value)

    def normalize(self, value: Any) -> str:
        """Normalización estándar: strip + lowercase."""
        if value is None:
            return ""
        return self._to_clean_str(value).strip().lower()


class ReferenciaMode(CompareMode):
    """Quita ceros a la izquierda. Todo ceros se obvia (string vacío)."""

    def normalize(self, value: Any) -> str:
        if value is None:
            return ""
        raw = self._to_clean_str(value).strip()
        stripped = raw.lstrip("0")
        return stripped.lower() if stripped else ""


class MontoMode(CompareMode):
    """Compara como float para ignorar diferencias de formato (444.50 vs 444.5)."""

    def normalize(self, value: Any) -> str:
        if value is None:
            return ""
        try:
            return str(float(value))
        except (ValueError, TypeError):
            return str(value).strip().lower()


# Agregar un modo = una entrada aquí (+ clase si necesita normalize especial)
MODES: dict[str, CompareMode] = {
    "Referencia": ReferenciaMode(
        label="Referencia",
        match_style=CellStyle(fill=(198, 239, 206)),
        no_match_style=CellStyle(fill=(255, 199, 206)),
    ),
    "Monto": MontoMode(
        label="Monto",
        match_style=CellStyle(fill=(219, 234, 254)),
        no_match_style=CellStyle(fill=(255, 199, 206)),
    ),
}


# ── Tarea ────────────────────────────────────────────────────────

class CompareColumnsTask(BaseTask):

    task_id = "compare_columns"
    task_name = "Comparar Columnas"
    task_description = (
        "Compara los valores de una columna del Excel 1 contra todos los "
        "valores de la columna del Excel 2. Colorea la columna indicada "
        "según si el valor existe o no en el otro archivo."
    )
    task_icon = "🔍"

    _REQUIRED = (
        "file_1", "sheet_1", "column_1", "color_col_1",
        "file_2", "sheet_2", "column_2", "color_col_2",
        "mode",
    )

    def get_params(self) -> list[TaskParam]:
        """Declara los parámetros: modo, dos archivos con columna clave y columna de color."""
        return [
            TaskParam("mode", "Modo de comparación", ParamType.SELECT,
                      options=tuple(MODES.keys()), group="Opciones"),

            TaskParam("file_1", "Archivo Excel 1", ParamType.FILE, group="Excel 1"),
            TaskParam("sheet_1", "Hoja", ParamType.SHEET, depends_on="file_1", group="Excel 1"),
            TaskParam("column_1", "Columna clave", ParamType.COLUMN, depends_on="sheet_1", group="Excel 1"),
            TaskParam("color_col_1", "Columna a colorear", ParamType.COLUMN, depends_on="sheet_1", group="Excel 1"),

            TaskParam("file_2", "Archivo Excel 2", ParamType.FILE, group="Excel 2"),
            TaskParam("sheet_2", "Hoja", ParamType.SHEET, depends_on="file_2", group="Excel 2"),
            TaskParam("column_2", "Columna clave", ParamType.COLUMN, depends_on="sheet_2", group="Excel 2"),
            TaskParam("color_col_2", "Columna a colorear", ParamType.COLUMN, depends_on="sheet_2", group="Excel 2"),
        ]

    def validate(self, params: dict[str, Any]) -> None:
        """Verifica parámetros requeridos, archivos distintos y modo válido."""
        for key in self._REQUIRED:
            if not params.get(key):
                raise ValidationTaskError(f"Falta el parámetro: {key}")

        if abspath(str(params["file_1"])) == abspath(str(params["file_2"])):
            raise ValidationTaskError("Los dos archivos deben ser diferentes")

        if params["mode"] not in MODES:
            raise ValidationTaskError(f"Modo no válido: {params['mode']}")

    def _run(self, params: dict[str, Any]) -> TaskResult:
        """
        Compara cada valor de una columna contra todos los valores de la otra.

        1. Lee ambas columnas y construye sets normalizados
        2. Clasifica cada fila como match o diff
        3. Aplica colores en batch (una llamada COM por color por archivo)
        """
        mode = MODES[str(params["mode"])]

        wb1 = ExcelHandler.load(str(params["file_1"]))
        wb2 = ExcelHandler.load(str(params["file_2"]))

        ws1: xw.Sheet = wb1.sheets[str(params["sheet_1"])]
        ws2: xw.Sheet = wb2.sheets[str(params["sheet_2"])]

        vals1 = ExcelHandler.read_column_values(ws1, int(params["column_1"]))
        vals2 = ExcelHandler.read_column_values(ws2, int(params["column_2"]))

        set1: set[str] = {mode.normalize(v.value) for v in vals1}
        set2: set[str] = {mode.normalize(v.value) for v in vals2}

        color_col_1 = int(params["color_col_1"])
        color_col_2 = int(params["color_col_2"])

        m1, d1 = self._classify_and_color(ws1, vals1, set2, color_col_1, mode)
        m2, d2 = self._classify_and_color(ws2, vals2, set1, color_col_2, mode)

        return TaskResult(
            success=True,
            message=(
                f"Comparación [{mode.label}] completada.\n"
                f"  Excel 1 → {m1} coincidencias, {d1} diferencias\n"
                f"  Excel 2 → {m2} coincidencias, {d2} diferencias\n"
                f"  Ctrl+Z para deshacer, Ctrl+S para guardar"
            ),
            output_files=[str(params["file_1"]), str(params["file_2"])],
            details={
                "mode": mode.label,
                "excel_1": {"matches": m1, "diffs": d1},
                "excel_2": {"matches": m2, "diffs": d2},
            },
        )

    @staticmethod
    def _classify_and_color(
        ws: xw.Sheet,
        values: list[CellValue],
        reference_set: set[str],
        color_col: int,
        mode: CompareMode,
    ) -> tuple[int, int]:
        """
        Clasifica cada valor como match/diff y aplica colores en batch.

        En vez de llamar set_cell por cada celda (lento), agrupa las filas
        por resultado y aplica el color en dos llamadas COM.
        Retorna (matches, diffs).
        """
        match_rows: list[int] = []
        diff_rows: list[int] = []

        for cv in values:
            if mode.normalize(cv.value) in reference_set:
                match_rows.append(cv.row)
            else:
                diff_rows.append(cv.row)

        if mode.match_style.fill:
            ExcelHandler.color_rows(ws, color_col, match_rows, mode.match_style.fill)
        if mode.no_match_style.fill:
            ExcelHandler.color_rows(ws, color_col, diff_rows, mode.no_match_style.fill)

        return len(match_rows), len(diff_rows)