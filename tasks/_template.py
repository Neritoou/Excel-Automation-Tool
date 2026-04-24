from typing import Any

import xlwings as xw

from core.base_task import BaseTask, ParamType, TaskParam, TaskResult
from core.exceptions import ValidationTaskError
from core.excel_handler import ExcelHandler

class MiNuevaTarea(BaseTask):

    task_id: str = "mi_nueva_tarea"
    task_name: str = "Mi Nueva Tarea"
    task_description: str = "Descripción de lo que hace esta tarea."
    task_icon: str = "🔧"

    def get_params(self) -> list[TaskParam]:
        """Declara los parámetros que necesita la tarea."""
        return [
            TaskParam("file_1", "Archivo Excel", ParamType.FILE, group="Entrada"),
            TaskParam("sheet_1", "Hoja", ParamType.SHEET, depends_on="file_1", group="Entrada"),
        ]

    def validate(self, params: dict[str, Any]) -> None:
        """Valida que el archivo esté seleccionado."""
        if not params.get("file_1"):
            raise ValidationTaskError("Selecciona un archivo")

    def _run(self, params: dict[str, Any]) -> TaskResult:
        """Lógica principal de la tarea."""
        file_path = str(params["file_1"])
        wb = ExcelHandler.load(file_path)
        ws: xw.Sheet = wb.sheets[str(params["sheet_1"])]

        # LÓGICA DE LA TAREA. . .
        # ExcelHandler.set_cell(ws, row, col, value="dato", fill=(198, 239, 206), bold=True)

        return TaskResult(
            success=True,
            message="Tarea completada. Ctrl+Z para deshacer, Ctrl+S para guardar.",
            output_files=[file_path],
        )