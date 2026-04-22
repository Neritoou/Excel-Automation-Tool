"""
PLANTILLA PARA NUEVAS TAREAS (Pylance strict / Pyright standard)
────────────────────────────────────────────────────────────────

Instrucciones:
  1. Copiar este archivo a /tasks/mi_nueva_tarea.py
  2. Renombrar la clase y definir task_id, task_name, task_description, task_icon
  3. Implementar get_params(), validate(), _run()
  4. Ejecutar main.py → la tarea aparecerá automáticamente en la GUI

IMPORTANTE:
  - Los parámetros de validate() y _run() deben llamarse "params"
    (mismo nombre que en BaseTask) para cumplir con Pyright standard.
  - Al asignar cell.value, cell.font, cell.fill, cell.border añadir:
      # type: ignore[union-attr]
    ya que ws.cell() devuelve Cell | MergedCell según los stubs de openpyxl.
"""

from typing import Any
from openpyxl.worksheet.worksheet import Worksheet
from core.base_task import BaseTask, TaskParam, TaskResult
from core.excel_handler import ExcelHandler


class MiNuevaTarea(BaseTask):

    task_id: str = "mi_nueva_tarea"
    task_name: str = "Mi Nueva Tarea"
    task_description: str = "Descripción de lo que hace esta tarea."
    task_icon: str = "🔧"

    def get_params(self) -> list[TaskParam]:
        return [
            TaskParam("file_1", "Archivo Excel", "file", group="Entrada"),
            TaskParam("sheet_1", "Hoja", "sheet", depends_on="file_1", group="Entrada"),
        ]

    def validate(self, params: dict[str, Any]) -> tuple[bool, str]:
        if not params.get("file_1"):
            return False, "Selecciona un archivo"
        return True, ""

    def _run(self, params: dict[str, Any]) -> TaskResult:
        handler = ExcelHandler()
        file_path: str = str(params["file_1"])
        wb = handler.load(file_path)
        ws: Worksheet = wb[str(params["sheet_1"])]  # type: ignore[assignment]

        # Tu lógica aquí...

        wb.save(file_path)
        return TaskResult(
            success=True,
            message="Tarea completada exitosamente.",
            output_files=[file_path],
        )
