"""
Clase base abstracta para todas las tareas de automatización Excel.

Patrón: Template Method
- define el esqueleto del algoritmo en execute()
- las subclases implementan los pasos concretos

Patrón: Strategy
- cada tarea es una estrategia intercambiable
- la GUI y el engine las consumen de forma uniforme
"""

from abc import ABC, abstractmethod
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any


@dataclass
class TaskParam:
    """Descriptor de un parámetro que la tarea necesita del usuario."""
    name: str
    label: str
    param_type: str  # "file", "sheet", "column", "text", "number", "bool"
    required: bool = True
    default: Any = None
    depends_on: str | None = None  # p.ej. "sheet" depende de "file"
    group: str = "default"


@dataclass
class TaskResult:
    """Resultado estandarizado de cualquier tarea."""
    success: bool
    message: str
    output_files: list[str] = field(default_factory=list)
    details: dict[str, Any] = field(default_factory=dict)


class BaseTask(ABC):
    """
    Contrato que toda tarea debe cumplir.

    Para crear una nueva funcionalidad:
      1. Heredar de BaseTask
      2. Implementar los 4 métodos abstractos
      3. Colocar el archivo en /tasks/
      4. El registro la descubre automáticamente
    """

    # ── Metadatos (las subclases los sobreescriben) ──────────────────
    task_id: str = ""
    task_name: str = ""
    task_description: str = ""
    task_icon: str = "⚙️"

    @abstractmethod
    def get_params(self) -> list[TaskParam]:
        """Declara qué parámetros necesita la tarea (la GUI los renderiza)."""
        ...

    @abstractmethod
    def validate(self, params: dict[str, Any]) -> tuple[bool, str]:
        """Valida los parámetros antes de ejecutar."""
        ...

    @abstractmethod
    def _run(self, params: dict[str, Any]) -> TaskResult:
        """Lógica principal de la tarea (paso protegido del Template Method)."""
        ...

    # ── Template Method ──────────────────────────────────────────────
    def execute(self, params: dict[str, Any], progress_cb: Callable[[float, str], None] | None = None) -> TaskResult:
        """
        Algoritmo fijo: validar → notificar inicio → ejecutar → notificar fin.
        Las subclases NO sobreescriben este método.
        """
        ok, msg = self.validate(params)
        if not ok:
            return TaskResult(success=False, message=f"Validación fallida: {msg}")

        if progress_cb:
            progress_cb(0, f"Iniciando: {self.task_name}...")

        try:
            result = self._run(params)
        except Exception as e:
            result = TaskResult(success=False, message=f"Error inesperado: {e}")

        if progress_cb:
            progress_cb(100, result.message)

        return result
