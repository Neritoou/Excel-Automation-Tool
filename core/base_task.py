from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from enum import Enum
from typing import Any

from core.exceptions import ValidationTaskError

class ParamType(Enum):
    """Tipos de parámetro soportados por el constructor dinámico de formularios."""
    FILE = "file"
    SHEET = "sheet"
    COLUMN = "column"
    TEXT = "text"
    NUMBER = "number"
    BOOL = "bool"
    SELECT = "select"


@dataclass(frozen=True, slots=True)
class TaskParam:
    """Descriptor inmutable de un parámetro que la tarea necesita del usuario."""
    name: str
    label: str
    param_type: ParamType
    required: bool = True
    default: Any = None
    depends_on: str | None = None
    group: str = "default"
    options: tuple[str, ...] = ()


@dataclass(slots=True)
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
      3. Colocar el archivo en /tasks/ e importar en el __init__
    """
    task_id: str = ""
    task_name: str = ""
    task_description: str = ""
    task_icon: str = "⚙️"

    @abstractmethod
    def get_params(self) -> list[TaskParam]:
        """Declara qué parámetros necesita la tarea (la GUI los renderiza)."""
        ...

    @abstractmethod
    def validate(self, params: dict[str, Any]) -> None:
        """Valida los parámetros antes de ejecutar. Lanza ValidationError si falla."""
        ...

    @abstractmethod
    def _run(self, params: dict[str, Any]) -> TaskResult:
        """Lógica principal de la tarea (paso protegido del Template Method)."""
        ...

    def execute(self, params: dict[str, Any]) -> TaskResult:
        """Las subclases NO sobreescriben este método."""
        try:
            self.validate(params)
            return self._run(params)
        except ValidationTaskError as e:
            return TaskResult(success=False, message=f"Validación fallida: {e}")
        except Exception as e:
            return TaskResult(success=False, message=f"Error inesperado: {e}")
        
    # -- HELPERS ---

    @staticmethod
    def normalize(value: Any) -> str:
        """Normaliza un valor a string en minúsculas sin espacios laterales."""
        if value is None:
            return ""
        return str(value).strip().lower()