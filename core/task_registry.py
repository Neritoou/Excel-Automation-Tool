"""
Registro centralizado de tareas con autodescubrimiento.

Escanea el paquete 'tasks/' e importa automáticamente cualquier
subclase de BaseTask, haciéndola disponible para la GUI sin
configuración adicional.
"""

from __future__ import annotations

import importlib
import pkgutil
from pathlib import Path
from .base_task import BaseTask


class TaskRegistry:
    _instance: TaskRegistry | None = None
    _tasks: dict[str, BaseTask]

    def __new__(cls) -> TaskRegistry:
        if cls._instance is None:
            instance = super().__new__(cls)
            instance._tasks = {}
            cls._instance = instance
        return cls._instance

    def discover(self, package_path: str = "tasks") -> None:
        """Importa todos los módulos en el paquete y registra las tareas."""
        pkg = importlib.import_module(package_path)

        if pkg.__file__ is None:
            raise RuntimeError(f"El paquete '{package_path}' no tiene __file__ definido")

        pkg_dir = Path(pkg.__file__).parent

        for _, module_name, _ in pkgutil.iter_modules([str(pkg_dir)]):
            if module_name.startswith("_"):
                continue
            full_name = f"{package_path}.{module_name}"
            mod = importlib.import_module(full_name)

            for attr_name in dir(mod):
                attr = getattr(mod, attr_name)
                if (
                    isinstance(attr, type)
                    and issubclass(attr, BaseTask)
                    and attr is not BaseTask
                    and attr.task_id
                ):
                    self._tasks[attr.task_id] = attr()

    def get_all(self) -> dict[str, BaseTask]:
        return dict(self._tasks)

    def get(self, task_id: str) -> BaseTask | None:
        return self._tasks.get(task_id)
