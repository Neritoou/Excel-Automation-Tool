class ExcelAutomatorError(Exception):
    """Excepción base del proyecto. Todas las demás heredan de esta."""
    pass

class ValidationTaskError(ExcelAutomatorError):
    """Error lanzado cuando los parámetros de una tarea no son válidos."""
    def __init__(self, message: str, field: str | None = None):
        super().__init__(message)
        self.field = field