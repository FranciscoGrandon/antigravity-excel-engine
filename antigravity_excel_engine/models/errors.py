# Creado por Francisco Grandón Vergara
"""
Definiciones de excepciones y errores específicos para Antigravity Excel Engine.
"""

from typing import Optional, Dict, Any


class ExcelEngineError(Exception):
    """Excepción base para todos los errores de Antigravity Excel Engine."""

    def __init__(self, message: str, details: Optional[Dict[str, Any]] = None):
        super().__init__(message)
        self.message = message
        self.details = details or {}

    def to_dict(self) -> Dict[str, Any]:
        """Serializa la excepción a un diccionario descriptivo."""
        return {
            "error_type": self.__class__.__name__,
            "message": self.message,
            "details": self.details,
        }


class ExcelLockError(ExcelEngineError):
    """Lanzada cuando un libro o recurso COM de Excel está bloqueado o retenido por otro proceso."""
    pass


class WorkbookNotFoundError(ExcelEngineError):
    """Lanzada cuando el archivo de libro de Excel especificado no existe o no se puede encontrar."""
    pass


class FormulaValidationError(ExcelEngineError):
    """Lanzada cuando una fórmula de Excel no supera la validación sintáctica o genera un error intrínseco (#REF!, #DIV/0!, etc.)."""
    pass


class PipeConnectionError(ExcelEngineError):
    """Lanzada cuando la conexión IPC vía Windows Named Pipe falla o se interrumpe."""
    pass


class UnsupportedOperationError(ExcelEngineError):
    """Lanzada cuando una operación solicitada no está soportada por el backend actual (Live COM vs Headless)."""
    pass


class InvalidRangeError(ExcelEngineError):
    """Lanzada cuando una coordenada o cadena de rango (A1, A1:B10, etc.) tiene un formato inválido o fuera de límites."""
    pass


class ExcelBusyError(ExcelEngineError):
    """Lanzada cuando la aplicación Excel COM se encuentra ocupada (RPC_E_CALL_REJECTED / RPC_E_SERVERCALL_RETRYLATER)."""
    pass
