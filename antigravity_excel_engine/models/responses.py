# Creado por Francisco Grandón Vergara
"""
Estructuras de respuesta y estados operativos de Antigravity Excel Engine.
"""

from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, Optional


class OperationStatus(str, Enum):
    """Estados estandarizados de las operaciones del motor."""
    SUCCESS = "success"
    ERROR = "error"
    WARNING = "warning"
    PENDING = "pending"
    BUSY = "busy"


@dataclass
class BaseResponse:
    """Estructura base de respuesta de cualquier operación o comando."""
    status: OperationStatus = OperationStatus.SUCCESS
    message: Optional[str] = None
    execution_time_ms: Optional[float] = None
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        data: Dict[str, Any] = {
            "status": self.status.value if isinstance(self.status, OperationStatus) else str(self.status)
        }
        if self.message is not None:
            data["message"] = self.message
        if self.execution_time_ms is not None:
            data["execution_time_ms"] = self.execution_time_ms
        if self.metadata:
            data["metadata"] = self.metadata
        return data


@dataclass
class SuccessResponse(BaseResponse):
    """Respuesta exitosa devuelta por el motor o el Daemon IPC."""
    status: OperationStatus = OperationStatus.SUCCESS
    result: Any = None

    def to_dict(self) -> Dict[str, Any]:
        base = super().to_dict()
        base["result"] = self.result
        return base


@dataclass
class ErrorResponse(BaseResponse):
    """Respuesta de error devuelta por el motor o el Daemon IPC."""
    status: OperationStatus = OperationStatus.ERROR
    error: str = ""
    error_type: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        base = super().to_dict()
        base["error"] = self.error
        if self.error_type:
            base["error_type"] = self.error_type
        if self.details:
            base["details"] = self.details
        return base
