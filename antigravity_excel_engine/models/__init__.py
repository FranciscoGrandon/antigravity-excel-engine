# Creado por Francisco Grandón Vergara
"""
Módulo de modelos, tipos y excepciones de Antigravity Excel Engine.
"""

from .errors import (
    ExcelBusyError,
    ExcelEngineError,
    ExcelLockError,
    FormulaValidationError,
    InvalidRangeError,
    PipeConnectionError,
    UnsupportedOperationError,
    WorkbookNotFoundError,
)
from .requests import (
    CellWriteRequest,
    CuratePipelineRequest,
    DaemonRequest,
    PipeCommandRequest,
)
from .responses import (
    BaseResponse,
    ErrorResponse,
    OperationStatus,
    SuccessResponse,
)
from .result_types import (
    AggregateResult,
    AuditReport,
    CellWriteResult,
    CurateResult,
    OutlierResult,
    PivotResult,
)

__all__ = [
    # Errors
    "ExcelEngineError",
    "ExcelLockError",
    "WorkbookNotFoundError",
    "FormulaValidationError",
    "PipeConnectionError",
    "UnsupportedOperationError",
    "InvalidRangeError",
    "ExcelBusyError",
    # Requests
    "DaemonRequest",
    "PipeCommandRequest",
    "CellWriteRequest",
    "CuratePipelineRequest",
    # Responses
    "OperationStatus",
    "BaseResponse",
    "SuccessResponse",
    "ErrorResponse",
    # Results
    "CellWriteResult",
    "AuditReport",
    "CurateResult",
    "AggregateResult",
    "PivotResult",
    "OutlierResult",
]
