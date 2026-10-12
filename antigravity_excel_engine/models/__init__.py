# Creado por Francisco Grandón Vergara
"""
Módulo de modelos, tipos y excepciones de Antigravity Excel Engine.
"""

from .errors import (
    ExcelEngineError,
    ExcelLockError,
    WorkbookNotFoundError,
    FormulaValidationError,
    PipeConnectionError,
    UnsupportedOperationError,
    InvalidRangeError,
    ExcelBusyError,
)

from .requests import (
    DaemonRequest,
    PipeCommandRequest,
    CellWriteRequest,
    CuratePipelineRequest,
)

from .responses import (
    OperationStatus,
    BaseResponse,
    SuccessResponse,
    ErrorResponse,
)

from .result_types import (
    CellWriteResult,
    AuditReport,
    CurateResult,
    AggregateResult,
    PivotResult,
    OutlierResult,
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
