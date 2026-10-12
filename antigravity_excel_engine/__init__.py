# Creado por Francisco Grandón Vergara
"""
Paquete principal de Antigravity Excel Engine.
Motor modular de automatización, análisis y curación inteligente de Excel.
"""

__version__ = "1.0.0"
__author__ = "Francisco Grandón Vergara"

from .models import (
    # Errors
    ExcelEngineError,
    ExcelLockError,
    WorkbookNotFoundError,
    FormulaValidationError,
    PipeConnectionError,
    UnsupportedOperationError,
    InvalidRangeError,
    ExcelBusyError,
    # Requests
    DaemonRequest,
    PipeCommandRequest,
    CellWriteRequest,
    CuratePipelineRequest,
    # Responses
    OperationStatus,
    BaseResponse,
    SuccessResponse,
    ErrorResponse,
    # Results
    CellWriteResult,
    AuditReport,
    CurateResult,
    AggregateResult,
    PivotResult,
    OutlierResult,
)

from .utils import (
    # Logging
    LOGGER_NAME,
    get_logger,
    log_operation,
    log_exception,
    log_duration,
    # Ranges
    coordinate_from_string,
    column_index_from_string,
    col_to_index,
    get_column_letter,
    index_to_col,
    range_boundaries,
    parse_range,
    chunk_cell_addresses,
    # Dates
    EXCEL_BASE_DATE,
    EXCEL_BASE_DATETIME,
    parse_excel_date_serial,
    coerce_date_to_excel_serial,
    excel_serial_to_datetime,
    excel_serial_to_date,
    # Validation
    validate_sheet_name,
    sanitize_sheet_name,
    validate_range_string,
    validate_workbook_exists,
    validate_formula_text,
    is_formula_error,
)

from .formatting import (
    FinancialStyleGuide,
    hex_to_rgb_int,
    hex_to_argb_hex,
    apply_number_format,
    apply_borders_openpyxl,
    apply_style_preset_openpyxl,
    apply_style_preset_com,
    apply_style_preset,
)

from .formulas import (
    MODERN_EXCEL_FUNCTIONS,
    sanitize_formula_compatibility,
    strip_xlfn_prefix,
    COM_EXCEL_ERROR_MAP,
    KNOWN_FORMULA_ERRORS,
    is_formula_error_value,
    get_formula_error_name,
    audit_openpyxl_formula_errors,
    audit_com_formula_errors,
    propagate_formula_row,
    replace_formula_placeholders,
    make_absolute_reference,
    make_relative_reference,
    make_row_absolute_reference,
    make_col_absolute_reference,
)

from .backends import (
    BaseExcelBackend,
    LiveExcelCOMBackend,
    HeadlessOpenPyXLBackend,
    retry_on_excel_busy,
    excel_fast_mode,
    get_com_clean_address,
)

from .engine import AntigravityExcelEngine, ExcelEngine

from .operations import (
    read_ranges,
    get_range_as_csv,
    write_cell_range,
    write_cells_declarative,
    copy_paste_range,
    audit_data_quality,
    check_formula_errors,
    detect_and_flag_outliers,
    aggregate,
    add_calculated_column,
    remove_duplicates,
    clean_table_dataset,
    curate_pipeline,
    create_pivot_table,
    create_chart,
    create_summary_table,
)

__all__ = [
    "__version__",
    "__author__",
    # Engine principal
    "AntigravityExcelEngine",
    "ExcelEngine",
    # Backends
    "BaseExcelBackend",
    "LiveExcelCOMBackend",
    "HeadlessOpenPyXLBackend",
    "retry_on_excel_busy",
    "excel_fast_mode",
    "get_com_clean_address",
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
    # Utilities
    "LOGGER_NAME",
    "get_logger",
    "log_operation",
    "log_exception",
    "log_duration",
    "coordinate_from_string",
    "column_index_from_string",
    "col_to_index",
    "get_column_letter",
    "index_to_col",
    "range_boundaries",
    "parse_range",
    "chunk_cell_addresses",
    "EXCEL_BASE_DATE",
    "EXCEL_BASE_DATETIME",
    "parse_excel_date_serial",
    "coerce_date_to_excel_serial",
    "excel_serial_to_datetime",
    "excel_serial_to_date",
    "validate_sheet_name",
    "sanitize_sheet_name",
    "validate_range_string",
    "validate_workbook_exists",
    "validate_formula_text",
    "is_formula_error",
    # Formatting
    "FinancialStyleGuide",
    "hex_to_rgb_int",
    "hex_to_argb_hex",
    "apply_number_format",
    "apply_borders_openpyxl",
    "apply_style_preset_openpyxl",
    "apply_style_preset_com",
    "apply_style_preset",
    # Formulas
    "MODERN_EXCEL_FUNCTIONS",
    "sanitize_formula_compatibility",
    "strip_xlfn_prefix",
    "COM_EXCEL_ERROR_MAP",
    "KNOWN_FORMULA_ERRORS",
    "is_formula_error_value",
    "get_formula_error_name",
    "audit_openpyxl_formula_errors",
    "audit_com_formula_errors",
    "propagate_formula_row",
    "replace_formula_placeholders",
    "make_absolute_reference",
    "make_relative_reference",
    "make_row_absolute_reference",
    "make_col_absolute_reference",
    # Operations
    "read_ranges",
    "get_range_as_csv",
    "write_cell_range",
    "write_cells_declarative",
    "copy_paste_range",
    "audit_data_quality",
    "check_formula_errors",
    "detect_and_flag_outliers",
    "aggregate",
    "add_calculated_column",
    "remove_duplicates",
    "clean_table_dataset",
    "curate_pipeline",
    "create_pivot_table",
    "create_chart",
    "create_summary_table",
]
