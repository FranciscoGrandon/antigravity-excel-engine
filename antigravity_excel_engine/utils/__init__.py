# Creado por Francisco Grandón Vergara
"""
Utilidades auxiliares (logging, rangos, fechas, validación) para Antigravity Excel Engine.
"""

from .logging import (
    LOGGER_NAME,
    get_logger,
    log_operation,
    log_exception,
    log_duration,
)

from .ranges import (
    coordinate_from_string,
    column_index_from_string,
    col_to_index,
    get_column_letter,
    index_to_col,
    range_boundaries,
    parse_range,
    chunk_cell_addresses,
)

from .dates import (
    EXCEL_BASE_DATE,
    EXCEL_BASE_DATETIME,
    parse_excel_date_serial,
    coerce_date_to_excel_serial,
    excel_serial_to_datetime,
    excel_serial_to_date,
)

from .validation import (
    validate_sheet_name,
    sanitize_sheet_name,
    validate_range_string,
    validate_workbook_exists,
    validate_formula_text,
    is_formula_error,
)

__all__ = [
    # Logging
    "LOGGER_NAME",
    "get_logger",
    "log_operation",
    "log_exception",
    "log_duration",
    # Ranges
    "coordinate_from_string",
    "column_index_from_string",
    "col_to_index",
    "get_column_letter",
    "index_to_col",
    "range_boundaries",
    "parse_range",
    "chunk_cell_addresses",
    # Dates
    "EXCEL_BASE_DATE",
    "EXCEL_BASE_DATETIME",
    "parse_excel_date_serial",
    "coerce_date_to_excel_serial",
    "excel_serial_to_datetime",
    "excel_serial_to_date",
    # Validation
    "validate_sheet_name",
    "sanitize_sheet_name",
    "validate_range_string",
    "validate_workbook_exists",
    "validate_formula_text",
    "is_formula_error",
]
