# Creado por Francisco Grandón Vergara
"""
Subpaquete para sanitización, compatibilidad, auditoría y manipulación de fórmulas de Excel.
"""

from .audit import (
    COM_EXCEL_ERROR_MAP,
    KNOWN_FORMULA_ERRORS,
    audit_com_formula_errors,
    audit_openpyxl_formula_errors,
    get_formula_error_name,
    is_formula_error_value,
)
from .compatibility import (
    make_absolute_reference,
    make_col_absolute_reference,
    make_relative_reference,
    make_row_absolute_reference,
    propagate_formula_row,
    replace_formula_placeholders,
)
from .sanitize import (
    MODERN_EXCEL_FUNCTIONS,
    sanitize_formula_compatibility,
    strip_xlfn_prefix,
)
from .structured import STRUCTURED_REF_REGEX, transpile_structured_formula

__all__ = [
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
    "STRUCTURED_REF_REGEX",
    "transpile_structured_formula",
]
