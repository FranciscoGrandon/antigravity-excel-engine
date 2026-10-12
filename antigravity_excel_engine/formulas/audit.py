# Creado por Francisco Grandón Vergara
"""
Auditoría y reconocimiento de errores de cálculo en fórmulas de Excel.
"""

from typing import Any, Dict, List, Optional, Set, Union
from ..utils.ranges import range_boundaries

# Mapa de constantes de error de win32com (códigos numéricos HRESULT retornados por Excel en .Value)
COM_EXCEL_ERROR_MAP: Dict[int, str] = {
    -2146826281: "#DIV/0!",
    -2146826246: "#N/A",
    -2146826259: "#NAME?",
    -2146826288: "#NULL!",
    -2146826252: "#NUM!",
    -2146826265: "#REF!",
    -2146826273: "#VALUE!",
}

# Conjunto de cadenas de errores textuales canónicos de Excel
KNOWN_FORMULA_ERRORS: Set[str] = {
    "#VALUE!",
    "#REF!",
    "#DIV/0!",
    "#N/A",
    "#NAME?",
    "#NUM!",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
}


def is_formula_error_value(val: Any) -> bool:
    """
    Determina si un valor (int en win32com o str en openpyxl/csv) corresponde
    a un error de evaluación de fórmula en Excel.
    """
    if val is None:
        return False
    if isinstance(val, int) and val in COM_EXCEL_ERROR_MAP:
        return True
    if isinstance(val, str):
        cleaned = val.strip().upper()
        return cleaned in KNOWN_FORMULA_ERRORS
    return False


def get_formula_error_name(val: Any) -> Optional[str]:
    """Retorna el nombre canónico del error ('#REF!', '#DIV/0!', etc.) o None si no es error."""
    if isinstance(val, int) and val in COM_EXCEL_ERROR_MAP:
        return COM_EXCEL_ERROR_MAP[val]
    if isinstance(val, str):
        cleaned = val.strip().upper()
        if cleaned in KNOWN_FORMULA_ERRORS:
            return cleaned
    return None


def audit_openpyxl_formula_errors(ws: Any, range_str: str) -> List[Dict[str, Any]]:
    """
    Inspecciona un rango en una hoja de trabajo de openpyxl en búsqueda de errores de fórmulas.
    Retorna una lista de diccionarios con información detallada de cada celda con error.
    """
    min_col, min_row, max_col, max_row = range_boundaries(range_str)
    errors: List[Dict[str, Any]] = []

    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            err_name = get_formula_error_name(cell.value)
            if err_name:
                errors.append({
                    "cell": cell.coordinate,
                    "error": err_name,
                    "formula": str(cell.value)
                })
    return errors


def audit_com_formula_errors(ws: Any, range_str: str) -> List[Dict[str, Any]]:
    """
    Inspecciona un rango en una hoja activa de Excel COM en búsqueda de errores de fórmulas.
    """
    rng = ws.Range(range_str)
    errors: List[Dict[str, Any]] = []

    for cell in rng:
        val = cell.Value
        err_name = get_formula_error_name(val)
        if err_name:
            try:
                addr = cell.Address(False, False)
            except Exception:
                addr = str(getattr(cell, "Address", ""))
            try:
                formula_str = cell.Formula
            except Exception:
                formula_str = ""

            errors.append({
                "cell": addr,
                "error": err_name,
                "formula": formula_str
            })
    return errors
