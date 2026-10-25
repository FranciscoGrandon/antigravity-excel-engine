# Creado por Francisco Grandón Vergara
"""
Helpers y validadores de entradas para operaciones en Excel (nombres de hojas, rangos, fórmulas, etc.).
"""

import os
import re
from typing import Any, Optional, Union

from ..models.errors import (
    FormulaValidationError,
    InvalidRangeError,
    WorkbookNotFoundError,
)
from .ranges import range_boundaries

# Caracteres no permitidos en nombres de hojas de Excel: \ / ? * : [ ]
_INVALID_SHEET_CHARS = set(r'\/?*:[')
_SHEET_MAX_LEN = 31

# Errores comunes de fórmulas en Excel
_EXCEL_FORMULA_ERRORS = {
    "#NULL!",
    "#DIV/0!",
    "#VALUE!",
    "#REF!",
    "#NAME?",
    "#NUM!",
    "#N/A",
    "#SPILL!",
    "#CALC!",
}


def validate_sheet_name(sheet_name: str) -> str:
    """
    Valida que un nombre de hoja cumpla con las restricciones de Microsoft Excel:
    - No vacío y menor o igual a 31 caracteres.
    - No contener los caracteres: \\ / ? * : [ ]
    - No empezar ni terminar con comilla simple (').
    Retorna el nombre sanitizado o lanza ValueError si es inválido.
    """
    if not isinstance(sheet_name, str):
        raise ValueError(f"Sheet name must be string, got {type(sheet_name).__name__}")

    name = sheet_name.strip()
    if not name:
        raise ValueError("Sheet name cannot be empty or whitespace only.")

    if len(name) > _SHEET_MAX_LEN:
        raise ValueError(f"Sheet name '{name}' exceeds Excel maximum limit of 31 characters ({len(name)} chars).")

    for ch in name:
        if ch in _INVALID_SHEET_CHARS:
            raise ValueError(f"Sheet name '{name}' contains forbidden character '{ch}'. Forbidden: \\ / ? * : [ ]")

    if name.startswith("'") or name.endswith("'"):
        raise ValueError(f"Sheet name '{name}' cannot begin or end with an apostrophe (').")

    return name


def sanitize_sheet_name(sheet_name: str, fallback: str = "Sheet1") -> str:
    """
    Sanitiza un nombre de hoja potencialmente inválido reemplazando caracteres prohibidos por guiones bajos
    y truncándolo a 31 caracteres de forma segura.
    """
    if not sheet_name or not isinstance(sheet_name, str):
        return fallback

    cleaned = "".join("_" if ch in _INVALID_SHEET_CHARS else ch for ch in sheet_name.strip())
    cleaned = cleaned.strip("'")
    if not cleaned:
        return fallback
    return cleaned[:_SHEET_MAX_LEN]


def validate_range_string(range_str: str) -> str:
    """
    Valida que una cadena de rango (ej. 'A1', 'B2:F20') tenga una sintaxis reconocible.
    Lanza InvalidRangeError si la sintaxis es incorrecta.
    """
    try:
        range_boundaries(range_str)
        return range_str.strip()
    except Exception as e:
        raise InvalidRangeError(f"Range validation failed for '{range_str}': {e}") from e


def validate_workbook_exists(file_path: str) -> str:
    """
    Verifica que el archivo de libro de Excel especificado exista en el sistema de archivos.
    Lanza WorkbookNotFoundError si el archivo no se encuentra.
    """
    if not file_path:
        raise WorkbookNotFoundError("Workbook path is empty or not specified.")

    norm_path = os.path.abspath(file_path)
    if not os.path.exists(norm_path):
        raise WorkbookNotFoundError(f"Workbook file does not exist: {norm_path}")
    if not os.path.isfile(norm_path):
        raise WorkbookNotFoundError(f"Specified path is not a file: {norm_path}")

    return norm_path


def validate_formula_text(formula: str) -> str:
    """
    Valida la sintaxis básica de una fórmula de Excel:
    - Debe comenzar con '=' o '@='
    - Balance de paréntesis
    Lanza FormulaValidationError si es inválida.
    """
    if not isinstance(formula, str):
        raise FormulaValidationError(f"Formula must be a string, got {type(formula).__name__}")

    f_str = formula.strip()
    if not f_str.startswith("=") and not f_str.startswith("@="):
        raise FormulaValidationError(f"Formula must begin with '=': '{f_str}'")

    # Chequeo rápido de balance de paréntesis
    open_count = 0
    in_quotes = False
    for ch in f_str:
        if ch == '"':
            in_quotes = not in_quotes
        elif not in_quotes:
            if ch == '(':
                open_count += 1
            elif ch == ')':
                open_count -= 1
                if open_count < 0:
                    raise FormulaValidationError(f"Unbalanced parentheses in formula: '{f_str}'")

    if open_count != 0:
        raise FormulaValidationError(f"Unbalanced parentheses (unclosed) in formula: '{f_str}'")

    return f_str


def is_formula_error(value: Any) -> bool:
    """Retorna True si un valor devuelto por una celda corresponde a un error de cálculo de Excel."""
    if not isinstance(value, str):
        return False
    return value.strip().upper() in _EXCEL_FORMULA_ERRORS


def is_file_locked(file_path: str) -> bool:
    """
    Determina si un archivo está bloqueado por otro proceso o por Microsoft Excel.
    En Windows intenta abrir con win32file.CreateFile solicitando GENERIC_WRITE,
    FILE_SHARE_READ y OPEN_EXISTING. Si falla por ERROR_SHARING_VIOLATION o PermissionError retorna True.
    """
    if not file_path or not os.path.exists(file_path):
        return False

    abs_path = os.path.abspath(file_path)

    # Intento nativo en Windows vía pywin32
    try:
        import pywintypes
        import win32con
        import win32file

        handle = win32file.CreateFile(
            abs_path,
            win32con.GENERIC_WRITE,
            win32con.FILE_SHARE_READ,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None,
        )
        win32file.CloseHandle(handle)
        return False
    except ImportError:
        pass
    except Exception:
        # Si ocurre violación de compartición o error de permisos
        return True

    # Fallback multiplataforma estándar (abrir en modo append/escritura)
    try:
        with open(abs_path, "a+b"):
            pass
        return False
    except (PermissionError, IOError, OSError):
        return True


def atomic_save_file(source_temp_path: str, target_path: str, backup: bool = False) -> str:
    """
    Reemplaza target_path de forma atómica con source_temp_path.
    Si backup es True y target_path existe, crea primero una copia de respaldo con sufijo .bak.
    Utiliza os.replace para garantizar atomicidad y prevenir corrupción de archivos.
    """
    import shutil

    target_abs = os.path.abspath(target_path)
    source_abs = os.path.abspath(source_temp_path)

    if not os.path.exists(source_abs):
        raise FileNotFoundError(f"El archivo temporal de origen no existe: {source_abs}")

    if backup and os.path.exists(target_abs):
        backup_path = f"{target_abs}.bak"
        shutil.copy2(target_abs, backup_path)

    os.replace(source_abs, target_abs)
    return target_abs
