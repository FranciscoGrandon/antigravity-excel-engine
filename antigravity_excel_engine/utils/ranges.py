# Creado por Francisco Grandón Vergara
"""
Utilidades de coordenadas, celdas y rangos compatibles con Excel (1-based index).
Optimizadas con LRU cache y sin dependencias pesadas de terceros.
"""

import functools
import re
from typing import List, Optional, Tuple, Union

from ..models.errors import InvalidRangeError

_COORD_RE = re.compile(r"^[$]?([A-Za-z]{1,3})[$]?(\d+)$")
_ABSOLUTE_RE = re.compile(
    r"^[$]?(?P<min_col>[A-Za-z]{1,3})?[$]?(?P<min_row>\d+)?(?::[$]?(?P<max_col>[A-Za-z]{1,3})?[$]?(?P<max_row>\d+)?)?$",
    re.IGNORECASE,
)


def coordinate_from_string(coord_string: str) -> Tuple[str, int]:
    """
    Convierte una coordenada como 'B12' o '$B$12' en una tupla ('B', 12).
    Lanza InvalidRangeError si la coordenada no es válida o la fila es 0.
    """
    if not isinstance(coord_string, str):
        raise InvalidRangeError(f"Coordinate must be a string, got {type(coord_string).__name__}")
    m = _COORD_RE.match(coord_string.strip())
    if not m:
        raise InvalidRangeError(f"Invalid cell coordinates ({coord_string})")
    col, row = m.groups()
    r = int(row)
    if not r:
        raise InvalidRangeError(f"There is no row 0 ({coord_string})")
    return col.upper(), r


@functools.lru_cache(maxsize=1024)
def column_index_from_string(col: str) -> int:
    """
    Convierte un nombre de columna alfanumérico (ej. 'A', 'Z', 'AA', 'XFD') en un índice decimal 1-based.
    Rango permitido: 1 a 18278 (hasta 'ZZZ').
    """
    col_clean = str(col).strip().upper()
    if not col_clean or len(col_clean) > 3:
        raise InvalidRangeError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
    idx = 0
    for char in col_clean:
        if not ("A" <= char <= "Z"):
            raise InvalidRangeError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    if not (1 <= idx <= 18278):
        raise InvalidRangeError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
    return idx


# Alias comunes
col_to_index = column_index_from_string


@functools.lru_cache(maxsize=1024)
def get_column_letter(col_idx: int) -> str:
    """
    Convierte un índice de columna 1-based en su letra correspondiente ('A', 'AA', etc.).
    Lanza InvalidRangeError si el índice está fuera del rango [1, 18278].
    """
    try:
        col_idx = int(col_idx)
    except (ValueError, TypeError):
        raise InvalidRangeError(f"Invalid column index: {col_idx}")

    if not (1 <= col_idx <= 18278):
        raise InvalidRangeError(f"Invalid column index {col_idx}")
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


# Alias comunes
index_to_col = get_column_letter


def range_boundaries(range_string: str) -> Tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """
    Convierte una cadena de rango (ej. 'A1:B10', 'C5', 'A:B', '1:10')
    en una tupla de enteros (min_col, min_row, max_col, max_row).
    """
    if not isinstance(range_string, str):
        raise InvalidRangeError(f"Range must be a string, got {type(range_string).__name__}")
    m = _ABSOLUTE_RE.match(range_string.strip())
    if not m:
        raise InvalidRangeError(f"{range_string} is not a valid coordinate or range")
    min_col, min_row, max_col, max_row = m.group("min_col"), m.group("min_row"), m.group("max_col"), m.group("max_row")
    c_min = column_index_from_string(min_col) if min_col else None
    r_min = int(min_row) if min_row else None
    c_max = column_index_from_string(max_col) if max_col else c_min
    r_max = int(max_row) if max_row else r_min
    return c_min, r_min, c_max, r_max


parse_range = range_boundaries


def chunk_cell_addresses(cell_list: List[str], max_chars: int = 200) -> List[str]:
    """
    Agrupa direcciones de celdas ('A1', 'B2', ...) en cadenas unificadas separadas por comas
    de longitud segura para llamadas multi-rango en COM (evita el límite de ~255 caracteres de Range).
    """
    chunks = []
    current: List[str] = []
    current_len = 0
    for c in cell_list:
        c_str = str(c).strip()
        if not c_str:
            continue
        c_len = len(c_str) + (1 if current else 0)
        if current and (current_len + c_len > max_chars):
            chunks.append(",".join(current))
            current = [c_str]
            current_len = len(c_str)
        else:
            current.append(c_str)
            current_len += c_len
    if current:
        chunks.append(",".join(current))
    return chunks
