# Creado por Francisco Grandón Vergara
"""
Módulo para transpilación de fórmulas de referencias estructuradas de tablas Excel.
Traduce sintaxis de tablas tipo '=[@Ventas] - [@Costos]' a coordenadas relativas como '=B2 - C2',
soporta sintaxis formal 'Tabla[[#This Row],[Col]]', agregaciones de columna completa '=SUM([Ventas])',
y preserva literales de texto entre comillas dobles.
"""

from __future__ import annotations

import re
from typing import Dict, Optional, Tuple

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("formulas.structured")

# Regex para referencias estructuradas en la fila actual:
# Soporta:
# - [@Columna]
# - [@[Columna Con Espacio]]
# - Tabla[@Columna]
# - Tabla[@[Columna Con Espacio]]
# - [[#This Row],[Columna]]
# - [[#This Row], [Columna con espacios]]
# - Tabla[[#This Row],[Columna]]
STRUCTURED_REF_REGEX = re.compile(
    r"(?:[A-Za-z0-9_\u00C0-\u00FF\.]+\s*)?(?:\[@(?:\[([^\]]+)\]|([^\]]+))\]|\[\s*\[#(?:This\s+Row|this\s+row)\]\s*,\s*(?:\[([^\]]+)\]|([^\]]+))\s*\])"
)

# Regex para referencias a columnas completas (agregación sin @ ni #):
# Soporta:
# - [Columna]
# - Tabla[Columna]
# - Tabla[[Columna]]
# - [[Columna]]
STRUCTURED_FULL_COL_REGEX = re.compile(
    r"(?:[A-Za-z0-9_\u00C0-\u00FF\.]+\s*)?\[(?:\[([^\]]+)\]|([^\]@#]+))\]"
)


def transpile_structured_formula(
    formula: str,
    headers_map: Dict[str, str],
    row_idx: int,
    table_range: Optional[str] = None,
    row_bounds: Optional[Tuple[int, int]] = None,
) -> str:
    """
    Toma fórmulas con referencias estructuradas de tablas y las traduce
    a coordenadas relativas de celda para la fila dada o rangos de columnas completas.

    Ejemplos:
        >>> headers = {"Ventas": "B", "Costos": "C", "Precio Unitario": "D"}
        >>> transpile_structured_formula("=[@Ventas] - [@Costos]", headers, 2)
        '=B2 - C2'
        >>> transpile_structured_formula("=[@[Precio Unitario]] * 1.19", headers, 5)
        '=D5 * 1.19'
        >>> transpile_structured_formula("=SUM([Ventas])", headers, 2, row_bounds=(2, 10))
        '=SUM(B$2:B$10)'
        >>> transpile_structured_formula('="Total [@Ventas]: " & [@Ventas]', headers, 2)
        '="Total [@Ventas]: " & B2'

    Args:
        formula: Expresión o fórmula con referencias estructuradas.
        headers_map: Diccionario { "NombreColumna": "LetraColumna" } (ej. {"Ventas": "B"}).
        row_idx: Número de fila en la hoja de Excel (ej. 2, 3...).
        table_range: Rango de la tabla opcional (ej. "A2:E50") para inferir límites verticales.
        row_bounds: Tupla opcional (start_row, end_row) para rangos de agregación de columnas.

    Returns:
        Fórmula transpilada con coordenadas relativas de celda y rangos.
    """
    if not formula or not isinstance(formula, str):
        return formula

    # Mapa insensible a mayúsculas/minúsculas y espacios en blanco
    headers_clean = {k.strip(): v for k, v in headers_map.items()}
    headers_lower = {k.strip().lower(): v for k, v in headers_map.items()}

    # Resolver límites verticales para agregación de columnas completas
    start_row: Optional[int] = None
    end_row: Optional[int] = None
    if row_bounds and len(row_bounds) == 2:
        try:
            start_row, end_row = int(row_bounds[0]), int(row_bounds[1])
        except (ValueError, TypeError):
            pass
    elif table_range and isinstance(table_range, str):
        nums = re.findall(r"\d+", table_range)
        if len(nums) >= 2:
            try:
                start_row, end_row = int(nums[0]), int(nums[1])
            except (ValueError, TypeError):
                pass

    # 1. Separar literales de texto entre comillas dobles para protegerlos
    parts = re.split(r'("(?:""|[^"])*")', formula)

    for i in range(0, len(parts), 2):
        segment = parts[i]
        if not segment:
            continue

        # 2. Reemplazar sintaxis de fila actual ([@Col] o [[#This Row],[Col]])
        def _replace_this_row(match: re.Match) -> str:
            col_name = match.group(1) or match.group(2) or match.group(3) or match.group(4)
            if not col_name:
                return match.group(0)

            clean_col = col_name.strip()
            col_letter: Optional[str] = headers_clean.get(clean_col)
            if col_letter is None:
                col_letter = headers_lower.get(clean_col.lower())

            if col_letter:
                return f"{col_letter}{row_idx}"

            logger.warning(
                "Columna estructurada '%s' no encontrada en headers_map: %s",
                col_name, list(headers_map.keys())
            )
            return match.group(0)

        segment = STRUCTURED_REF_REGEX.sub(_replace_this_row, segment)

        # 3. Reemplazar sintaxis de columna completa (ej. [Ventas] o Tabla[Ventas])
        def _replace_full_col(match: re.Match) -> str:
            col_name = match.group(1) or match.group(2)
            if not col_name:
                return match.group(0)

            clean_col = col_name.strip()
            col_letter: Optional[str] = headers_clean.get(clean_col)
            if col_letter is None:
                col_letter = headers_lower.get(clean_col.lower())

            if col_letter:
                if start_row is not None and end_row is not None:
                    return f"{col_letter}${start_row}:{col_letter}${end_row}"
                return f"{col_letter}:{col_letter}"

            return match.group(0)

        segment = STRUCTURED_FULL_COL_REGEX.sub(_replace_full_col, segment)
        parts[i] = segment

    return "".join(parts)
