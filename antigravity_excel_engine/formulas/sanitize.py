# Creado por Francisco Grandón Vergara
"""
Sanitización y normalización de compatibilidad XML para fórmulas de Excel.
"""

import re
from typing import Optional

# Catálogo ampliado de funciones modernas de Excel que requieren prefijo XML _xlfn. en openpyxl
MODERN_EXCEL_FUNCTIONS = (
    # Funciones texto y lógica post-2007/2016
    "TEXTJOIN",
    "CONCAT",
    "IFS",
    "SWITCH",
    "MAXIFS",
    "MINIFS",
    # Matrices dinámicas y fórmulas modernas de Excel 365/2021
    "XLOOKUP",
    "XMATCH",
    "UNIQUE",
    "FILTER",
    "SORT",
    "SORTBY",
    "SEQUENCE",
    "RANDARRAY",
    "LET",
    "LAMBDA",
    "BYROW",
    "BYCOL",
    "SCAN",
    "REDUCE",
    "MAP",
    "MAKEARRAY",
    "CHOOSECOLS",
    "CHOOSEROWS",
    "DROP",
    "TAKE",
    "EXPAND",
    "HSTACK",
    "VSTACK",
    "TOROW",
    "TOCOL",
    "WRAPROWS",
    "WRAPCOLS",
    "TEXTSPLIT",
    "TEXTBEFORE",
    "TEXTAFTER",
)

_XLFN_PATTERN = re.compile(
    r'(?<!_xlfn\.)\b(' + '|'.join(MODERN_EXCEL_FUNCTIONS) + r')\b(?=\s*\()',
    flags=re.IGNORECASE
)


def sanitize_formula_compatibility(formula: str) -> str:
    """
    Detecta y normaliza nombres de funciones Excel modernas (post-2007 y dinámicas de Excel 365)
    que requieren el prefijo de compatibilidad XML '_xlfn.' en bibliotecas como openpyxl.
    Funciones tratadas:
      - TEXTJOIN, CONCAT, IFS, SWITCH, MAXIFS, MINIFS.
      - Fórmulas de matriz dinámica: XLOOKUP, XMATCH, UNIQUE, FILTER, SORT, SORTBY, SEQUENCE, RANDARRAY.
      - Lambdas y nuevas utilidades: LET, LAMBDA, VSTACK, HSTACK, TEXTSPLIT, etc.
    Si ya traen el prefijo, se preservan sin duplicarlo.
    """
    if not formula or not isinstance(formula, str):
        return formula

    def _replace_match(m):
        func_name = m.group(1).upper()
        return f"_xlfn.{func_name}"

    return _XLFN_PATTERN.sub(_replace_match, formula)


def strip_xlfn_prefix(formula: str) -> str:
    """Elimina los prefijos de compatibilidad '_xlfn.' para presentación limpia al usuario o interfaces COM."""
    if not formula or not isinstance(formula, str):
        return formula
    return re.sub(r'_xlfn\.', '', formula, flags=re.IGNORECASE)
