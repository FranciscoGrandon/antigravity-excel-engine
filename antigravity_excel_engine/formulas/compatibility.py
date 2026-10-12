# Creado por Francisco Grandón Vergara
"""
Utilidades de propagación, compatibilidad de referencias y reemplazo de placeholders en fórmulas.
"""

import re
from typing import Optional


def propagate_formula_row(formula_template: str, row_idx: int, base_row: int = 2) -> str:
    """
    Propaga una plantilla de fórmula para una fila destino `row_idx`, respetando referencias
    absolutas y relativas.
    - Soporta placeholders explícitos `{r}` y `{row}`.
    - Si la fórmula contiene referencias relativas a la fila base (por defecto fila 2, ej. `A2`, `$B2`, `Sheet!C2`),
      sustituye dinámicamente el número de fila base por `row_idx`.
    - Preserva estrictamente intactas las referencias absolutas en fila (ej. `$C$4`, `A$2`, `$A$2`, `'Params'!$C$4`).
    """
    if not formula_template or not isinstance(formula_template, str):
        return formula_template

    if "{r}" in formula_template:
        return formula_template.replace("{r}", str(row_idx))
    if "{row}" in formula_template:
        return formula_template.replace("{row}", str(row_idx))

    # Identificar celdas relativas en la fila base_row:
    # 1. Columna con opcional '$' inicial y 1-3 letras: (\$?[A-Za-z]{1,3})
    # 2. Número de fila base_row exacto sin '$' entre columna y número
    # 3. No seguido de dígitos adicionales ni de paréntesis abierto (evitar funciones como LOG2)
    pattern = rf"(?<![\$A-Za-z0-9_])(\$?[A-Za-z]{{1,3}}){base_row}(?!\d|\s*\()"
    return re.sub(pattern, rf"\g<1>{row_idx}", formula_template)


def replace_formula_placeholders(
    formula: str,
    row: Optional[int] = None,
    col: Optional[str] = None,
    table_name: Optional[str] = None,
    extra_vars: Optional[dict] = None
) -> str:
    """
    Reemplaza placeholders dinámicos comunes en plantillas de fórmulas.
    Soporta:
      - {r} o {row} -> row
      - {c} o {col} -> col
      - {table} -> table_name
      - Variables personalizadas pasadas en extra_vars (ej. {tax_rate} -> 0.19)
    """
    if not formula or not isinstance(formula, str):
        return formula

    res = formula
    if row is not None:
        res = res.replace("{r}", str(row)).replace("{row}", str(row))
    if col is not None:
        res = res.replace("{c}", str(col)).replace("{col}", str(col))
    if table_name is not None:
        res = res.replace("{table}", str(table_name))

    if extra_vars:
        for k, v in extra_vars.items():
            placeholder = f"{{{k}}}"
            if placeholder in res:
                res = res.replace(placeholder, str(v))

    return res


def make_absolute_reference(col_letter: str, row_idx: int) -> str:
    """Retorna una referencia absoluta de celda (ej. '$C$4')."""
    return f"${col_letter.upper()}${row_idx}"


def make_relative_reference(col_letter: str, row_idx: int) -> str:
    """Retorna una referencia relativa estándar de celda (ej. 'C4')."""
    return f"{col_letter.upper()}{row_idx}"


def make_row_absolute_reference(col_letter: str, row_idx: int) -> str:
    """Retorna una referencia con fila fija y columna relativa (ej. 'C$4')."""
    return f"{col_letter.upper()}${row_idx}"


def make_col_absolute_reference(col_letter: str, row_idx: int) -> str:
    """Retorna una referencia con columna fija y fila relativa (ej. '$C4')."""
    return f"${col_letter.upper()}{row_idx}"
