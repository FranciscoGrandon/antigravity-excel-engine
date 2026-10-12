# Creado por Francisco Grandón Vergara
"""
Operaciones de agregación estadística y columnas calculadas.
Delegación al backend activo sin lógica de negocio propia.
"""

from typing import Any, Dict, List, Optional

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.aggregate")


def aggregate(
    backend,
    range_str: str = "",
    group_by_col: str = "",
    metric_col: str = "",
    agg_func: str = "sum",
    top_n: int = 10,
    sheet: Optional[str] = None,
    dest_cell: Optional[str] = None,
    dest_sheet: Optional[str] = None,
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Agrupa filas por una columna y calcula una métrica agregada.

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango de datos con encabezados (ej. "A1:F500").
        group_by_col: Nombre de columna para agrupar.
        metric_col: Nombre de columna para agregar.
        agg_func: Función de agregación ("sum", "avg", "count", "min", "max").
        top_n: Número de grupos a retornar en el resultado.
        sheet: Nombre de hoja origen; None usa la activa.
        dest_cell: Celda de inicio donde escribir el resultado (opcional).
        dest_sheet: Hoja donde escribir el resultado (opcional).

    Returns:
        Diccionario con los datos agregados y metadatos.
    """
    logger.debug(
        "aggregate | range=%s group_by=%s metric=%s func=%s top_n=%d",
        range_str, group_by_col, metric_col, agg_func, top_n
    )
    # Soporta tanto llamadas por posición (group_by_col, metric_col) como con range_str
    # Si group_by_col está vacío pero range_str tiene el nombre de columna:
    if not metric_col and group_by_col and range_str:
        metric_col = group_by_col
        group_by_col = range_str
        range_str = ""

    call_kwargs: Dict[str, Any] = {
        "group_by_col": group_by_col,
        "metric_col": metric_col,
        "agg_func": agg_func,
        "top_n": top_n,
    }
    if sheet is not None:
        call_kwargs["sheet"] = sheet
    return backend.aggregate(**call_kwargs)


def add_calculated_column(
    backend,
    header: str = "",
    formula_template: str = "",
    sheet: Optional[str] = None,
    col_format: Optional[str] = None,
    autofit: bool = True,
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Agrega una columna calculada con fórmula dinámica.

    Args:
        backend: Instancia de BaseExcelBackend.
        header: Encabezado de la nueva columna (o 'new_col_name').
        formula_template: Plantilla de fórmula base fila 2 o con {r} (ej. "=L2/J2").
        sheet: Nombre de hoja; None usa la activa.
        col_format: Formato numérico opcional (o 'number_format').
        autofit: Si True, autoajusta el ancho de la columna.

    Returns:
        Diccionario con la columna destino, fila de inicio/fin y fórmulas aplicadas.
    """
    # Compatibilidad con múltiples nombres de argumentos
    resolved_header = header or kwargs.get("new_col_name") or kwargs.get("header") or kwargs.get("range_str", "")
    resolved_formula = formula_template or kwargs.get("formula") or kwargs.get("formula_template", "")
    resolved_format = col_format or kwargs.get("number_format") or kwargs.get("col_format")
    resolved_sheet = sheet if sheet is not None else kwargs.get("sheet")
    resolved_autofit = autofit if "autofit" not in kwargs else kwargs.get("autofit", True)

    logger.debug(
        "add_calculated_column | header=%s formula=%s format=%s",
        resolved_header, resolved_formula, resolved_format
    )

    call_kwargs: Dict[str, Any] = {
        "header": resolved_header,
        "formula_template": resolved_formula,
        "autofit": resolved_autofit,
    }
    if resolved_format is not None:
        call_kwargs["number_format"] = resolved_format
    if resolved_sheet is not None:
        call_kwargs["sheet"] = resolved_sheet
    return backend.add_calculated_column(**call_kwargs)

