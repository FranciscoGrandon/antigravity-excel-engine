# Creado por Francisco Grandón Vergara
"""
Operaciones de visualización: tablas dinámicas, gráficos y tablas de resumen.
Delegación al backend activo sin lógica de negocio propia.
"""

from typing import Any, Dict, List, Optional, Union

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.charts")


def create_pivot_table(
    backend,
    source_range: Optional[str] = None,
    dest_sheet: Optional[str] = None,
    dest_cell: str = "A3",
    rows: Optional[List[str]] = None,
    cols: Optional[List[str]] = None,
    value_field: Optional[str] = None,
    value_func: str = "sum",
    table_name: Optional[str] = None,
    source_sheet: Optional[str] = None,
    source: Optional[Union[str, Dict[str, Any]]] = None,
    value_format: Optional[str] = None,
    **extra_kwargs: Any
) -> Dict[str, Any]:
    """
    Crea una tabla dinámica (pivot table) a partir de un rango de datos.

    Args:
        backend: Instancia de BaseExcelBackend.
        source_range: Rango de datos con encabezados (ej. "A1:J500").
        dest_sheet: Nombre de la hoja de destino donde se insertará la tabla.
        dest_cell: Celda de inicio del pivot (por defecto "A3").
        rows: Campos de fila del pivot.
        cols: Campos de columna del pivot.
        value_field: Campo de valor a agregar.
        value_func: Función de agregación ("sum", "count", "avg", "max", "min").
        table_name: Nombre identificador de la tabla dinámica.
        source_sheet: Hoja de origen; None usa la activa.

    Returns:
        Diccionario con metadatos del pivot creado.
    """
    logger.debug(
        "create_pivot_table | src=%s dest_sheet=%s dest=%s rows=%s cols=%s val=%s",
        source_range, dest_sheet, dest_cell, rows, cols, value_field
    )
    src = source if source is not None else source_range
    kwargs: Dict[str, Any] = {
        "source": src,
        "dest_sheet": dest_sheet,
        "dest_cell": dest_cell,
        "value_func": value_func,
    }
    if rows is not None:
        kwargs["rows"] = rows
    if cols is not None:
        kwargs["cols"] = cols
    if value_field is not None:
        kwargs["value_field"] = value_field
    if table_name is not None:
        kwargs["table_name"] = table_name
    if source_sheet is not None:
        kwargs["source_sheet"] = source_sheet
    if value_format is not None:
        kwargs["value_format"] = value_format
    kwargs.update(extra_kwargs)
    return backend.create_pivot_table(**kwargs)


def create_chart(
    backend,
    source_range: Optional[str] = None,
    chart_type: str = "column_clustered",
    dest_sheet: Optional[str] = None,
    dest_cell: Optional[str] = None,
    title: Optional[str] = None,
    x_axis: Optional[str] = None,
    y_axis: Optional[str] = None,
    sheet: Optional[str] = None,
    width: Optional[int] = None,
    height: Optional[int] = None,
    source: Optional[Union[str, Dict[str, Any]]] = None,
    placement: str = "auto",
    anchor: Optional[str] = None,
    **extra_kwargs: Any
) -> Dict[str, Any]:
    """
    Crea un gráfico Excel a partir de un rango de datos.

    Args:
        backend: Instancia de BaseExcelBackend.
        source_range: Rango de datos para el gráfico (ej. "A1:B50").
        chart_type: Tipo de gráfico ("column_clustered", "bar", "line", "pie", etc.).
        dest_sheet: Hoja de destino; None usa la hoja activa.
        dest_cell: Celda de anclaje del gráfico (ej. "D3").
        title: Título del gráfico.
        x_axis: Etiqueta del eje X.
        y_axis: Etiqueta del eje Y.
        sheet: Hoja de origen de los datos; None usa la activa.
        width: Ancho del gráfico en puntos.
        height: Alto del gráfico en puntos.

    Returns:
        Diccionario con metadatos del gráfico creado.
    """
    src = source if source is not None else source_range
    logger.debug(
        "create_chart | src=%s type=%s dest_sheet=%s dest=%s",
        src, chart_type, dest_sheet, dest_cell
    )
    kwargs: Dict[str, Any] = {
        "source": src,
        "chart_type": chart_type,
    }
    if dest_sheet is not None:
        kwargs["dest_sheet"] = dest_sheet
    if dest_cell is not None:
        kwargs["anchor"] = dest_cell
    elif anchor is not None:
        kwargs["anchor"] = anchor
    if placement is not None:
        kwargs["placement"] = placement
    if title is not None:
        kwargs["title"] = title
    if x_axis is not None:
        kwargs["x_axis"] = x_axis
    if y_axis is not None:
        kwargs["y_axis"] = y_axis
    if sheet is not None:
        kwargs["source_sheet"] = sheet
    if width is not None:
        kwargs["width"] = width
    if height is not None:
        kwargs["height"] = height
    kwargs.update(extra_kwargs)
    return backend.create_chart(**kwargs)


def create_summary_table(
    backend,
    source_range: str,
    metrics: Optional[List[str]] = None,
    dest_sheet: Optional[str] = None,
    dest_cell: Optional[str] = None,
    sheet: Optional[str] = None,
    title: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Crea una tabla de resumen con métricas estadísticas globales de un rango.

    Args:
        backend: Instancia de BaseExcelBackend.
        source_range: Rango de datos con encabezados.
        metrics: Lista de métricas a incluir (ej. ["sum", "avg", "count", "min", "max"]).
                 None incluye todas las métricas disponibles.
        dest_sheet: Hoja de destino; None usa la hoja activa.
        dest_cell: Celda de inicio del resumen.
        sheet: Hoja de origen; None usa la activa.
        title: Título de la tabla de resumen.

    Returns:
        Diccionario con la tabla generada y su ubicación.
    """
    logger.debug(
        "create_summary_table | src=%s metrics=%s dest=%s", source_range, metrics, dest_cell
    )
    kwargs: Dict[str, Any] = {"source_range": source_range}
    if metrics is not None:
        kwargs["metrics"] = metrics
    if dest_sheet is not None:
        kwargs["dest_sheet"] = dest_sheet
    if dest_cell is not None:
        kwargs["dest_cell"] = dest_cell
    if sheet is not None:
        kwargs["sheet"] = sheet
    if title is not None:
        kwargs["title"] = title
    return backend.create_summary_table(**kwargs)
