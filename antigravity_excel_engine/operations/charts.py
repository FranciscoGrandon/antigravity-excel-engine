# Creado por Francisco Grandón Vergara
"""
Operaciones de visualización: tablas dinámicas, gráficos y tablas de resumen.
Delegación al backend activo sin lógica de negocio propia.
"""

import re
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


def parse_series_formula(formula: str) -> Dict[str, str]:
    """
    Parsea una fórmula =SERIES(name_ref, x_range, y_range, order) de Excel,
    respetando comillas simples de nombres de hoja y comillas dobles de texto.

    Args:
        formula: Cadena con la fórmula de la serie.

    Returns:
        Diccionario con claves: name_ref, x_range, y_range, order.
    """
    res = {"name_ref": "", "x_range": "", "y_range": "", "order": "1"}
    if not formula:
        return res

    raw = formula.strip()
    m = re.match(r"^=?SERIES\s*\((.*)\)\s*$", raw, re.IGNORECASE | re.DOTALL)
    if not m:
        return res

    inner = m.group(1).strip()
    parts = []
    current = []
    in_single_quote = False
    in_double_quote = False
    paren_depth = 0

    for ch in inner:
        if ch == "'" and not in_double_quote:
            in_single_quote = not in_single_quote
            current.append(ch)
        elif ch == '"' and not in_single_quote:
            in_double_quote = not in_double_quote
            current.append(ch)
        elif ch == '(' and not in_single_quote and not in_double_quote:
            paren_depth += 1
            current.append(ch)
        elif ch == ')' and not in_single_quote and not in_double_quote:
            paren_depth = max(0, paren_depth - 1)
            current.append(ch)
        elif ch == ',' and not in_single_quote and not in_double_quote and paren_depth == 0:
            parts.append("".join(current).strip())
            current = []
        else:
            current.append(ch)

    parts.append("".join(current).strip())

    if len(parts) >= 4:
        res["name_ref"] = parts[0]
        res["x_range"] = parts[1]
        res["y_range"] = parts[2]
        res["order"] = parts[3]
    elif len(parts) == 3:
        res["name_ref"] = ""
        res["x_range"] = parts[0]
        res["y_range"] = parts[1]
        res["order"] = parts[2]
    elif len(parts) == 2:
        res["x_range"] = parts[0]
        res["y_range"] = parts[1]
    elif len(parts) == 1:
        res["y_range"] = parts[0]

    return res


def _extend_area(area_ref: str, to_col: Optional[str] = None, to_row: Optional[int] = None) -> str:
    """Helper interno para extender una coordenada o rango A1."""
    if not area_ref:
        return area_ref

    sheet_prefix = ""
    target_part = area_ref.strip()

    m_sheet = re.match(r"^((?:'[^']+'|[^!]+)!)(.*)$", target_part)
    if m_sheet:
        sheet_prefix = m_sheet.group(1)
        target_part = m_sheet.group(2).strip()

    if ":" in target_part:
        start_cell, end_cell = target_part.split(":", 1)
    else:
        start_cell = target_part
        end_cell = target_part

    m_cell = re.match(r"^(\$?)([A-Za-z]+)(\$?)(-?\d+)$", end_cell.strip())
    if not m_cell:
        return area_ref

    col_dollar, col_letters, row_dollar, row_num = m_cell.groups()

    new_col = to_col.upper() if to_col else col_letters
    new_row = str(to_row) if to_row is not None else row_num

    new_end_cell = f"{col_dollar}{new_col}{row_dollar}{new_row}"

    return f"{sheet_prefix}{start_cell}:{new_end_cell}"


def extend_range_reference(ref: str, to_col: Optional[str] = None, to_row: Optional[int] = None) -> str:
    """
    Extiende la referencia de rango hacia to_col y/o to_row,
    soportando referencias múltiples entre paréntesis.
    """
    if not ref:
        return ref
    ref_stripped = ref.strip()
    if ref_stripped.startswith("(") and ref_stripped.endswith(")"):
        inner = ref_stripped[1:-1]
        areas = [a.strip() for a in inner.split(",")]
        if areas:
            areas[-1] = _extend_area(areas[-1], to_col=to_col, to_row=to_row)
        return f"({','.join(areas)})"
    return _extend_area(ref_stripped, to_col=to_col, to_row=to_row)


def inspect_charts(backend, sheet: Optional[str] = None) -> Dict[str, Any]:
    """Inspecciona gráficos en la hoja indicada (o todas)."""
    return backend.inspect_charts(sheet=sheet)


def extend_chart_series(
    backend,
    to_col: Optional[str] = None,
    to_row: Optional[int] = None,
    sheet: Optional[str] = None,
    charts: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Extiende series y categorías de los gráficos indicados hasta to_col y/o to_row."""
    return backend.extend_chart_series(to_col=to_col, to_row=to_row, sheet=sheet, charts=charts)


def update_chart_series(
    backend,
    chart_name: str,
    series_index: int,
    x_range: Optional[str] = None,
    y_range: Optional[str] = None,
    name_ref: Optional[str] = None,
    formula: Optional[str] = None,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """Actualiza una serie específica de un gráfico."""
    return backend.update_chart_series(
        chart_name=chart_name,
        series_index=series_index,
        x_range=x_range,
        y_range=y_range,
        name_ref=name_ref,
        formula=formula,
        sheet=sheet,
    )


def inspect_workbook(backend, sheet: Optional[str] = None, include_charts: bool = True) -> Dict[str, Any]:
    """Inspecciona metadatos del libro y detalle por hoja."""
    return backend.inspect_workbook(sheet=sheet, include_charts=include_charts)

