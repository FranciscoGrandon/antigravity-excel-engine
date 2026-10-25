# Creado por Francisco Grandón Vergara
"""
Operaciones de escritura de celdas y rangos.
Delegación al backend activo sin lógica de negocio propia.
"""

from typing import Any, Dict, List, Optional

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.write")


def write_cell_range(
    backend,
    start_cell: str,
    values: List[List[Any]],
    sheet: Optional[str] = None,
    formulas: Optional[List[List[str]]] = None,
    styles: Optional[Dict[str, Any]] = None,
    autofit: bool = False,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """
    Escribe una matriz rectangular de celdas a partir de una celda inicial.

    Args:
        backend: Instancia de BaseExcelBackend.
        start_cell: Celda de inicio (ej. "B3").
        values: Matriz de filas y columnas con los valores a escribir.
        sheet: Nombre de hoja; None usa la hoja activa.
        formulas: Matriz opcional de fórmulas en formato Excel.
        styles: Diccionario de estilos aplicado al rango.
        autofit: Si True, ajusta automáticamente el ancho de columnas.
        dry_run: Si True, simula la operación sin modificar el libro.

    Returns:
        Diccionario con resultado de la operación.
    """
    logger.debug(
        "write_cell_range | start=%s sheet=%s rows=%d dry_run=%s",
        start_cell, sheet, len(values) if values else 0, dry_run
    )
    return backend.set_cell_range(
        start_cell=start_cell,
        values=values,
        sheet=sheet,
        formulas=formulas,
        styles=styles,
        autofit=autofit,
        dry_run=dry_run,
    )


def write_cells_declarative(
    backend,
    cells_map: Dict[str, Dict[str, Any]],
    sheet: Optional[str] = None,
    autofit: bool = False,
    dry_run: bool = False,
) -> Dict[str, Any]:
    """
    Escritura declarativa de celdas individuales con valores, fórmulas y estilos.

    Args:
        backend: Instancia de BaseExcelBackend.
        cells_map: Diccionario { "A1": {"value": ..., "formula": ..., "style": ...} }.
        sheet: Nombre de hoja; None usa la hoja activa.
        autofit: Si True, ajusta automáticamente el ancho de columnas.
        dry_run: Si True, simula la operación sin modificar el libro.

    Returns:
        Diccionario con resultado de la operación.
    """
    logger.debug("write_cells_declarative | cells=%d sheet=%s dry_run=%s", len(cells_map), sheet, dry_run)
    return backend.set_cells(cells_map=cells_map, sheet=sheet, autofit=autofit, dry_run=dry_run)


def copy_paste_range(
    backend,
    source_range: str,
    dest_cell: str,
    source_sheet: Optional[str] = None,
    dest_sheet: Optional[str] = None,
    paste_values_only: bool = False,
) -> Dict[str, Any]:
    """
    Copia un rango de celdas a una ubicación destino.

    Args:
        backend: Instancia de BaseExcelBackend.
        source_range: Rango origen (ej. "A1:D10").
        dest_cell: Celda destino (ej. "F1").
        source_sheet: Hoja de origen; None usa la activa.
        dest_sheet: Hoja de destino; None usa la activa.
        paste_values_only: Si True, pega solo valores (no fórmulas ni formatos).

    Returns:
        Diccionario con resultado de la operación.
    """
    logger.debug(
        "copy_paste_range | src=%s dest=%s paste_values_only=%s",
        source_range, dest_cell, paste_values_only
    )
    paste_type = "values" if paste_values_only else "all"
    return backend.copy_paste_range(
        src_range=source_range,
        dst_start_cell=dest_cell,
        src_sheet=source_sheet,
        dst_sheet=dest_sheet,
        paste_type=paste_type,
    )
