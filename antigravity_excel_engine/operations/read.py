# Creado por Francisco Grandón Vergara
"""
Operaciones de lectura de rangos y celdas.
Delegación al backend activo sin lógica de negocio propia.
"""

from typing import Any, Dict, List, Optional

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.read")


def read_ranges(
    backend,
    ranges: List[str],
    sheet: Optional[str] = None,
    include_formulas: bool = False,
    include_styles: bool = False,
) -> Dict[str, Any]:
    """
    Lee uno o varios rangos de celdas desde el backend dado.

    Args:
        backend: Instancia de BaseExcelBackend (live COM o headless).
        ranges: Lista de referencias de rango en formato A1 (ej. ["A1:C10", "E5"]).
        sheet: Nombre de la hoja; None usa la hoja activa.
        include_formulas: Si True, devuelve fórmulas en lugar de valores calculados.
        include_styles: Si True, incluye información de estilo por celda.

    Returns:
        Diccionario con clave = rango, valor = datos leídos.
    """
    logger.debug(
        "read_ranges | ranges=%s sheet=%s include_formulas=%s",
        ranges, sheet, include_formulas
    )
    return backend.get_cell_ranges(
        ranges=ranges,
        sheet=sheet,
        include_formulas=include_formulas,
        include_styles=include_styles,
    )


def get_range_as_csv(
    backend,
    range_str: str,
    sheet: Optional[str] = None,
    delimiter: str = ";",
    include_headers: bool = True,
) -> str:
    """
    Exporta un rango de celdas como texto CSV.

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Referencia de rango (ej. "A1:G100").
        sheet: Nombre de hoja; None usa la hoja activa.
        delimiter: Carácter separador (por defecto ";").
        include_headers: Si True incluye la primera fila como encabezados.

    Returns:
        Cadena CSV con los datos del rango.
    """
    logger.debug("get_range_as_csv | range=%s sheet=%s delimiter=%r", range_str, sheet, delimiter)
    return backend.get_range_as_csv(
        range_str=range_str,
        sheet=sheet,
        delimiter=delimiter,
    )
