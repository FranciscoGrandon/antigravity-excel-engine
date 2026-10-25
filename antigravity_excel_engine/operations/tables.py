# Creado por Francisco Grandón Vergara
"""
Operaciones para gestión de tablas de Excel (ListObject).
Permite crear tablas estructuradas, listar tablas existentes y redimensionarlas
tanto en Live COM como en OpenPyXL Headless.
"""

from typing import Any, Dict, List, Optional
from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.tables")


def create_table(
    backend,
    range_str: str,
    table_name: Optional[str] = None,
    style: str = "TableStyleMedium9",
    has_headers: bool = True,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Convierte un rango libre en tabla estructurada oficial (ListObject).

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango a convertir en tabla (ej. "A1:D10").
        table_name: Nombre único para la tabla (ej. "TblVentas"). Si None, se autogenera.
        style: Estilo visual de tabla (ej. "TableStyleMedium9", "TableStyleLight1").
        has_headers: Si True, la primera fila contiene los encabezados.
        sheet: Nombre de hoja; None usa la hoja activa.

    Returns:
        Diccionario con metadatos de la tabla creada.
    """
    logger.debug(
        "create_table | range=%s name=%s style=%s sheet=%s",
        range_str, table_name, style, sheet
    )
    return backend.create_table(
        range_str=range_str,
        table_name=table_name,
        style=style,
        has_headers=has_headers,
        sheet=sheet,
    )


def list_tables(
    backend,
    sheet: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Lista todas las tablas estructuradas en la hoja especificada o en el libro.

    Args:
        backend: Instancia de BaseExcelBackend.
        sheet: Nombre de hoja; si None, lista las tablas de todas las hojas.

    Returns:
        Lista de diccionarios con nombre, rango, hoja y columnas de cada tabla.
    """
    logger.debug("list_tables | sheet=%s", sheet)
    return backend.list_tables(sheet=sheet)


def resize_table(
    backend,
    table_name: str,
    new_range: str,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Redimensiona el rango de una tabla estructurada existente.

    Args:
        backend: Instancia de BaseExcelBackend.
        table_name: Nombre de la tabla a redimensionar.
        new_range: Nuevo rango (ej. "A1:E20").
        sheet: Nombre de hoja; None usa la hoja activa o busca en el libro.

    Returns:
        Diccionario con confirmación del redimensionamiento.
    """
    logger.debug(
        "resize_table | name=%s new_range=%s sheet=%s",
        table_name, new_range, sheet
    )
    return backend.resize_table(
        table_name=table_name,
        new_range=new_range,
        sheet=sheet,
    )
