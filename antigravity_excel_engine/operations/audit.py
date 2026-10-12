# Creado por Francisco Grandón Vergara
"""
Operaciones de auditoría de integridad de datos y detección de anomalías.
Delegación al backend activo con retorno de tipos estructurados cuando aplica.
"""

from typing import Any, Dict, List, Optional

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.audit")


def audit_data_quality(
    backend,
    range_str: Optional[str] = None,
    sheet: Optional[str] = None,
    checks: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Audita la calidad de datos en un rango u hoja completa.

    Verifica: celdas vacías, errores de fórmula, tipos inconsistentes,
    duplicados y valores atípicos según los checks seleccionados.

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango a auditar (ej. "A1:F200"); None audita la hoja activa.
        sheet: Nombre de hoja; None usa la activa.
        checks: Lista de chequeos a ejecutar (ej. ["blanks", "types", "formulas"]).
                None ejecuta todos los disponibles.

    Returns:
        Diccionario con informe de auditoría.
    """
    logger.debug("audit_data_quality | range=%s sheet=%s checks=%s", range_str, sheet, checks)
    kwargs: Dict[str, Any] = {}
    if range_str is not None:
        kwargs["range_str"] = range_str
    if sheet is not None:
        kwargs["sheet"] = sheet
    if checks is not None:
        kwargs["checks"] = checks
    return backend.audit_data_quality(**kwargs)


def check_formula_errors(
    backend,
    range_str: Optional[str] = None,
    sheet: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Detecta celdas con errores de fórmula (#VALUE!, #REF!, #DIV/0!, etc.).

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango a inspeccionar; None inspecciona la hoja activa.
        sheet: Nombre de hoja; None usa la activa.

    Returns:
        Diccionario con lista de celdas erróneas y tipos de error encontrados.
    """
    logger.debug("check_formula_errors | range=%s sheet=%s", range_str, sheet)
    kwargs: Dict[str, Any] = {}
    if range_str is not None:
        kwargs["range_str"] = range_str
    if sheet is not None:
        kwargs["sheet"] = sheet
    return backend.check_formula_errors(**kwargs)


def detect_and_flag_outliers(
    backend,
    range_str: str,
    column: Optional[str] = None,
    method: str = "iqr",
    flag_style: str = "warning",
    sheet: Optional[str] = None,
    threshold: float = 1.5,
) -> Dict[str, Any]:
    """
    Detecta valores atípicos en un rango numérico y los marca visualmente.

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango de datos (ej. "A1:A500").
        column: Nombre de columna si el rango tiene encabezados.
        method: Método de detección ("iqr" o "zscore").
        flag_style: Preset de estilo para marcar outliers ("warning", "alert", etc.).
        sheet: Nombre de hoja; None usa la activa.
        threshold: Multiplicador IQR o umbral z-score.

    Returns:
        Diccionario con lista de celdas marcadas, umbrales calculados y conteo.
    """
    logger.debug(
        "detect_and_flag_outliers | range=%s col=%s method=%s threshold=%s",
        range_str, column, method, threshold
    )
    kwargs: Dict[str, Any] = {
        "range_str": range_str,
        "method": method,
        "flag_style": flag_style,
        "threshold": threshold,
    }
    if column is not None:
        kwargs["column"] = column
    if sheet is not None:
        kwargs["sheet"] = sheet
    return backend.detect_and_flag_outliers(**kwargs)
