# Creado por Francisco Grandón Vergara
"""
Operaciones de curación y limpieza de datos.
Deduplicación, normalización de tipos y pipeline unificado de curación.
"""

from typing import Any, Dict, List, Optional

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.curation")


def remove_duplicates(
    backend,
    range_str: str = "",
    key_columns: Optional[List[str]] = None,
    sheet: Optional[str] = None,
    keep: str = "first",
    range_or_table: Optional[str] = None,
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Elimina filas duplicadas en el rango o tabla especificada.

    Args:
        backend: Instancia de BaseExcelBackend.
        range_str: Rango con encabezados (ej. "A1:H500") o nombre de tabla.
        key_columns: Columnas clave para detectar duplicados; None usa todas.
        sheet: Nombre de hoja; None usa la activa.
        keep: "first" conserva la primera ocurrencia; "last" conserva la última.

    Returns:
        Diccionario con filas eliminadas, filas restantes y detalles.
    """
    target = range_or_table or range_str
    logger.debug(
        "remove_duplicates | target=%s keys=%s keep=%s", target, key_columns, keep
    )
    call_kwargs: Dict[str, Any] = {"range_or_table": target}
    if key_columns is not None:
        call_kwargs["key_columns"] = key_columns
    if sheet is not None:
        call_kwargs["sheet"] = sheet
    return backend.remove_duplicates(**call_kwargs)


def clean_table_dataset(
    backend,
    range_str: str = "",
    sheet: Optional[str] = None,
    rules: Optional[Dict[str, Any]] = None,
    target_range_or_table: Optional[str] = None,
    date_columns: Optional[List[str]] = None,
    numeric_columns: Optional[List[str]] = None,
    trim_strings: bool = True,
    normalize_booleans: bool = True,
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Limpia y normaliza un rango tabular de datos o tabla estructurada.
    """
    target = target_range_or_table or range_str
    logger.debug(
        "clean_table_dataset | target=%s sheet=%s", target, sheet
    )
    if rules is None:
        rules = {"trim_headers": trim_strings}
        if date_columns:
            rules["coerce_dates"] = date_columns
        if numeric_columns:
            rules["cast_types"] = {c: "float" for c in numeric_columns}

    call_kwargs: Dict[str, Any] = {
        "target_range_or_table": target,
        "rules": rules,
    }
    if sheet is not None:
        call_kwargs["sheet"] = sheet
    return backend.clean_table_dataset(**call_kwargs)


def curate_pipeline(
    backend,
    range_str: str = "",
    sheet: Optional[str] = None,
    pipeline_spec: Optional[Dict[str, Any]] = None,
    target_range_or_table: Optional[str] = None,
    steps: Optional[List[str]] = None,
    key_columns: Optional[List[str]] = None,
    date_columns: Optional[List[str]] = None,
    numeric_columns: Optional[List[str]] = None,
    outlier_column: Optional[str] = None,
    outlier_method: str = "iqr",
    **kwargs: Any,
) -> Dict[str, Any]:
    """
    Ejecuta un pipeline completo de curación de datos en una sola pasada en memoria.
    """
    target = target_range_or_table or range_str
    logger.debug(
        "curate_pipeline | target=%s sheet=%s", target, sheet
    )
    if pipeline_spec is None:
        pipeline_spec = {}
        if key_columns:
            pipeline_spec["dedup"] = {"key_columns": key_columns}
        clean_dict = {}
        if date_columns:
            clean_dict["coerce_dates"] = date_columns
        if numeric_columns:
            clean_dict["cast_types"] = {c: "float" for c in numeric_columns}
        if clean_dict:
            pipeline_spec["clean"] = clean_dict
        if outlier_column:
            pipeline_spec["outliers"] = [{"column": outlier_column, "method": outlier_method}]

    call_kwargs: Dict[str, Any] = {
        "target_range_or_table": target,
        "pipeline_spec": pipeline_spec,
    }
    if sheet is not None:
        call_kwargs["sheet"] = sheet
    return backend.curate_pipeline(**call_kwargs)
