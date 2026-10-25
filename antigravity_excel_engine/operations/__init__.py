# Creado por Francisco Grandón Vergara
"""
Subpaquete de operaciones de alto nivel para Antigravity Excel Engine.
Cada módulo encapsula un dominio funcional independiente:
  - read:      lectura de rangos y celdas
  - write:     escritura declarativa de celdas y rangos
  - audit:     auditoría de calidad de datos y detección de errores
  - aggregate: agregaciones estadísticas y columnas calculadas
  - curation:  deduplicación, limpieza y pipelines de curación
  - charts:    tablas dinámicas, gráficos y tablas de resumen
"""

from .aggregate import add_calculated_column, aggregate, compute_aggregation_metric
from .audit import audit_data_quality, check_formula_errors, detect_and_flag_outliers
from .charts import (
    create_chart,
    create_pivot_table,
    create_summary_table,
    extend_chart_series,
    extend_range_reference,
    inspect_charts,
    inspect_workbook,
    parse_series_formula,
    update_chart_series,
)
from .curation import clean_table_dataset, curate_pipeline, remove_duplicates
from .grounding import get_ai_grounding_context, inspect_table_schema
from .read import get_range_as_csv, read_ranges
from .tables import create_table, list_tables, resize_table
from .write import copy_paste_range, write_cell_range, write_cells_declarative

__all__ = [
    "read_ranges",
    "get_range_as_csv",
    "write_cell_range",
    "write_cells_declarative",
    "copy_paste_range",
    "audit_data_quality",
    "check_formula_errors",
    "detect_and_flag_outliers",
    "aggregate",
    "compute_aggregation_metric",
    "add_calculated_column",
    "remove_duplicates",
    "clean_table_dataset",
    "curate_pipeline",
    "create_pivot_table",
    "create_chart",
    "create_summary_table",
    "inspect_charts",
    "extend_chart_series",
    "update_chart_series",
    "inspect_workbook",
    "parse_series_formula",
    "extend_range_reference",
    "create_table",
    "list_tables",
    "resize_table",
    "inspect_table_schema",
    "get_ai_grounding_context",
]

