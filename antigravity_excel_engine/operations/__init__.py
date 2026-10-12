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

from .read import read_ranges, get_range_as_csv
from .write import write_cell_range, write_cells_declarative, copy_paste_range
from .audit import audit_data_quality, check_formula_errors, detect_and_flag_outliers
from .aggregate import aggregate, add_calculated_column
from .curation import remove_duplicates, clean_table_dataset, curate_pipeline
from .charts import create_pivot_table, create_chart, create_summary_table

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
    "add_calculated_column",
    "remove_duplicates",
    "clean_table_dataset",
    "curate_pipeline",
    "create_pivot_table",
    "create_chart",
    "create_summary_table",
]
