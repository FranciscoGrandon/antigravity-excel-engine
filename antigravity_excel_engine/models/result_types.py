# Creado por Francisco Grandón Vergara
"""
Dataclasses tipadas para los resultados de operaciones analíticas y de manipulación de datos en Excel.
"""

from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional


@dataclass
class CellWriteResult:
    """Resultado de una operación de escritura de celda individual o rango."""
    cell: str
    value: Any
    sheet: str
    previous_value: Optional[Any] = None
    applied_format: Optional[str] = None
    is_formula: bool = False

    def to_dict(self) -> Dict[str, Any]:
        return {
            "cell": self.cell,
            "value": self.value,
            "sheet": self.sheet,
            "previous_value": self.previous_value,
            "applied_format": self.applied_format,
            "is_formula": self.is_formula,
        }


@dataclass
class AuditReport:
    """Informe detallado de auditoría de calidad de datos en un rango u hoja."""
    sheet: str
    range_audited: str
    total_cells: int
    blank_cells: int
    formula_count: int
    formula_errors: List[Dict[str, Any]] = field(default_factory=list)
    data_types: Dict[str, int] = field(default_factory=dict)
    summary: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "sheet": self.sheet,
            "range_audited": self.range_audited,
            "total_cells": self.total_cells,
            "blank_cells": self.blank_cells,
            "formula_count": self.formula_count,
            "formula_errors": self.formula_errors,
            "data_types": self.data_types,
            "summary": self.summary,
        }


@dataclass
class CurateResult:
    """Resultado de la ejecución de un pipeline compuesto de curación."""
    target: str
    sheet: str
    steps_executed: List[str] = field(default_factory=list)
    rows_affected: int = 0
    duplicates_removed: int = 0
    outliers_flagged: int = 0
    cleaning_stats: Dict[str, Any] = field(default_factory=dict)
    execution_time_sec: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "target": self.target,
            "sheet": self.sheet,
            "steps_executed": self.steps_executed,
            "rows_affected": self.rows_affected,
            "duplicates_removed": self.duplicates_removed,
            "outliers_flagged": self.outliers_flagged,
            "cleaning_stats": self.cleaning_stats,
            "execution_time_sec": self.execution_time_sec,
        }


@dataclass
class AggregateResult:
    """Resultado de agregación analítica (SUM, AVG, COUNT, MIN, MAX)."""
    group_by_col: str
    metric_col: str
    agg_func: str
    data: List[Dict[str, Any]] = field(default_factory=list)
    top_n: int = 10
    total_groups: int = 0

    def to_dict(self) -> Dict[str, Any]:
        return {
            "group_by_col": self.group_by_col,
            "metric_col": self.metric_col,
            "agg_func": self.agg_func,
            "data": self.data,
            "top_n": self.top_n,
            "total_groups": self.total_groups,
        }


@dataclass
class PivotResult:
    """Resultado de la creación o actualización de una tabla dinámica."""
    table_name: str
    source_range: str
    dest_sheet: str
    dest_cell: str
    rows: List[str] = field(default_factory=list)
    cols: List[str] = field(default_factory=list)
    value_field: Optional[str] = None
    value_func: str = "sum"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "table_name": self.table_name,
            "source_range": self.source_range,
            "dest_sheet": self.dest_sheet,
            "dest_cell": self.dest_cell,
            "rows": self.rows,
            "cols": self.cols,
            "value_field": self.value_field,
            "value_func": self.value_func,
        }


@dataclass
class OutlierResult:
    """Resultado de detección y etiquetado de valores atípicos (outliers)."""
    column: str
    method: str
    outliers_count: int
    outlier_cells: List[str] = field(default_factory=list)
    thresholds: Dict[str, Any] = field(default_factory=dict)
    flag_style: str = "warning"

    def to_dict(self) -> Dict[str, Any]:
        return {
            "column": self.column,
            "method": self.method,
            "outliers_count": self.outliers_count,
            "outlier_cells": self.outlier_cells,
            "thresholds": self.thresholds,
            "flag_style": self.flag_style,
        }
