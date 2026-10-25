# Creado por Francisco Grandón Vergara
"""
Módulo para AI Grounding & Table Schemas en Excel Engine.
Inspecciona esquemas estructurados de tablas y rangos para inyección de contexto
en prompts de modelos de lenguaje (LLMs) con costo mínimo de tokens.
"""

import datetime
import json
import re
from typing import Any, Dict, List, Optional, Union

from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("operations.grounding")


def _is_date_value(val: Any) -> bool:
    """Detecta si un valor representa una fecha o datetime."""
    if isinstance(val, (datetime.date, datetime.datetime)):
        return True
    if isinstance(val, str):
        val_str = val.strip()
        # ISO date YYYY-MM-DD o DD-MM-YYYY / DD/MM/YYYY
        if re.match(r"^\d{4}[-/]\d{2}[-/]\d{2}$", val_str) or re.match(r"^\d{2}[-/]\d{2}[-/]\d{4}$", val_str):
            return True
    return False


def _is_currency_value(val: Any) -> bool:
    """Detecta si un valor textual tiene formato explícito de moneda."""
    if isinstance(val, str):
        val_str = val.strip()
        if any(sym in val_str for sym in ("$", "€", "£", "CLP", "UF", "USD")):
            # Eliminar símbolos y ver si es numérico
            cleaned = re.sub(r"[\$\€\£\s]|CLP|UF|USD", "", val_str).replace(".", "").replace(",", ".")
            try:
                float(cleaned)
                return True
            except ValueError:
                pass
    return False


def _is_boolean_value(val: Any) -> bool:
    """Detecta si un valor representa un booleano."""
    if isinstance(val, bool):
        return True
    if isinstance(val, str):
        return val.strip().lower() in ("true", "false", "verdadero", "falso")
    return False


def _infer_column_type(values: List[Any], col_name: str = "") -> str:
    """
    Infiere el tipo de datos semántico de una columna a partir de sus valores no nulos.
    Retorna uno de: 'integer', 'float', 'date', 'currency', 'boolean', 'text', 'empty'.
    """
    if not values:
        return "empty"

    type_counts: Dict[str, int] = {
        "boolean": 0,
        "currency": 0,
        "integer": 0,
        "float": 0,
        "date": 0,
        "text": 0,
    }

    for val in values:
        if _is_boolean_value(val):
            type_counts["boolean"] += 1
        elif _is_currency_value(val):
            type_counts["currency"] += 1
        elif _is_date_value(val):
            type_counts["date"] += 1
        elif isinstance(val, int) and not isinstance(val, bool):
            type_counts["integer"] += 1
        elif isinstance(val, float):
            # Si el float es entero exacto (ej. 10.0), podría ser int
            if val.is_integer():
                type_counts["integer"] += 1
            else:
                type_counts["float"] += 1
        elif isinstance(val, str):
            val_s = val.strip()
            # Si es código con ceros a la izquierda (ej. "00123", "0987") o RUT/teléfono con guión (ej. "12345678-9")
            if (val_s.startswith("0") and len(val_s) > 1 and val_s.isdigit()) or (
                bool(re.search(r"\d+-\d+", val_s) or re.search(r"\d+-[\dkK]$", val_s))
            ):
                type_counts["text"] += 1
                continue

            # Intentar parseo de int o float
            try:
                iv = int(val_s)
                type_counts["integer"] += 1
                continue
            except ValueError:
                pass
            try:
                fv = float(val_s.replace(",", "."))
                if fv.is_integer():
                    type_counts["integer"] += 1
                else:
                    type_counts["float"] += 1
                continue
            except ValueError:
                pass
            type_counts["text"] += 1
        else:
            type_counts["text"] += 1

    total = len(values)
    # Si la mayoría son booleanos
    if type_counts["boolean"] > total * 0.5:
        return "boolean"
    # Si hay moneda explícita
    if type_counts["currency"] > total * 0.5:
        return "currency"
    # Si la mayoría son fechas
    if type_counts["date"] > total * 0.5:
        return "date"
    # Si hay números (int + float)
    num_count = type_counts["integer"] + type_counts["float"]
    if num_count > total * 0.5:
        # Verificar si el nombre sugiere moneda
        name_lower = col_name.lower()
        if any(w in name_lower for w in ("precio", "monto", "venta", "costo", "ingreso", "tarifa", "total", "saldo", "valor")):
            return "currency"
        if type_counts["float"] > 0:
            return "float"
        return "integer"

    return "text"


def inspect_table_schema(
    backend,
    table_or_range: str,
    sheet: Optional[str] = None,
    max_sample_rows: int = 1000,
) -> Dict[str, Any]:
    """
    Inspecciona y describe exhaustivamente el esquema de una tabla de Excel o rango libre.

    Args:
        backend: Instancia de BaseExcelBackend.
        table_or_range: Nombre de la tabla (ej. "TblFacturacion") o rango (ej. "A1:E20").
        sheet: Nombre opcional de la hoja.

    Returns:
        Diccionario con nombre de tabla, hoja, rango, dimensiones y detalles de columnas
        (nombre, índice, tipo inferido, no nulos, nulos, muestra de valores, min, max, fórmulas).
    """
    resolved_sheet = sheet
    range_str = table_or_range
    table_name: Optional[str] = None

    # Intentar resolver como tabla estructurada si existe
    try:
        tables = backend.list_tables(sheet=sheet)
        for tbl in tables:
            if tbl.get("name", "").lower() == table_or_range.strip().lower():
                table_name = tbl["name"]
                range_str = tbl["range"]
                if not resolved_sheet:
                    resolved_sheet = tbl.get("sheet")
                break
    except Exception as exc:
        logger.debug("No se pudieron consultar tablas en el backend: %s", exc)

    # Si no se encontró tabla y no tiene dos puntos, verificar si es nombre de tabla o celda
    if table_name is None and ":" not in range_str:
        table_name = table_or_range

    read_data = backend.read_range(
        range_str=range_str,
        sheet=resolved_sheet,
        include_formulas=True,
    )

    values: List[List[Any]] = read_data.get("values", [])
    formulas: List[List[Any]] = read_data.get("formulas", [])

    if not values:
        return {
            "table_name": table_name,
            "sheet": resolved_sheet,
            "range": range_str,
            "dimensions": {"rows": 0, "cols": 0, "data_rows": 0},
            "columns": [],
        }

    header_row = values[0]
    data_rows = values[1:]
    formula_rows = formulas[1:] if formulas and len(formulas) > 1 else []
    total_cols = len(header_row)
    data_rows_count = len(data_rows)

    # Muestreo eficiente para grandes volúmenes de datos
    is_sampled = bool(max_sample_rows and data_rows_count > max_sample_rows)
    if is_sampled:
        step = max(1, data_rows_count // max_sample_rows)
        sampled_data_rows = data_rows[::step][:max_sample_rows]
        sampled_formula_rows = formula_rows[::step][:max_sample_rows] if formula_rows else []
    else:
        sampled_data_rows = data_rows
        sampled_formula_rows = formula_rows

    # Transposición matricial vectorizada en un solo paso con zip(*...)
    padded_rows = [
        row + [None] * (total_cols - len(row)) if len(row) < total_cols else row[:total_cols]
        for row in sampled_data_rows
    ]
    transposed_cols = list(zip(*padded_rows)) if padded_rows else [() for _ in range(total_cols)]

    padded_formula_rows = [
        row + [None] * (total_cols - len(row)) if len(row) < total_cols else row[:total_cols]
        for row in sampled_formula_rows
    ]
    transposed_formulas = list(zip(*padded_formula_rows)) if padded_formula_rows else [() for _ in range(total_cols)]

    columns: List[Dict[str, Any]] = []

    for c_idx in range(total_cols):
        raw_header = header_row[c_idx]
        col_name = str(raw_header).strip() if raw_header is not None else f"Column{c_idx + 1}"

        # Extraer valores y fórmulas de esta columna vía tupla transpuesta
        col_values = transposed_cols[c_idx] if c_idx < len(transposed_cols) else ()
        col_formulas = transposed_formulas[c_idx] if c_idx < len(transposed_formulas) else ()

        has_formula = any(
            isinstance(f, str) and f.strip().startswith("=")
            for f in col_formulas
            if f is not None
        )

        non_null_values = [v for v in col_values if v is not None and v != ""]
        if is_sampled and col_values:
            ratio = len(non_null_values) / len(col_values)
            non_null_count = int(round(ratio * data_rows_count))
            null_count = data_rows_count - non_null_count
        else:
            non_null_count = len(non_null_values)
            null_count = data_rows_count - non_null_count

        inferred_type = _infer_column_type(non_null_values, col_name=col_name)

        # Muestra de hasta 3 valores representativos únicos
        sample_values: List[Any] = []
        seen = set()
        for v in non_null_values:
            v_rep = str(v)
            if v_rep not in seen:
                seen.add(v_rep)
                # Normalizar si es objeto fecha
                if isinstance(v, (datetime.date, datetime.datetime)):
                    sample_values.append(v.isoformat())
                else:
                    sample_values.append(v)
                if len(sample_values) >= 3:
                    break

        # Calcular min y max numéricos o de fecha
        col_min: Optional[Any] = None
        col_max: Optional[Any] = None

        if inferred_type in ("integer", "float", "currency"):
            numeric_vals: List[Union[int, float]] = []
            for v in non_null_values:
                if isinstance(v, (int, float)) and not isinstance(v, bool):
                    numeric_vals.append(v)
                elif isinstance(v, str):
                    try:
                        cleaned = re.sub(r"[\$\€\£\s]|CLP|UF|USD", "", v).replace(".", "").replace(",", ".")
                        numeric_vals.append(float(cleaned))
                    except ValueError:
                        pass
            if numeric_vals:
                col_min = min(numeric_vals)
                col_max = max(numeric_vals)
                if inferred_type == "integer" and isinstance(col_min, float) and col_min.is_integer():
                    col_min = int(col_min)
                if inferred_type == "integer" and isinstance(col_max, float) and col_max.is_integer():
                    col_max = int(col_max)

        elif inferred_type == "date":
            date_vals: List[Any] = []
            for v in non_null_values:
                if isinstance(v, (datetime.date, datetime.datetime)):
                    date_vals.append(v)
                elif isinstance(v, str):
                    date_vals.append(v)
            if date_vals:
                try:
                    col_min = str(min(date_vals))
                    col_max = str(max(date_vals))
                except Exception:
                    pass

        columns.append({
            "name": col_name,
            "index": c_idx + 1,
            "inferred_type": inferred_type,
            "non_null_count": non_null_count,
            "null_count": null_count,
            "sample_values": sample_values,
            "min": col_min,
            "max": col_max,
            "has_formula": has_formula,
        })

    return {
        "table_name": table_name,
        "sheet": resolved_sheet,
        "range": range_str,
        "dimensions": {
            "rows": len(values),
            "cols": total_cols,
            "data_rows": data_rows_count,
        },
        "columns": columns,
    }


def get_ai_grounding_context(
    backend,
    table_or_range: str,
    sheet: Optional[str] = None,
    format: str = "json",
    max_sample_rows: int = 1000,
) -> Union[str, Dict[str, Any]]:
    """
    Serializa el esquema de una tabla o rango para inyección directa en prompts de LLMs.

    Args:
        backend: Instancia de BaseExcelBackend.
        table_or_range: Nombre de la tabla o rango de celdas.
        sheet: Hoja de trabajo opcional.
        format: Formato de salida: 'json' o 'markdown'.
        max_sample_rows: Límite máximo de filas de muestra para inferencia rápida.

    Returns:
        Cadena formateada en JSON o Markdown lista para grounding de prompts.
    """
    schema = inspect_table_schema(backend, table_or_range, sheet=sheet, max_sample_rows=max_sample_rows)

    if format.lower() == "json":
        return json.dumps(schema, indent=2, ensure_ascii=False, default=str)

    if format.lower() == "markdown":
        tbl_label = schema.get("table_name") or schema.get("range", "Tabla")
        sheet_label = schema.get("sheet") or "Activa"
        dims = schema.get("dimensions", {})
        total_rows = dims.get("rows", 0)
        total_cols = dims.get("cols", 0)
        data_rows = dims.get("data_rows", 0)
        rng = schema.get("range", "")

        lines: List[str] = [
            f"### Esquema de Tabla: `{tbl_label}` (Hoja: `{sheet_label}`)",
            f"- **Rango:** `{rng}` | **Dimensiones:** {total_rows} filas × {total_cols} columnas ({data_rows} registros de datos)",
            "",
            "| Columna | Índice | Tipo | No Nulos | Nulos | Muestra de Valores | Min | Max | Fórmula |",
            "|:---|:---:|:---:|:---:|:---:|:---|:---:|:---:|:---:|",
        ]

        for col in schema.get("columns", []):
            name = col["name"]
            idx = col["index"]
            itype = col["inferred_type"]
            non_null = col["non_null_count"]
            nulls = col["null_count"]
            samples = ", ".join(f"'{s}'" if isinstance(s, str) else str(s) for s in col["sample_values"])
            if not samples:
                samples = "*(vacío)*"
            min_val = str(col["min"]) if col["min"] is not None else "-"
            max_val = str(col["max"]) if col["max"] is not None else "-"
            formula_flag = "Sí" if col["has_formula"] else "No"

            lines.append(
                f"| `{name}` | {idx} | `{itype}` | {non_null} | {nulls} | {samples} | {min_val} | {max_val} | {formula_flag} |"
            )

        return "\n".join(lines)

    return json.dumps(schema, indent=2, ensure_ascii=False, default=str)
