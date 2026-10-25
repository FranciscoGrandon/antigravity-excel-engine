# Creado por Francisco Grandón Vergara
"""
CLI modular de Antigravity Excel Engine.

Responsabilidad exclusiva: parsear argumentos y despachar al engine o daemon.
No contiene lógica de negocio Excel.

Integración IPC: intenta primero el daemon residente vía Named Pipe.
Si no está disponible, fallback a ejecución standalone con AntigravityExcelEngine.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any, Optional

from antigravity_excel_engine.ipc.protocol import DEFAULT_TIMEOUT_MS
from antigravity_excel_engine.ipc.client import (
    call_command,
    start_daemon_process,
    status_daemon,
    stop_daemon_process,
)
from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("cli")

_DEFAULT_CURRENCY_FORMAT = '$ #,##0.00;($ #,##0.00);"-"'


# ──────────────────────────────────────────────────────────────────────────────
# Utilidades
# ──────────────────────────────────────────────────────────────────────────────

def _configure_stdout() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def load_json_or_file(arg_val: str) -> Any:
    """Carga JSON desde string, desde archivo @path o desde ruta en disco."""
    if not arg_val:
        return None
    if arg_val.startswith("@"):
        path = arg_val[1:]
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    if os.path.exists(arg_val):
        with open(arg_val, "r", encoding="utf-8") as f:
            return json.load(f)
    return json.loads(arg_val)


def _output(data: Any, as_json: bool) -> None:
    if as_json or isinstance(data, (dict, list)):
        print(json.dumps(data, indent=2, ensure_ascii=False, default=str))
    else:
        print(data)


def _try_daemon(command: str, args_payload: dict, timeout_ms: int = DEFAULT_TIMEOUT_MS) -> Optional[dict]:
    """
    Intenta despachar el comando al daemon residente.
    Retorna la respuesta si el daemon está activo, o None para fallback.
    """
    try:
        return call_command(command, args=args_payload, timeout_ms=timeout_ms)
    except Exception:
        return None


# ──────────────────────────────────────────────────────────────────────────────
# Parser builder
# ──────────────────────────────────────────────────────────────────────────────

def build_parser() -> argparse.ArgumentParser:
    """Construye el parser completo del CLI sin ejecutar nada."""
    parent = argparse.ArgumentParser(add_help=False)
    parent.add_argument("--json", action="store_true", help="Salida en formato JSON")
    parent.add_argument("--mode", choices=["auto", "live", "file"], default="auto")
    parent.add_argument("--file", help="Ruta al archivo .xlsx")

    parser = argparse.ArgumentParser(
        description="Antigravity Excel Engine CLI",
        parents=[parent],
    )
    parser.add_argument("--daemon-start", action="store_true")
    parser.add_argument("--daemon-stop", action="store_true")
    parser.add_argument("--daemon-status", action="store_true")
    sub = parser.add_subparsers(dest="command", required=False)

    # status
    sub.add_parser("status", parents=[parent])

    # get-csv
    p = sub.add_parser("get-csv", parents=[parent])
    p.add_argument("range"); p.add_argument("--sheet"); p.add_argument("--delimiter", default=";")

    # get-ranges
    p = sub.add_parser("get-ranges", parents=[parent])
    p.add_argument("ranges", nargs="+"); p.add_argument("--sheet")
    p.add_argument("--formulas", action="store_true")

    # set-range
    p = sub.add_parser("set-range", parents=[parent])
    p.add_argument("--start-cell", required=True); p.add_argument("--values-json")
    p.add_argument("--input-file"); p.add_argument("--stdin", action="store_true")
    p.add_argument("--sheet"); p.add_argument("--autofit", action="store_true")
    p.add_argument("--dry-run", action="store_true")

    # set-cells
    p = sub.add_parser("set-cells", parents=[parent])
    p.add_argument("--cells-json"); p.add_argument("--input-file")
    p.add_argument("--stdin", action="store_true"); p.add_argument("--sheet")
    p.add_argument("--autofit", action="store_true")
    p.add_argument("--dry-run", action="store_true")

    # copy-paste
    p = sub.add_parser("copy-paste", parents=[parent])
    p.add_argument("--src", required=True); p.add_argument("--dst", required=True)
    p.add_argument("--src-sheet"); p.add_argument("--dst-sheet")
    p.add_argument("--type", choices=["all", "values", "formats", "formulas"], default="all")

    # check-errors
    p = sub.add_parser("check-errors", parents=[parent])
    p.add_argument("range"); p.add_argument("--sheet")

    # audit-quality
    p = sub.add_parser("audit-quality", parents=[parent])
    p.add_argument("--range"); p.add_argument("--sheet")

    # aggregate
    p = sub.add_parser("aggregate", parents=[parent])
    p.add_argument("--group-by", required=True); p.add_argument("--metric", required=True)
    p.add_argument("--func", default="sum",
                   choices=["sum", "avg", "mean", "count", "min", "max"])
    p.add_argument("--sheet"); p.add_argument("--top", type=int, default=10)
    p.add_argument("--ascending", action="store_true")

    # add-column
    p = sub.add_parser("add-column", parents=[parent])
    p.add_argument("--header", required=True); p.add_argument("--formula", required=True)
    p.add_argument("--number-format"); p.add_argument("--sheet")
    p.add_argument("--no-autofit", action="store_true")
    p.add_argument("--dry-run", action="store_true")

    # create-pivot
    p = sub.add_parser("create-pivot", parents=[parent])
    p.add_argument("--source", required=True); p.add_argument("--rows", required=True, nargs="+")
    p.add_argument("--cols", nargs="*"); p.add_argument("--value-field", required=True)
    p.add_argument("--value-func", default="sum",
                   choices=["sum", "count", "average", "avg", "min", "max"])
    p.add_argument("--value-format", default=_DEFAULT_CURRENCY_FORMAT)
    p.add_argument("--dest-sheet"); p.add_argument("--dest-cell", default="A3")
    p.add_argument("--table-name"); p.add_argument("--source-sheet")

    # create-chart
    p = sub.add_parser("create-chart", parents=[parent])
    p.add_argument("--source", required=True)
    p.add_argument("--chart-type", default="column_clustered",
                   choices=["column_clustered", "column", "bar_clustered", "bar",
                            "line", "pie", "area"])
    p.add_argument("--title"); p.add_argument("--dest-sheet")
    p.add_argument("--placement", default="auto", choices=["auto", "right", "bottom"])
    p.add_argument("--anchor"); p.add_argument("--width", type=int, default=480)
    p.add_argument("--height", type=int, default=300); p.add_argument("--source-sheet")

    # apply-preset
    p = sub.add_parser("apply-preset", parents=[parent])
    p.add_argument("--range", required=True)
    p.add_argument("--preset", required=True,
                   choices=["header", "input_cell", "total_row", "currency",
                            "percent", "delta_positive", "delta_negative"])
    p.add_argument("--sheet")

    # summary-table
    p = sub.add_parser("summary-table", parents=[parent])
    p.add_argument("--title"); p.add_argument("--headers", nargs="+")
    p.add_argument("--rows-json"); p.add_argument("--parameters-json")
    p.add_argument("--spec-json"); p.add_argument("--input-file")
    p.add_argument("--stdin", action="store_true")
    p.add_argument("--dest-sheet"); p.add_argument("--start-cell", default="B2")

    # dedup
    p = sub.add_parser("dedup", parents=[parent])
    p.add_argument("--source", required=True)
    p.add_argument("--key-cols", nargs="*"); p.add_argument("--sheet")

    # flag-outliers
    p = sub.add_parser("flag-outliers", parents=[parent])
    p.add_argument("--source"); p.add_argument("--col", required=True)
    p.add_argument("--method", choices=["iqr", "zscore"], default="iqr")
    p.add_argument("--no-comment", action="store_true"); p.add_argument("--sheet")

    # clean
    p = sub.add_parser("clean", parents=[parent])
    p.add_argument("--source", required=True); p.add_argument("--rules-json")
    p.add_argument("--input-file"); p.add_argument("--stdin", action="store_true")
    p.add_argument("--sheet")

    # curate
    p = sub.add_parser("curate", parents=[parent])
    p.add_argument("--source", required=True); p.add_argument("--spec-json")
    p.add_argument("--input-file"); p.add_argument("--stdin", action="store_true")
    p.add_argument("--sheet")

    # wait-until-ready
    p = sub.add_parser("wait-until-ready", aliases=["wait_until_ready"], parents=[parent])
    p.add_argument("--timeout", type=float, default=10.0)
    p.add_argument("--interval", type=float, default=0.25)

    # dispatch-com
    p = sub.add_parser("dispatch-com", aliases=["dispatch_com"], parents=[parent])
    p.add_argument("--target", default="app"); p.add_argument("--member", required=True)
    p.add_argument("--args-json"); p.add_argument("--kwargs-json")
    p.add_argument("--input-file"); p.add_argument("--stdin", action="store_true")

    # inspect-workbook
    p = sub.add_parser("inspect-workbook", parents=[parent])
    p.add_argument("--sheet")
    p.add_argument("--no-charts", action="store_true")

    # inspect-charts
    p = sub.add_parser("inspect-charts", parents=[parent])
    p.add_argument("--sheet")

    # extend-chart-series
    p = sub.add_parser("extend-chart-series", parents=[parent])
    p.add_argument("--to-col")
    p.add_argument("--to-row", type=int)
    p.add_argument("--sheet")
    p.add_argument("--charts", nargs="*")

    # update-chart-series
    p = sub.add_parser("update-chart-series", parents=[parent])
    p.add_argument("--chart", "--chart-name", dest="chart", required=True)
    p.add_argument("--series-index", type=int, required=True)
    p.add_argument("--x-range")
    p.add_argument("--y-range")
    p.add_argument("--name-ref")
    p.add_argument("--formula")
    p.add_argument("--sheet")

    # batch
    p = sub.add_parser("batch", parents=[parent])
    p.add_argument("--operations-json")
    p.add_argument("--input-file")
    p.add_argument("--stdin", action="store_true")

    # create-table
    p = sub.add_parser("create-table", aliases=["create_table"], parents=[parent])
    p.add_argument("--range", "--range-str", dest="range", required=True)
    p.add_argument("--name", "--table-name", dest="table_name")
    p.add_argument("--style", default="TableStyleMedium9")
    p.add_argument("--no-headers", action="store_true")
    p.add_argument("--sheet")

    # inspect-table
    p = sub.add_parser("inspect-table", aliases=["inspect_table"], parents=[parent])
    p.add_argument("--table", "--range", "--table-or-range", dest="table_or_range", required=True)
    p.add_argument("--sheet")
    p.add_argument("--format", choices=["json", "markdown"], default="json")

    return parser


# ──────────────────────────────────────────────────────────────────────────────
# Helpers de lectura de stdin/archivo
# ──────────────────────────────────────────────────────────────────────────────

def _read_json_input(args: argparse.Namespace, json_attr: str) -> Optional[Any]:
    """Lee JSON desde stdin, input-file o atributo inline."""
    if getattr(args, "stdin", False):
        return json.loads(sys.stdin.read())
    if getattr(args, "input_file", None):
        with open(args.input_file, "r", encoding="utf-8") as f:
            return json.load(f)
    raw = getattr(args, json_attr, None)
    if raw:
        return load_json_or_file(raw)
    return None


# ──────────────────────────────────────────────────────────────────────────────
# Dispatcher principal
# ──────────────────────────────────────────────────────────────────────────────

def _build_payload(args: argparse.Namespace) -> dict:
    """Construye el diccionario de argumentos para el daemon o el engine."""
    cmd = args.command
    base = {"mode": args.mode, "file": args.file}

    if cmd == "status":
        return base
    if cmd == "get-csv":
        return {**base, "range": args.range, "sheet": args.sheet, "delimiter": args.delimiter}
    if cmd == "get-ranges":
        return {**base, "ranges": args.ranges, "sheet": args.sheet, "formulas": args.formulas}
    if cmd == "set-range":
        values = _read_json_input(args, "values_json")
        if not values:
            raise ValueError("Debe proporcionar matriz de valores vía --values-json, --input-file o --stdin")
        return {**base, "start_cell": args.start_cell, "values": values,
                "sheet": args.sheet, "autofit": args.autofit,
                "dry_run": getattr(args, "dry_run", False)}
    if cmd == "set-cells":
        cells = _read_json_input(args, "cells_json")
        if not cells:
            raise ValueError("Debe proporcionar mapa de celdas vía --cells-json, --input-file o --stdin")
        return {**base, "cells": cells, "sheet": args.sheet, "autofit": args.autofit,
                "dry_run": getattr(args, "dry_run", False)}
    if cmd == "copy-paste":
        return {**base, "src": args.src, "dst": args.dst,
                "src_sheet": args.src_sheet, "dst_sheet": args.dst_sheet, "type": args.type}
    if cmd == "check-errors":
        return {**base, "range": args.range, "sheet": args.sheet}
    if cmd == "audit-quality":
        return {**base, "range": args.range, "sheet": args.sheet}
    if cmd == "aggregate":
        return {**base, "group_by": args.group_by, "metric": args.metric,
                "func": args.func, "sheet": args.sheet, "top": args.top,
                "ascending": args.ascending}
    if cmd == "add-column":
        return {**base, "header": args.header, "formula": args.formula,
                "number_format": args.number_format, "sheet": args.sheet,
                "autofit": not args.no_autofit,
                "dry_run": getattr(args, "dry_run", False)}
    if cmd == "create-pivot":
        return {**base, "source": args.source, "rows": args.rows, "cols": args.cols,
                "value_field": args.value_field, "value_func": args.value_func,
                "value_format": args.value_format, "dest_sheet": args.dest_sheet,
                "dest_cell": args.dest_cell, "table_name": args.table_name,
                "source_sheet": args.source_sheet}
    if cmd == "create-chart":
        return {**base, "source": args.source, "chart_type": args.chart_type,
                "title": args.title, "dest_sheet": args.dest_sheet,
                "placement": args.placement, "anchor": args.anchor,
                "width": args.width, "height": args.height,
                "source_sheet": args.source_sheet}
    if cmd == "apply-preset":
        return {**base, "range": args.range, "preset": args.preset, "sheet": args.sheet}
    if cmd == "summary-table":
        spec = _read_json_input(args, "spec_json")
        rows_spec = None
        title = args.title or "Resumen"
        headers = args.headers or []
        dest_sheet = args.dest_sheet or "Resumen"
        start_cell = args.start_cell or "B2"
        parameters = None
        if isinstance(spec, dict) and ("rows_spec" in spec or "rows" in spec):
            title = spec.get("title", title)
            headers = spec.get("headers", headers)
            rows_spec = spec.get("rows_spec") or spec.get("rows", [])
            dest_sheet = spec.get("dest_sheet", dest_sheet)
            start_cell = spec.get("start_cell", start_cell)
            parameters = spec.get("parameters")
        else:
            rows_spec = spec if isinstance(spec, list) else None
            if not rows_spec and getattr(args, "rows_json", None):
                rows_spec = load_json_or_file(args.rows_json)
            if not rows_spec:
                raise ValueError("Debe proporcionar especificación de filas vía --rows-json, --spec-json, --input-file o --stdin")
            if getattr(args, "parameters_json", None):
                parameters = load_json_or_file(args.parameters_json)
        return {**base, "title": title, "headers": headers, "rows_spec": rows_spec,
                "dest_sheet": dest_sheet, "start_cell": start_cell, "parameters": parameters,
                "source": ""}
    if cmd == "dedup":
        return {**base, "source": args.source, "key_cols": args.key_cols, "sheet": args.sheet}
    if cmd == "flag-outliers":
        return {**base, "col": args.col, "method": args.method,
                "source": args.source, "sheet": args.sheet,
                "add_comment": not args.no_comment}
    if cmd == "clean":
        rules = _read_json_input(args, "rules_json")
        if not rules or not isinstance(rules, dict):
            raise ValueError("Debe proporcionar un diccionario de reglas vía --rules-json, --input-file o --stdin")
        return {**base, "source": args.source, "rules": rules, "sheet": args.sheet}
    if cmd == "curate":
        spec = _read_json_input(args, "spec_json")
        if not spec or not isinstance(spec, dict):
            raise ValueError("Debe proporcionar especificación de pipeline vía --spec-json, --input-file o --stdin")
        return {**base, "source": args.source, "spec": spec, "sheet": args.sheet}
    if cmd in ("wait-until-ready", "wait_until_ready"):
        return {**base, "timeout": args.timeout, "interval": args.interval}
    if cmd in ("dispatch-com", "dispatch_com"):
        d_args = _read_json_input(args, "args_json")
        d_kwargs = load_json_or_file(args.kwargs_json) if getattr(args, "kwargs_json", None) else None
        if d_args is not None and not isinstance(d_args, list):
            d_args = [d_args]
        return {**base, "target": args.target, "member": args.member,
                "args": d_args, "kwargs": d_kwargs}
    if cmd == "inspect-workbook":
        return {**base, "sheet": args.sheet, "include_charts": not getattr(args, "no_charts", False)}
    if cmd == "inspect-charts":
        return {**base, "sheet": args.sheet}
    if cmd == "extend-chart-series":
        return {**base, "to_col": args.to_col, "to_row": args.to_row,
                "sheet": args.sheet, "charts": args.charts}
    if cmd == "update-chart-series":
        return {**base, "chart_name": args.chart, "series_index": args.series_index,
                "x_range": args.x_range, "y_range": args.y_range,
                "name_ref": args.name_ref, "formula": args.formula, "sheet": args.sheet}
    if cmd == "batch":
        ops = _read_json_input(args, "operations_json")
        if not ops or not isinstance(ops, list):
            raise ValueError("Debe proporcionar una lista de operaciones vía --operations-json, --input-file o --stdin")
        return {**base, "operations": ops}
    if cmd in ("create-table", "create_table"):
        return {**base, "range": args.range, "table_name": getattr(args, "table_name", None),
                "style": getattr(args, "style", "TableStyleMedium9"),
                "has_headers": not getattr(args, "no_headers", False), "sheet": args.sheet}
    if cmd in ("inspect-table", "inspect_table"):
        return {**base, "table_or_range": getattr(args, "table_or_range", ""),
                "sheet": args.sheet, "format": getattr(args, "format", "json")}
    return base


def _run_standalone(cmd: str, payload: dict, as_json: bool) -> None:
    """Ejecuta el comando en modo standalone sin daemon."""
    from antigravity_excel_engine.engine import AntigravityExcelEngine
    engine = AntigravityExcelEngine(mode=payload.get("mode", "auto"), file_path=payload.get("file"))

    try:
        if cmd == "status":
            _output(engine.status(), as_json)

        elif cmd == "get-csv":
            csv_data = engine.get_range_as_csv(
                payload["range"], sheet=payload.get("sheet"), delimiter=payload.get("delimiter", ";")
            )
            if as_json:
                _output({"status": "success", "csv": csv_data}, True)
            else:
                print(csv_data)

        elif cmd == "get-ranges":
            res = engine.get_cell_ranges(
                payload["ranges"], sheet=payload.get("sheet"),
                include_formulas=payload.get("formulas", False)
            )
            _output({"status": "success", "ranges": res}, as_json)

        elif cmd == "set-range":
            dry_run = bool(payload.get("dry_run", False))
            res = engine.set_cell_range(
                payload["start_cell"], payload["values"],
                sheet=payload.get("sheet"), autofit=payload.get("autofit", False),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "set-cells":
            dry_run = bool(payload.get("dry_run", False))
            res = engine.set_cells(
                payload["cells"], sheet=payload.get("sheet"),
                autofit=payload.get("autofit", False),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "copy-paste":
            res = engine.copy_paste_range(
                payload["src"], payload["dst"],
                source_sheet=payload.get("src_sheet"),
                dest_sheet=payload.get("dst_sheet"),
                paste_values_only=(payload.get("type") == "values"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "check-errors":
            errs = engine.check_formula_errors(
                range_str=payload.get("range"), sheet=payload.get("sheet")
            )
            _output({"status": "success", "errors_count": len(errs) if isinstance(errs, list) else 0,
                     "errors": errs}, as_json)

        elif cmd == "audit-quality":
            report = engine.audit_data_quality(
                range_str=payload.get("range"), sheet=payload.get("sheet")
            )
            _output({"status": "success", "audit_report": report}, as_json)

        elif cmd == "aggregate":
            res = engine.aggregate(
                range_str=payload.get("range", ""),
                group_by_col=payload["group_by"],
                metric_col=payload["metric"],
                agg_func=payload.get("func", "sum"),
                sheet=payload.get("sheet"),
                top_n=payload.get("top", 10),
            )
            _output({"status": "success", "aggregation": res}, as_json)

        elif cmd == "add-column":
            dry_run = bool(payload.get("dry_run", False))
            res = engine.add_calculated_column(
                range_str=payload.get("range", ""),
                new_col_name=payload["header"],
                formula_template=payload["formula"],
                sheet=payload.get("sheet"),
                col_format=payload.get("number_format"),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "create-pivot":
            res = engine.create_pivot_table(
                source_range=payload["source"],
                dest_sheet=payload.get("dest_sheet", "Pivot"),
                dest_cell=payload.get("dest_cell", "A3"),
                rows=payload.get("rows"),
                cols=payload.get("cols"),
                value_field=payload.get("value_field"),
                value_func=payload.get("value_func", "sum"),
                table_name=payload.get("table_name"),
                source_sheet=payload.get("source_sheet"),
                value_format=payload.get("value_format"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "create-chart":
            res = engine.create_chart(
                source_range=payload["source"],
                chart_type=payload.get("chart_type", "bar"),
                dest_sheet=payload.get("dest_sheet"),
                dest_cell=payload.get("anchor"),
                title=payload.get("title"),
                sheet=payload.get("source_sheet"),
                width=payload.get("width"),
                height=payload.get("height"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "apply-preset":
            res = engine.apply_style_preset(
                range_str=payload["range"],
                preset_name=payload["preset"],
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "summary-table":
            res = engine.create_summary_table(
                source_range=payload.get("source", ""),
                dest_sheet=payload.get("dest_sheet"),
                dest_cell=payload.get("start_cell"),
                sheet=payload.get("sheet"),
                title=payload.get("title"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "dedup":
            res = engine.remove_duplicates(
                range_str=payload["source"],
                key_columns=payload.get("key_cols"),
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "flag-outliers":
            res = engine.detect_and_flag_outliers(
                range_str=payload.get("source", ""),
                column=payload.get("col"),
                method=payload.get("method", "iqr"),
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "clean":
            res = engine.clean_table_dataset(
                range_str=payload["source"],
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "curate":
            res = engine.curate_pipeline(
                range_str=payload["source"],
                sheet=payload.get("sheet"),
                steps=payload.get("spec", {}).get("steps"),
                key_columns=payload.get("spec", {}).get("key_cols"),
                date_columns=payload.get("spec", {}).get("date_cols"),
                numeric_columns=payload.get("spec", {}).get("numeric_cols"),
                outlier_column=payload.get("spec", {}).get("outlier_col"),
                outlier_method=payload.get("spec", {}).get("outlier_method", "iqr"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd in ("wait-until-ready", "wait_until_ready"):
            res = engine.wait_until_ready(
                timeout_sec=payload.get("timeout", 10.0),
                retry_interval=payload.get("interval", 0.25),
            )
            _output(res, as_json)

        elif cmd in ("dispatch-com", "dispatch_com"):
            res = engine.dispatch_com(
                target=payload.get("target", "app"),
                member=payload.get("member", ""),
                args=payload.get("args"),
                kwargs=payload.get("kwargs"),
            )
            _output(res, as_json)

        elif cmd == "inspect-workbook":
            res = engine.inspect_workbook(
                sheet=payload.get("sheet"),
                include_charts=payload.get("include_charts", True),
            )
            _output(res, as_json)

        elif cmd == "inspect-charts":
            res = engine.inspect_charts(
                sheet=payload.get("sheet"),
            )
            _output(res, as_json)

        elif cmd == "extend-chart-series":
            res = engine.extend_chart_series(
                to_col=payload.get("to_col"),
                to_row=payload.get("to_row"),
                sheet=payload.get("sheet"),
                charts=payload.get("charts"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "update-chart-series":
            res = engine.update_chart_series(
                chart_name=payload["chart_name"],
                series_index=int(payload["series_index"]),
                x_range=payload.get("x_range"),
                y_range=payload.get("y_range"),
                name_ref=payload.get("name_ref"),
                formula=payload.get("formula"),
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd == "batch":
            res = engine.batch(payload.get("operations", []))
            _output(res, as_json)

        elif cmd in ("create-table", "create_table"):
            res = engine.create_table(
                range_str=payload["range"],
                table_name=payload.get("table_name"),
                style=payload.get("style", "TableStyleMedium9"),
                has_headers=payload.get("has_headers", True),
                sheet=payload.get("sheet"),
            )
            res["saved_to"] = engine.save()
            _output(res, as_json)

        elif cmd in ("inspect-table", "inspect_table"):
            fmt = payload.get("format", "json")
            if fmt.lower() == "markdown":
                md_out = engine.get_ai_grounding_context(
                    table_or_range=payload["table_or_range"],
                    sheet=payload.get("sheet"),
                    format="markdown",
                )
                if as_json:
                    _output({"status": "success", "markdown": md_out}, True)
                else:
                    print(md_out)
            else:
                res = engine.inspect_table_schema(
                    table_or_range=payload["table_or_range"],
                    sheet=payload.get("sheet"),
                )
                _output(res, as_json)

        else:
            _output({"status": "error", "error": f"Comando no reconocido: '{cmd}'"}, True)
            sys.exit(1)

    finally:
        if payload.get("mode") == "file" or payload.get("file"):
            try:
                engine.close()
            except Exception:
                pass


# ──────────────────────────────────────────────────────────────────────────────
# Punto de entrada
# ──────────────────────────────────────────────────────────────────────────────

def main() -> None:
    _configure_stdout()
    parser = build_parser()
    args = parser.parse_args()
    as_json = args.json

    # Gestión del ciclo de vida del daemon
    if getattr(args, "daemon_start", False):
        res = start_daemon_process(mode=args.mode, file_path=args.file)
        _output(res, True)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if getattr(args, "daemon_stop", False):
        res = stop_daemon_process()
        _output(res, True)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if getattr(args, "daemon_status", False):
        res = status_daemon()
        _output(res, True)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Construir payload
    try:
        payload = _build_payload(args)
    except (ValueError, FileNotFoundError) as exc:
        logger.error("Error construyendo payload: %s", exc)
        _output({"status": "error", "error": str(exc)}, as_json)
        sys.exit(1)

    # Intentar daemon primero (< 30ms)
    file_path = payload.get("file") or getattr(args, "file", None) or ""
    timeout_ms = 60000 if (str(file_path).startswith(r"\\") or str(file_path).startswith("//")) else DEFAULT_TIMEOUT_MS
    daemon_resp = _try_daemon(args.command, payload, timeout_ms=timeout_ms)
    if daemon_resp is not None:
        if daemon_resp.get("status") == "success":
            result = daemon_resp.get("result")
            if args.command == "get-csv" and not as_json:
                # Para get-csv sin --json: salida plana del CSV
                csv_val = result.get("csv") if isinstance(result, dict) else result
                print(csv_val)
            else:
                _output(result, as_json)
            sys.exit(0)
        else:
            err = daemon_resp.get("error", "Error desconocido del daemon")
            logger.warning("Daemon respondió error: %s", err)
            _output({"status": "error", "error": err}, as_json)
            sys.exit(1)

    # Fallback a ejecución standalone
    try:
        _run_standalone(args.command, payload, as_json)
    except Exception as exc:
        logger.exception("Error en ejecución standalone | cmd=%s", args.command)
        _output({"status": "error", "error": str(exc)}, as_json)
        sys.exit(1)


if __name__ == "__main__":
    main()
