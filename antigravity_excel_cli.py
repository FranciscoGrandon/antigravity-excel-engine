# Creado por Francisco Grandón Vergara
"""
CLI Oficial de Antigravity Excel Engine.
Proporciona salida JSON determinista para subagentes y llamadas por lotes.
"""

import argparse
import json
import sys
import os
from typing import Any, Dict, Optional

# Formato por defecto para argumentos CLI sin forzar import anticipado de core
_DEFAULT_CURRENCY_FORMAT = "$ #,##0.00;($ #,##0.00);\"-\""


def _try_call_daemon(command: str, args_payload: dict, timeout_ms: int = 50) -> Optional[dict]:
    """
    Intenta enviar la petición al daemon residente vía Named Pipe con timeout ultra-corto.
    Retorna la respuesta JSON del daemon si está activo (< 30ms), o None para fallback transparente.
    """
    try:
        from antigravity_excel_daemon import send_pipe_message
        payload = {"command": command, "args": args_payload}
        return send_pipe_message(payload, timeout_ms=timeout_ms)
    except Exception:
        return None


def load_json_or_file(arg_val: str) -> Any:
    """Carga JSON desde string o desde archivo si comienza con '@' o existe en disco."""
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


def main():
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parent_parser = argparse.ArgumentParser(add_help=False)

    parent_parser.add_argument("--json", action="store_true", help="Salida en formato JSON")
    parent_parser.add_argument("--mode", choices=["auto", "live", "file"], default="auto", help="Modo de ejecución")
    parent_parser.add_argument("--file", help="Ruta al archivo .xlsx")

    parser = argparse.ArgumentParser(description="Antigravity Excel Engine CLI", parents=[parent_parser])
    parser.add_argument("--daemon-start", action="store_true", help="Iniciar daemon en segundo plano")
    parser.add_argument("--daemon-stop", action="store_true", help="Detener daemon vía Named Pipe")
    parser.add_argument("--daemon-status", action="store_true", help="Consultar estado y latencia del daemon")
    subparsers = parser.add_subparsers(dest="command", required=False)

    # status
    subparsers.add_parser("status", parents=[parent_parser], help="Estado del motor y libro activo")

    # get-csv
    get_csv_parser = subparsers.add_parser("get-csv", parents=[parent_parser], help="Extraer rango como texto CSV")
    get_csv_parser.add_argument("range", help="Rango A1 (ej. A1:Z100)")
    get_csv_parser.add_argument("--sheet", help="Pestaña de la hoja")
    get_csv_parser.add_argument("--delimiter", default=";", help="Delimitador (default: ';')")

    # get-ranges
    get_ranges_parser = subparsers.add_parser("get-ranges", parents=[parent_parser], help="Leer uno o más rangos de celdas")
    get_ranges_parser.add_argument("ranges", nargs="+", help="Rangos en notación A1")
    get_ranges_parser.add_argument("--sheet", help="Pestaña de la hoja")
    get_ranges_parser.add_argument("--formulas", action="store_true", help="Incluir fórmulas")

    # set-range (matriz 2D)
    set_range_parser = subparsers.add_parser("set-range", parents=[parent_parser], help="Escribir matriz de celdas")
    set_range_parser.add_argument("--start-cell", required=True, help="Celda superior izquierda (ej. A1)")
    set_range_parser.add_argument("--values-json", help="JSON array de arrays o @archivo.json")
    set_range_parser.add_argument("--input-file", help="Ruta a archivo JSON con matriz de datos")
    set_range_parser.add_argument("--stdin", action="store_true", help="Leer matriz JSON desde stdin")
    set_range_parser.add_argument("--sheet", help="Pestaña de la hoja")
    set_range_parser.add_argument("--autofit", action="store_true", help="Autoajustar ancho de columnas")

    # set-cells (mapa declarativo A1 -> cellData con recálculo atómico)
    set_cells_parser = subparsers.add_parser("set-cells", parents=[parent_parser], help="Escribir celdas declarativas A1 -> cellData")
    set_cells_parser.add_argument("--cells-json", help="JSON dict con mapa de celdas o @archivo.json")
    set_cells_parser.add_argument("--input-file", help="Ruta a archivo JSON con dict de celdas")
    set_cells_parser.add_argument("--stdin", action="store_true", help="Leer dict JSON desde stdin")
    set_cells_parser.add_argument("--sheet", help="Pestaña de la hoja")
    set_cells_parser.add_argument("--autofit", action="store_true", help="Autoajustar ancho de columnas")

    # copy-paste
    copy_paste_parser = subparsers.add_parser("copy-paste", parents=[parent_parser], help="Copiar y pegar rangos")
    copy_paste_parser.add_argument("--src", required=True, help="Rango origen")
    copy_paste_parser.add_argument("--dst", required=True, help="Celda destino")
    copy_paste_parser.add_argument("--src-sheet", help="Pestaña origen")
    copy_paste_parser.add_argument("--dst-sheet", help="Pestaña destino")
    copy_paste_parser.add_argument("--type", choices=["all", "values", "formats", "formulas"], default="all")

    # check-errors
    check_errors_parser = subparsers.add_parser("check-errors", parents=[parent_parser], help="Verificar errores de fórmulas (#VALUE!, #REF!)")
    check_errors_parser.add_argument("range", help="Rango a auditar")
    check_errors_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # audit-quality
    audit_quality_parser = subparsers.add_parser("audit-quality", parents=[parent_parser], help="Auditoría de calidad e higiene de datos")
    audit_quality_parser.add_argument("--range", help="Rango opcional a auditar (ej. A1:P701)")
    audit_quality_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # aggregate
    agg_parser = subparsers.add_parser("aggregate", parents=[parent_parser], help="Agrupación analítica (Pivot/Groupby)")
    agg_parser.add_argument("--group-by", required=True, help="Nombre de la columna para agrupar (ej. Segment)")
    agg_parser.add_argument("--metric", required=True, help="Nombre de la columna métrica numérica (ej. Profit)")
    agg_parser.add_argument("--func", default="sum", choices=["sum", "avg", "mean", "count", "min", "max"], help="Función de agregación (default: sum)")
    agg_parser.add_argument("--sheet", help="Pestaña de la hoja")
    agg_parser.add_argument("--top", type=int, default=10, help="Límite de grupos a mostrar (default: 10)")
    agg_parser.add_argument("--ascending", action="store_true", help="Ordenar de menor a mayor")

    # add-column (columna calculada con fórmulas vectorizadas)
    add_col_parser = subparsers.add_parser("add-column", parents=[parent_parser], help="Agregar columna calculada con fórmulas")
    add_col_parser.add_argument("--header", required=True, help="Nombre del encabezado de la nueva columna")
    add_col_parser.add_argument("--formula", required=True, help="Fórmula relativa base fila 2 (ej. =L2/J2)")
    add_col_parser.add_argument("--number-format", help="Máscara de formato numérico (ej. 0.0%% o $#,##0.00)")
    add_col_parser.add_argument("--sheet", help="Pestaña de la hoja")
    add_col_parser.add_argument("--no-autofit", action="store_true", help="Desactivar autoajuste de ancho de columna")

    # create-pivot (tabla dinámica)
    pivot_parser = subparsers.add_parser("create-pivot", parents=[parent_parser], help="Crear tabla dinámica (Pivot Table)")
    pivot_parser.add_argument("--source", required=True, help="Rango o nombre de tabla origen (ej. financials o A1:P701)")
    pivot_parser.add_argument("--rows", required=True, nargs="+", help="Campos de fila (ej. Country)")
    pivot_parser.add_argument("--cols", nargs="*", help="Campos de columna opcionales")
    pivot_parser.add_argument("--value-field", required=True, help="Campo métrico de valor (ej. Profit)")
    pivot_parser.add_argument("--value-func", default="sum", choices=["sum", "count", "average", "avg", "min", "max"], help="Función de agregación (default: sum)")
    pivot_parser.add_argument("--value-format", default=_DEFAULT_CURRENCY_FORMAT, help=f"Máscara de formato numérico (default: {_DEFAULT_CURRENCY_FORMAT})")
    pivot_parser.add_argument("--dest-sheet", help="Hoja de destino (creada automáticamente si no existe)")
    pivot_parser.add_argument("--dest-cell", default="A3", help="Celda superior izquierda destino (default: A3)")
    pivot_parser.add_argument("--table-name", help="Nombre identificador de la tabla dinámica")
    pivot_parser.add_argument("--source-sheet", help="Hoja origen si difiere de la activa")

    # create-chart (gráfico analítico anti-solapamiento)
    chart_parser = subparsers.add_parser("create-chart", parents=[parent_parser], help="Crear gráfico estadístico no invasivo")
    chart_parser.add_argument("--source", required=True, help="Rango origen o nombre de tabla dinámica")
    chart_parser.add_argument("--chart-type", default="column_clustered", choices=["column_clustered", "column", "bar_clustered", "bar", "line", "pie", "area"], help="Tipo de gráfico")
    chart_parser.add_argument("--title", help="Título del gráfico")
    chart_parser.add_argument("--dest-sheet", help="Hoja destino donde se inserta el gráfico")
    chart_parser.add_argument("--placement", default="auto", choices=["auto", "right", "bottom"], help="Estrategia geométrica de ubicación")
    chart_parser.add_argument("--anchor", help="Celda específica de anclaje (ej. I3)")
    chart_parser.add_argument("--width", type=int, default=480, help="Ancho en píxeles (default: 480)")
    chart_parser.add_argument("--height", type=int, default=300, help="Alto en píxeles (default: 300)")
    chart_parser.add_argument("--source-sheet", help="Hoja origen si difiere")

    # apply-preset (presets de estilo universales)
    preset_parser = subparsers.add_parser("apply-preset", parents=[parent_parser], help="Aplicar preset de estilo universal")
    preset_parser.add_argument("--range", required=True, help="Rango A1 (ej. B2:D2 o C5)")
    preset_parser.add_argument("--preset", required=True, choices=["header", "input_cell", "total_row", "currency", "percent", "delta_positive", "delta_negative"], help="Preset de estilo")
    preset_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # summary-table (motor genérico de tablas de resumen / comparativas)
    summary_parser = subparsers.add_parser("summary-table", parents=[parent_parser], help="Crear tabla resumen o comparativa con parámetros")
    summary_parser.add_argument("--title", help="Título principal de la tabla")
    summary_parser.add_argument("--headers", nargs="+", help="Encabezados de columnas")
    summary_parser.add_argument("--rows-json", help="JSON con lista de especificaciones de fila o @archivo.json")
    summary_parser.add_argument("--parameters-json", help="JSON con lista de parámetros/supuestos o @archivo.json")
    summary_parser.add_argument("--spec-json", help="JSON con especificación completa o @archivo.json")
    summary_parser.add_argument("--input-file", help="Ruta a archivo JSON con filas o spec completa")
    summary_parser.add_argument("--stdin", action="store_true", help="Leer filas o spec JSON desde stdin")
    summary_parser.add_argument("--dest-sheet", help="Pestaña de destino")
    summary_parser.add_argument("--start-cell", default="B2", help="Celda superior izquierda (default: B2)")

    # dedup (eliminación de duplicados on-demand)
    dedup_parser = subparsers.add_parser("dedup", parents=[parent_parser], help="Eliminar filas duplicadas on-demand")
    dedup_parser.add_argument("--source", required=True, help="Rango o nombre de tabla (ej. financials o A1:P716)")
    dedup_parser.add_argument("--key-cols", nargs="*", help="Columnas clave para evaluar duplicados")
    dedup_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # flag-outliers (detección y marcado visual no destructivo de outliers on-demand)
    outliers_parser = subparsers.add_parser("flag-outliers", parents=[parent_parser], help="Detectar y marcar valores atípicos on-demand")
    outliers_parser.add_argument("--source", help="Rango o nombre de tabla (default: auto)")
    outliers_parser.add_argument("--col", required=True, help="Nombre de columna a evaluar")
    outliers_parser.add_argument("--method", choices=["iqr", "zscore"], default="iqr", help="Método de detección (iqr o zscore, default: iqr)")
    outliers_parser.add_argument("--no-comment", action="store_true", help="No agregar comentarios explicativos a las celdas")
    outliers_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # clean (limpieza, estandarización de texto, fechas y casteo de tipos on-demand)
    clean_parser = subparsers.add_parser("clean", parents=[parent_parser], help="Limpiar y estandarizar dataset on-demand")
    clean_parser.add_argument("--source", required=True, help="Rango o nombre de tabla a limpiar (ej. financials)")
    clean_parser.add_argument("--rules-json", help="JSON de reglas o @archivo.json")
    clean_parser.add_argument("--input-file", help="Ruta a archivo JSON con reglas de limpieza")
    clean_parser.add_argument("--stdin", action="store_true", help="Leer reglas JSON desde stdin")
    clean_parser.add_argument("--sheet", help="Pestaña de la hoja")

    # curate (pipeline unificado de curaduría: dedup + clean + outliers en una sola pasada)
    curate_parser = subparsers.add_parser("curate", parents=[parent_parser], help="Pipeline unificado de curaduría (dedup + clean + outliers)")
    curate_parser.add_argument("--source", required=True, help="Rango o nombre de tabla a procesar (ej. financials)")
    curate_parser.add_argument("--spec-json", help="JSON de especificación de pipeline o @archivo.json")
    curate_parser.add_argument("--input-file", help="Ruta a archivo JSON con especificación de pipeline")
    curate_parser.add_argument("--stdin", action="store_true", help="Leer especificación JSON desde stdin")
    curate_parser.add_argument("--sheet", help="Pestaña de la hoja")

    args = parser.parse_args()

    def output(data):
        if args.json:
            print(json.dumps(data, indent=2, ensure_ascii=False))
        else:
            if isinstance(data, dict) or isinstance(data, list):
                print(json.dumps(data, indent=2, ensure_ascii=False))
            else:
                print(data)

    if getattr(args, "daemon_start", False):
        try:
            from antigravity_excel_daemon import start_daemon_process
            res = start_daemon_process(mode=args.mode, file_path=args.file)
        except Exception as e:
            res = {"status": "error", "error": str(e)}
        output(res)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if getattr(args, "daemon_stop", False):
        try:
            from antigravity_excel_daemon import stop_daemon_process
            res = stop_daemon_process()
        except Exception as e:
            res = {"status": "error", "error": str(e)}
        output(res)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if getattr(args, "daemon_status", False):
        try:
            from antigravity_excel_daemon import status_daemon
            res = status_daemon()
        except Exception as e:
            res = {"status": "error", "error": str(e)}
        output(res)
        sys.exit(0 if res.get("status") in ("success", "info") else 1)

    if not args.command:
        parser.print_help()
        sys.exit(1)

    # Pre-carga de datos para comandos estructurados
    values = None
    cells = None
    spec = None
    rules = None
    rows_spec = None
    title = None
    headers = None
    dest_sheet = None
    start_cell = None
    parameters = None

    if args.command == "set-range":
        if args.stdin:
            values = json.loads(sys.stdin.read())
        elif args.input_file:
            with open(args.input_file, "r", encoding="utf-8") as f:
                values = json.load(f)
        elif args.values_json:
            values = load_json_or_file(args.values_json)
        if not values:
            raise ValueError("Debe proporcionar matriz de valores vía --values-json, --input-file o --stdin")

    elif args.command == "set-cells":
        if args.stdin:
            cells = json.loads(sys.stdin.read())
        elif args.input_file:
            with open(args.input_file, "r", encoding="utf-8") as f:
                cells = json.load(f)
        elif args.cells_json:
            cells = load_json_or_file(args.cells_json)
        if not cells:
            raise ValueError("Debe proporcionar mapa de celdas vía --cells-json, --input-file o --stdin")

    elif args.command == "summary-table":
        if args.stdin:
            spec = json.loads(sys.stdin.read())
        elif args.input_file:
            with open(args.input_file, "r", encoding="utf-8") as f:
                spec = json.load(f)
        elif args.spec_json:
            spec = load_json_or_file(args.spec_json)

        if isinstance(spec, dict) and ("rows_spec" in spec or "rows" in spec):
            title = spec.get("title", args.title or "Resumen")
            headers = spec.get("headers", args.headers or [])
            rows_spec = spec.get("rows_spec") or spec.get("rows", [])
            dest_sheet = spec.get("dest_sheet", args.dest_sheet or "Resumen")
            start_cell = spec.get("start_cell", args.start_cell or "B2")
            parameters = spec.get("parameters", None)
        else:
            rows_spec = spec if isinstance(spec, list) else None
            if not rows_spec and args.rows_json:
                rows_spec = load_json_or_file(args.rows_json)
            if not rows_spec:
                raise ValueError("Debe proporcionar especificación de filas vía --rows-json, --spec-json, --input-file o --stdin")
            title = args.title or "Resumen"
            headers = args.headers or []
            dest_sheet = args.dest_sheet or "Resumen"
            start_cell = args.start_cell or "B2"
            parameters = load_json_or_file(args.parameters_json) if args.parameters_json else None

    elif args.command == "clean":
        if args.stdin:
            rules = json.loads(sys.stdin.read())
        elif args.input_file:
            with open(args.input_file, "r", encoding="utf-8") as f:
                rules = json.load(f)
        elif args.rules_json:
            rules = load_json_or_file(args.rules_json)
        if not rules or not isinstance(rules, dict):
            raise ValueError("Debe proporcionar un diccionario de reglas vía --rules-json, --input-file o --stdin")

    elif args.command == "curate":
        if args.stdin:
            spec = json.loads(sys.stdin.read())
        elif args.input_file:
            with open(args.input_file, "r", encoding="utf-8") as f:
                spec = json.load(f)
        elif args.spec_json:
            spec = load_json_or_file(args.spec_json)
        if not spec or not isinstance(spec, dict):
            raise ValueError("Debe proporcionar una especificación de pipeline vía --spec-json, --input-file o --stdin")

    # Intentar llamada ultra-rápida al daemon residente (< 30ms)
    payload_args = {"mode": args.mode, "file": args.file}
    if args.command == "status":
        pass
    elif args.command == "get-csv":
        payload_args.update({"range": args.range, "sheet": args.sheet, "delimiter": args.delimiter})
    elif args.command == "get-ranges":
        payload_args.update({"ranges": args.ranges, "sheet": args.sheet, "formulas": args.formulas})
    elif args.command == "set-range":
        payload_args.update({"start_cell": args.start_cell, "values": values, "sheet": args.sheet, "autofit": args.autofit})
    elif args.command == "set-cells":
        payload_args.update({"cells": cells, "sheet": args.sheet, "autofit": args.autofit})
    elif args.command == "copy-paste":
        payload_args.update({"src": args.src, "dst": args.dst, "src_sheet": args.src_sheet, "dst_sheet": args.dst_sheet, "type": args.type})
    elif args.command == "check-errors":
        payload_args.update({"range": args.range, "sheet": args.sheet})
    elif args.command == "audit-quality":
        payload_args.update({"range": args.range, "sheet": args.sheet})
    elif args.command == "aggregate":
        payload_args.update({"group_by": args.group_by, "metric": args.metric, "func": args.func, "sheet": args.sheet, "top": args.top, "ascending": args.ascending})
    elif args.command == "add-column":
        payload_args.update({"header": args.header, "formula": args.formula, "number_format": args.number_format, "sheet": args.sheet, "autofit": not args.no_autofit})
    elif args.command == "create-pivot":
        payload_args.update({"source": args.source, "rows": args.rows, "cols": args.cols, "value_field": args.value_field, "value_func": args.value_func, "value_format": args.value_format, "dest_sheet": args.dest_sheet, "dest_cell": args.dest_cell, "table_name": args.table_name, "source_sheet": args.source_sheet})
    elif args.command == "create-chart":
        payload_args.update({"source": args.source, "chart_type": args.chart_type, "title": args.title, "dest_sheet": args.dest_sheet, "placement": args.placement, "anchor": args.anchor, "width": args.width, "height": args.height, "source_sheet": args.source_sheet})
    elif args.command == "apply-preset":
        payload_args.update({"range": args.range, "preset": args.preset, "sheet": args.sheet})
    elif args.command == "summary-table":
        payload_args.update({"title": title, "headers": headers, "rows_spec": rows_spec, "dest_sheet": dest_sheet, "start_cell": start_cell, "parameters": parameters})
    elif args.command == "dedup":
        payload_args.update({"source": args.source, "key_cols": args.key_cols, "sheet": args.sheet})
    elif args.command == "flag-outliers":
        payload_args.update({"col": args.col, "method": args.method, "source": args.source, "sheet": args.sheet, "no_comment": args.no_comment})
    elif args.command == "clean":
        payload_args.update({"source": args.source, "rules": rules, "sheet": args.sheet})
    elif args.command == "curate":
        payload_args.update({"source": args.source, "spec": spec, "sheet": args.sheet})

    daemon_resp = _try_call_daemon(args.command, payload_args)
    if daemon_resp is not None:
        if daemon_resp.get("status") == "success":
            res = daemon_resp.get("result")
            if args.command == "get-csv" and not args.json:
                print(res.get("csv") if isinstance(res, dict) else res)
            else:
                output(res)
            sys.exit(0)
        elif daemon_resp.get("status") == "error":
            err = daemon_resp.get("error", "Error desconocido del daemon")
            if args.json:
                print(json.dumps({"status": "error", "error": err}))
            else:
                print(f"Error: {err}")
            sys.exit(1)

    # Fallback automático: ejecución standalone local
    try:
        from antigravity_excel_core import AntigravityExcelEngine
        engine = AntigravityExcelEngine(mode=args.mode, file_path=args.file)
    except Exception as e:
        if args.json:
            print(json.dumps({"status": "error", "error": f"Error inicializando motor: {str(e)}"}))
        else:
            print(f"Error inicializando motor: {str(e)}")
        sys.exit(1)

    try:
        if args.command == "status":
            output(engine.status())

        elif args.command == "get-csv":
            csv_data = engine.get_range_as_csv(args.range, sheet=args.sheet, delimiter=args.delimiter)
            if args.json:
                output({"status": "success", "csv": csv_data})
            else:
                print(csv_data)

        elif args.command == "get-ranges":
            res = engine.get_cell_ranges(args.ranges, sheet=args.sheet, include_formulas=args.formulas)
            output({"status": "success", "ranges": res})

        elif args.command == "set-range":
            if values is None:
                if args.stdin:
                    values = json.loads(sys.stdin.read())
                elif args.input_file:
                    with open(args.input_file, "r", encoding="utf-8") as f:
                        values = json.load(f)
                elif args.values_json:
                    values = load_json_or_file(args.values_json)

            if not values:
                raise ValueError("Debe proporcionar matriz de valores vía --values-json, --input-file o --stdin")

            res = engine.set_cell_range(args.start_cell, values, sheet=args.sheet, autofit=args.autofit)
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "set-cells":
            if cells is None:
                if args.stdin:
                    cells = json.loads(sys.stdin.read())
                elif args.input_file:
                    with open(args.input_file, "r", encoding="utf-8") as f:
                        cells = json.load(f)
                elif args.cells_json:
                    cells = load_json_or_file(args.cells_json)

            if not cells:
                raise ValueError("Debe proporcionar mapa de celdas vía --cells-json, --input-file o --stdin")

            res = engine.set_cells(cells, sheet=args.sheet, autofit=args.autofit)
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "copy-paste":
            res = engine.copy_paste_range(args.src, args.dst, args.src_sheet, args.dst_sheet, args.type)
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "check-errors":
            errs = engine.check_formula_errors(args.range, sheet=args.sheet)
            output({"status": "success", "errors_count": len(errs), "errors": errs})

        elif args.command == "audit-quality":
            report = engine.audit_data_quality(range_str=args.range, sheet=args.sheet)
            output({"status": "success", "audit_report": report})

        elif args.command == "aggregate":
            res = engine.aggregate(
                group_by_col=args.group_by,
                metric_col=args.metric,
                agg_func=args.func,
                sheet=args.sheet,
                top_n=args.top,
                ascending=args.ascending
            )
            output({"status": "success", "aggregation": res})

        elif args.command == "add-column":
            res = engine.add_calculated_column(
                header=args.header,
                formula_template=args.formula,
                number_format=args.number_format,
                sheet=args.sheet,
                autofit=not args.no_autofit
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "create-pivot":
            res = engine.create_pivot_table(
                source=args.source,
                rows=args.rows,
                cols=args.cols,
                value_field=args.value_field,
                value_func=args.value_func,
                value_format=args.value_format,
                dest_sheet=args.dest_sheet,
                dest_cell=args.dest_cell,
                table_name=args.table_name,
                source_sheet=args.source_sheet
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "create-chart":
            res = engine.create_chart(
                source=args.source,
                chart_type=args.chart_type,
                title=args.title,
                dest_sheet=args.dest_sheet,
                placement=args.placement,
                anchor=args.anchor,
                width=args.width,
                height=args.height,
                source_sheet=args.source_sheet
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "apply-preset":
            res = engine.apply_style_preset(args.range, preset=args.preset, sheet=args.sheet)
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "summary-table":
            if rows_spec is None:
                if args.stdin:
                    spec = json.loads(sys.stdin.read())
                elif args.input_file:
                    with open(args.input_file, "r", encoding="utf-8") as f:
                        spec = json.load(f)
                elif args.spec_json:
                    spec = load_json_or_file(args.spec_json)

                if isinstance(spec, dict) and ("rows_spec" in spec or "rows" in spec):
                    title = spec.get("title", args.title or "Resumen")
                    headers = spec.get("headers", args.headers or [])
                    rows_spec = spec.get("rows_spec") or spec.get("rows", [])
                    dest_sheet = spec.get("dest_sheet", args.dest_sheet or "Resumen")
                    start_cell = spec.get("start_cell", args.start_cell or "B2")
                    parameters = spec.get("parameters", None)
                else:
                    rows_spec = spec if isinstance(spec, list) else None
                    if not rows_spec and args.rows_json:
                        rows_spec = load_json_or_file(args.rows_json)
                    if not rows_spec:
                        raise ValueError("Debe proporcionar especificación de filas vía --rows-json, --spec-json, --input-file o --stdin")
                    title = args.title or "Resumen"
                    headers = args.headers or []
                    dest_sheet = args.dest_sheet or "Resumen"
                    start_cell = args.start_cell or "B2"
                    parameters = load_json_or_file(args.parameters_json) if args.parameters_json else None

            res = engine.create_summary_table(
                title=title,
                headers=headers,
                rows_spec=rows_spec,
                dest_sheet=dest_sheet,
                start_cell=start_cell,
                parameters=parameters
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "dedup":
            res = engine.remove_duplicates(
                range_or_table=args.source,
                key_columns=args.key_cols,
                sheet=args.sheet
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "flag-outliers":
            res = engine.detect_and_flag_outliers(
                column=args.col,
                method=args.method,
                range_or_table=args.source,
                sheet=args.sheet,
                add_comment=not args.no_comment
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "clean":
            if rules is None:
                if args.stdin:
                    rules = json.loads(sys.stdin.read())
                elif args.input_file:
                    with open(args.input_file, "r", encoding="utf-8") as f:
                        rules = json.load(f)
                elif args.rules_json:
                    rules = load_json_or_file(args.rules_json)
            if not rules or not isinstance(rules, dict):
                raise ValueError("Debe proporcionar un diccionario de reglas vía --rules-json, --input-file o --stdin")

            res = engine.clean_table_dataset(
                target_range_or_table=args.source,
                rules=rules,
                sheet=args.sheet
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)

        elif args.command == "curate":
            if spec is None:
                if args.stdin:
                    spec = json.loads(sys.stdin.read())
                elif args.input_file:
                    with open(args.input_file, "r", encoding="utf-8") as f:
                        spec = json.load(f)
                elif args.spec_json:
                    spec = load_json_or_file(args.spec_json)
            if not spec or not isinstance(spec, dict):
                raise ValueError("Debe proporcionar una especificación de pipeline vía --spec-json, --input-file o --stdin")

            res = engine.curate_pipeline(
                target_range_or_table=args.source,
                pipeline_spec=spec,
                sheet=args.sheet
            )
            saved_to = engine.save()
            res["saved_to"] = saved_to
            output(res)


    except Exception as e:
        if args.json:
            print(json.dumps({"status": "error", "error": str(e)}))
        else:
            print(f"Error: {str(e)}")
        sys.exit(1)
    finally:
        if args.mode == "file" or (args.file and not engine.status().get("backend") == "live_com"):
            try:
                engine.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
