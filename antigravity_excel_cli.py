# Creado por Francisco Grandón Vergara
"""
CLI Oficial de Antigravity Excel Engine.
Proporciona salida JSON determinista para subagentes y llamadas por lotes.
"""

import argparse
import json
import sys
import os
from typing import Any
from antigravity_excel_core import AntigravityExcelEngine


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
    parent_parser = argparse.ArgumentParser(add_help=False)
    parent_parser.add_argument("--json", action="store_true", help="Salida en formato JSON")
    parent_parser.add_argument("--mode", choices=["auto", "live", "file"], default="auto", help="Modo de ejecución")
    parent_parser.add_argument("--file", help="Ruta al archivo .xlsx")

    parser = argparse.ArgumentParser(description="Antigravity Excel Engine CLI", parents=[parent_parser])
    subparsers = parser.add_subparsers(dest="command", required=True)

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

    args = parser.parse_args()

    def output(data):
        if args.json:
            print(json.dumps(data, indent=2, ensure_ascii=False))
        else:
            if isinstance(data, dict) or isinstance(data, list):
                print(json.dumps(data, indent=2, ensure_ascii=False))
            else:
                print(data)

    try:
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
            values = None
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
            cells = None
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
