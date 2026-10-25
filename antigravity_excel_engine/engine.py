# Creado por Francisco Grandón Vergara
"""
AntigravityExcelEngine — Orquestador central del motor Excel.

Responsabilidades exclusivas:
  - Seleccionar e instanciar el backend correcto (live COM o headless).
  - Exponer la API pública delegando a los submódulos de operations/.
  - Nunca ejecutar lógica de negocio de Excel directamente.

Compatibilidad total con la API pública de antigravity_excel_core.py.
"""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Union

from antigravity_excel_engine.backends import (
    BaseExcelBackend,
    HeadlessOpenPyXLBackend,
    LiveExcelCOMBackend,
)
from antigravity_excel_engine.operations import (
    add_calculated_column,
    aggregate,
    audit_data_quality,
    check_formula_errors,
    clean_table_dataset,
    copy_paste_range,
    create_chart,
    create_pivot_table,
    create_summary_table,
    create_table,
    curate_pipeline,
    detect_and_flag_outliers,
    extend_chart_series,
    get_ai_grounding_context,
    get_range_as_csv,
    inspect_charts,
    inspect_table_schema,
    inspect_workbook,
    list_tables,
    read_ranges,
    remove_duplicates,
    resize_table,
    update_chart_series,
    write_cell_range,
    write_cells_declarative,
)
from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("engine")


class AntigravityExcelEngine:
    """
    Fachada y coordinador central del motor de automatización Excel.

    Selecciona automáticamente el backend óptimo según el entorno:
      - "live"  → LiveExcelCOMBackend (Excel COM activo en pantalla)
      - "file"  → HeadlessOpenPyXLBackend (OpenPyXL en disco, sin GUI)
      - "auto"  → Detecta si hay Excel activo; si no, usa headless.

    Todos los métodos de negocio delegan a operations/* o al backend directamente.
    """

    def __init__(
        self,
        mode: str = "auto",
        file_path: Optional[str] = None,
        visible: bool = True,
        **backend_kwargs: Any,
    ) -> None:
        """
        Inicializa el motor con el backend seleccionado.

        Args:
            mode: "auto", "live" o "file".
            file_path: Ruta al archivo Excel (.xlsx, .xlsb, etc.).
            visible: Si True, Excel se abre visible en modo live.
            **backend_kwargs: Argumentos adicionales para el backend específico.
        """
        self.mode = mode.lower()
        self.file_path = file_path
        self.visible = visible
        self._backend_kwargs = backend_kwargs
        self.backend: BaseExcelBackend = self._select_backend()
        logger.info(
            "AntigravityExcelEngine iniciado | mode=%s backend=%s file=%s",
            self.mode, type(self.backend).__name__, file_path
        )

    # ─────────────────────────── Backend Selection ──────────────────────────

    def _select_backend(self) -> BaseExcelBackend:
        """Instancia el backend apropiado según mode y entorno disponible."""
        if self.mode == "live":
            return LiveExcelCOMBackend(self.file_path)
        elif self.mode == "file":
            return HeadlessOpenPyXLBackend(self.file_path)
        elif self.mode == "auto":
            return self._auto_detect_backend()
        else:
            raise ValueError(
                f"Modo no reconocido: '{self.mode}'. Use 'auto', 'live' o 'file'."
            )

    def _auto_detect_backend(self) -> BaseExcelBackend:
        """Detecta automáticamente si Excel COM está disponible."""
        # 1. Verificar si Excel está en la ROT activa
        try:
            import win32com.client  # type: ignore
            win32com.client.GetActiveObject("Excel.Application")
            logger.debug("auto-detect: Excel COM activo (ROT) → LiveExcelCOMBackend")
            return LiveExcelCOMBackend(self.file_path)
        except Exception:
            pass

        # 2. Verificar si hay ventana de Excel activa en el escritorio vía oleacc
        try:
            tester = LiveExcelCOMBackend.__new__(LiveExcelCOMBackend)
            if hasattr(tester, "_try_get_live_desktop_excel"):
                desktop_app = tester._try_get_live_desktop_excel(self.file_path)
                if desktop_app:
                    logger.debug("auto-detect: Excel en escritorio (oleacc) → LiveExcelCOMBackend")
                    return LiveExcelCOMBackend(self.file_path)
        except Exception:
            pass

        try:
            import openpyxl  # noqa: F401  # type: ignore
            logger.debug("auto-detect: sin COM → HeadlessOpenPyXLBackend")
            return HeadlessOpenPyXLBackend(self.file_path)
        except ImportError:
            pass

        try:
            import win32com.client  # type: ignore
            logger.debug("auto-detect: sin openpyxl → LiveExcelCOMBackend fallback")
            return LiveExcelCOMBackend(self.file_path)
        except ImportError:
            pass

        raise ImportError(
            "No se encontró backend disponible. Instale pywin32 u openpyxl."
        )

    # ─────────────────────────── Control ────────────────────────────────────

    def status(self) -> Dict[str, Any]:
        """Devuelve el estado actual del motor y del backend."""
        return self.backend.status()

    def wait_until_ready(
        self,
        timeout_sec: float = 10.0,
        retry_interval: float = 0.25,
    ) -> Dict[str, Any]:
        """Espera hasta que el backend esté listo para operar."""
        return self.backend.wait_until_ready(
            timeout_sec=timeout_sec, retry_interval=retry_interval
        )

    def save(self, file_path: Optional[str] = None) -> str:
        """Guarda el libro de trabajo."""
        return self.backend.save(file_path=file_path)

    def close(self) -> None:
        """Cierra el libro de trabajo y libera recursos."""
        self.backend.close()

    def dispatch_com(
        self,
        target: str,
        member: str,
        args: Optional[List[Any]] = None,
        kwargs: Optional[Dict[str, Any]] = None,
    ) -> Dict[str, Any]:
        """Despacha una llamada COM avanzada al backend live (solo modo live)."""
        return self.backend.dispatch_com(target=target, member=member, args=args, kwargs=kwargs)

    # ─────────────────────────── Read ───────────────────────────────────────

    def get_cell_ranges(
        self,
        ranges: List[str],
        sheet: Optional[str] = None,
        include_formulas: bool = False,
        include_styles: bool = False,
    ) -> Dict[str, Any]:
        """Lee uno o varios rangos de celdas."""
        return read_ranges(
            self.backend, ranges=ranges, sheet=sheet,
            include_formulas=include_formulas, include_styles=include_styles
        )

    def get_range_as_csv(
        self,
        range_str: str,
        sheet: Optional[str] = None,
        delimiter: str = ";",
        include_headers: bool = True,
    ) -> str:
        """Exporta un rango como texto CSV."""
        return get_range_as_csv(
            self.backend, range_str=range_str, sheet=sheet,
            delimiter=delimiter, include_headers=include_headers
        )

    # ─────────────────────────── Write ──────────────────────────────────────

    def set_cell_range(
        self,
        start_cell: str,
        values: List[List[Any]],
        sheet: Optional[str] = None,
        formulas: Optional[List[List[str]]] = None,
        styles: Optional[Dict[str, Any]] = None,
        autofit: bool = False,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Escribe una matriz rectangular de valores a partir de una celda de inicio."""
        return write_cell_range(
            self.backend, start_cell=start_cell, values=values,
            sheet=sheet, formulas=formulas, styles=styles, autofit=autofit, dry_run=dry_run
        )

    def set_cells(
        self,
        cells_map: Dict[str, Dict[str, Any]],
        sheet: Optional[str] = None,
        autofit: bool = False,
        dry_run: bool = False,
    ) -> Dict[str, Any]:
        """Escritura declarativa de celdas individuales."""
        return write_cells_declarative(
            self.backend, cells_map=cells_map, sheet=sheet, autofit=autofit, dry_run=dry_run
        )

    def copy_paste_range(
        self,
        source_range: str,
        dest_cell: str,
        source_sheet: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        paste_values_only: bool = False,
    ) -> Dict[str, Any]:
        """Copia un rango de celdas a una ubicación destino."""
        return copy_paste_range(
            self.backend, source_range=source_range, dest_cell=dest_cell,
            source_sheet=source_sheet, dest_sheet=dest_sheet,
            paste_values_only=paste_values_only
        )

    # ─────────────────────────── Audit ──────────────────────────────────────

    def audit_data_quality(
        self,
        range_str: Optional[str] = None,
        sheet: Optional[str] = None,
        checks: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Audita la calidad de datos en un rango u hoja."""
        return audit_data_quality(
            self.backend, range_str=range_str, sheet=sheet, checks=checks
        )

    def check_formula_errors(
        self,
        range_str: Optional[str] = None,
        sheet: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Detecta celdas con errores de fórmula."""
        return check_formula_errors(
            self.backend, range_str=range_str, sheet=sheet
        )

    def detect_and_flag_outliers(
        self,
        range_str: str,
        column: Optional[str] = None,
        method: str = "iqr",
        flag_style: str = "warning",
        sheet: Optional[str] = None,
        threshold: float = 1.5,
    ) -> Dict[str, Any]:
        """Detecta y marca visualmente valores atípicos."""
        return detect_and_flag_outliers(
            self.backend, range_str=range_str, column=column,
            method=method, flag_style=flag_style, sheet=sheet, threshold=threshold
        )

    # ─────────────────────────── Aggregate ──────────────────────────────────

    def aggregate(
        self,
        range_str: str = "",
        group_by_col: str = "",
        metric_col: str = "",
        agg_func: str = "sum",
        top_n: int = 10,
        sheet: Optional[str] = None,
        dest_cell: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Agrupa filas y calcula una métrica agregada."""
        return aggregate(
            self.backend, range_str=range_str, group_by_col=group_by_col,
            metric_col=metric_col, agg_func=agg_func, top_n=top_n,
            sheet=sheet, dest_cell=dest_cell, dest_sheet=dest_sheet, **kwargs
        )

    def add_calculated_column(
        self,
        header: str = "",
        formula_template: str = "",
        sheet: Optional[str] = None,
        col_format: Optional[str] = None,
        autofit: bool = True,
        dry_run: bool = False,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Agrega una columna calculada con fórmula dinámica."""
        return add_calculated_column(
            self.backend, header=header, formula_template=formula_template,
            sheet=sheet, col_format=col_format, autofit=autofit, dry_run=dry_run, **kwargs
        )

    # ─────────────────────────── Curation ───────────────────────────────────

    def remove_duplicates(
        self,
        range_str: str = "",
        key_columns: Optional[List[str]] = None,
        sheet: Optional[str] = None,
        keep: str = "first",
        range_or_table: Optional[str] = None,
        **kwargs: Any,
    ) -> Dict[str, Any]:
        """Elimina filas duplicadas en el rango especificado."""
        return remove_duplicates(
            self.backend, range_str=range_str, key_columns=key_columns,
            sheet=sheet, keep=keep, range_or_table=range_or_table, **kwargs
        )

    def clean_table_dataset(
        self,
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
        """Limpia y normaliza un rango tabular de datos."""
        return clean_table_dataset(
            self.backend, range_str=range_str, sheet=sheet, rules=rules,
            target_range_or_table=target_range_or_table,
            date_columns=date_columns, numeric_columns=numeric_columns,
            trim_strings=trim_strings, normalize_booleans=normalize_booleans, **kwargs
        )

    def curate_pipeline(
        self,
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
        """Ejecuta un pipeline completo de curación de datos."""
        return curate_pipeline(
            self.backend, range_str=range_str, sheet=sheet,
            pipeline_spec=pipeline_spec, target_range_or_table=target_range_or_table,
            steps=steps, key_columns=key_columns, date_columns=date_columns,
            numeric_columns=numeric_columns, outlier_column=outlier_column,
            outlier_method=outlier_method, **kwargs
        )

    # ─────────────────────────── Charts / Pivots ────────────────────────────

    def create_pivot_table(
        self,
        source: Optional[Union[str, Dict[str, Any]]] = None,
        source_range: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        dest_cell: str = "A3",
        rows: Optional[List[str]] = None,
        cols: Optional[List[str]] = None,
        value_field: Optional[str] = None,
        value_func: str = "sum",
        table_name: Optional[str] = None,
        source_sheet: Optional[str] = None,
        value_format: Optional[str] = None,
        **kwargs: Any
    ) -> Dict[str, Any]:
        """Crea una tabla dinámica a partir de un rango o tabla de datos."""
        src = source if source is not None else source_range
        return create_pivot_table(
            self.backend, source=src, dest_sheet=dest_sheet,
            dest_cell=dest_cell, rows=rows, cols=cols, value_field=value_field,
            value_func=value_func, table_name=table_name, source_sheet=source_sheet,
            value_format=value_format, **kwargs
        )

    def create_chart(
        self,
        source: Optional[Union[str, Dict[str, Any]]] = None,
        source_range: Optional[str] = None,
        chart_type: str = "column_clustered",
        dest_sheet: Optional[str] = None,
        dest_cell: Optional[str] = None,
        title: Optional[str] = None,
        x_axis: Optional[str] = None,
        y_axis: Optional[str] = None,
        sheet: Optional[str] = None,
        width: Optional[int] = None,
        height: Optional[int] = None,
        placement: str = "auto",
        anchor: Optional[str] = None,
        **kwargs: Any
    ) -> Dict[str, Any]:
        """Crea un gráfico Excel a partir de un rango o tabla de datos."""
        src = source if source is not None else source_range
        return create_chart(
            self.backend, source=src, chart_type=chart_type,
            dest_sheet=dest_sheet, dest_cell=dest_cell, title=title,
            x_axis=x_axis, y_axis=y_axis, sheet=sheet, width=width, height=height,
            placement=placement, anchor=anchor, **kwargs
        )

    def create_summary_table(
        self,
        source_range: str,
        metrics: Optional[List[str]] = None,
        dest_sheet: Optional[str] = None,
        dest_cell: Optional[str] = None,
        sheet: Optional[str] = None,
        title: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Crea una tabla de resumen estadístico."""
        return create_summary_table(
            self.backend, source_range=source_range, metrics=metrics,
            dest_sheet=dest_sheet, dest_cell=dest_cell, sheet=sheet, title=title
        )

    # ─────────────────────────── Estilo ─────────────────────────────────────

    def apply_style_preset(
        self,
        range_str: str,
        preset: Optional[str] = None,
        sheet: Optional[str] = None,
        preset_name: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Aplica un preset de estilo financiero a un rango de celdas."""
        p = preset or preset_name or "header"
        return self.backend.apply_style_preset(
            range_str=range_str, preset=p, sheet=sheet
        )

    # ─────────────────────────── Introspección y Gráficos Avanzados ──────────

    def inspect_workbook(
        self,
        sheet: Optional[str] = None,
        include_charts: bool = True,
    ) -> Dict[str, Any]:
        """Inspecciona metadatos globales del libro y estructura detallada por hoja."""
        return inspect_workbook(self.backend, sheet=sheet, include_charts=include_charts)

    def inspect_charts(
        self,
        sheet: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Inspecciona todos los gráficos en la hoja especificada o en el libro completo."""
        return inspect_charts(self.backend, sheet=sheet)

    def extend_chart_series(
        self,
        to_col: Optional[str] = None,
        to_row: Optional[int] = None,
        sheet: Optional[str] = None,
        charts: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """Extiende las series y rangos de categorías de los gráficos indicados."""
        return extend_chart_series(
            self.backend, to_col=to_col, to_row=to_row, sheet=sheet, charts=charts
        )

    def update_chart_series(
        self,
        chart_name: str,
        series_index: int,
        x_range: Optional[str] = None,
        y_range: Optional[str] = None,
        name_ref: Optional[str] = None,
        formula: Optional[str] = None,
        sheet: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Modifica directamente los parámetros o la fórmula de una serie de gráfico."""
        return update_chart_series(
            self.backend,
            chart_name=chart_name,
            series_index=series_index,
            x_range=x_range,
            y_range=y_range,
            name_ref=name_ref,
            formula=formula,
            sheet=sheet,
        )

    # ─────────────────────────── Ejecución Batch Atómica ────────────────────

    # ─────────────────────────── Copilot & Tables ─────────────────────────────

    def create_table(
        self,
        range_str: str,
        table_name: Optional[str] = None,
        style: str = "TableStyleMedium9",
        has_headers: bool = True,
        sheet: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Convierte un rango libre en tabla oficial ListObject."""
        return create_table(
            self.backend,
            range_str=range_str,
            table_name=table_name,
            style=style,
            has_headers=has_headers,
            sheet=sheet,
        )

    def list_tables(self, sheet: Optional[str] = None) -> List[Dict[str, Any]]:
        """Lista todas las tablas estructuradas en la hoja especificada o en el libro."""
        return list_tables(self.backend, sheet=sheet)

    def resize_table(
        self,
        table_name: str,
        new_range: str,
        sheet: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Redimensiona el rango de una tabla estructurada existente."""
        return resize_table(
            self.backend,
            table_name=table_name,
            new_range=new_range,
            sheet=sheet,
        )

    def inspect_table_schema(
        self,
        table_or_range: str,
        sheet: Optional[str] = None,
        max_sample_rows: int = 1000,
    ) -> Dict[str, Any]:
        """Inspecciona y describe exhaustivamente el esquema de una tabla o rango."""
        return inspect_table_schema(
            self.backend,
            table_or_range=table_or_range,
            sheet=sheet,
            max_sample_rows=max_sample_rows,
        )

    def get_ai_grounding_context(
        self,
        table_or_range: str,
        sheet: Optional[str] = None,
        format: str = "json",
        max_sample_rows: int = 1000,
    ) -> Union[str, Dict[str, Any]]:
        """Serializa el esquema para inyección de contexto en prompts de LLMs."""
        return get_ai_grounding_context(
            self.backend,
            table_or_range=table_or_range,
            sheet=sheet,
            format=format,
            max_sample_rows=max_sample_rows,
        )

    def batch(self, operations: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Ejecuta secuencialmente una lista de operaciones sobre la misma instancia del motor
        y realiza un único save() al finalizar todas las operaciones.
        """
        results = []
        for idx, op in enumerate(operations):
            cmd = op.get("command") or op.get("cmd")
            if not cmd:
                raise ValueError(f"Operación en índice {idx} no especifica 'command'.")

            args = op.get("args") if "args" in op else {k: v for k, v in op.items() if k != "command"}
            res = self._execute_batch_op(cmd, args)
            results.append({"command": cmd, "result": res})

        saved_path = self.save()
        return {
            "status": "success",
            "operations_executed": len(results),
            "results": results,
            "saved_to": saved_path,
        }

    def _execute_batch_op(self, cmd: str, args: Dict[str, Any]) -> Any:
        """Despacha una operación individual dentro de un lote sin invocar save()."""
        if cmd == "set-range":
            return self.set_cell_range(
                start_cell=args["start_cell"],
                values=args["values"],
                sheet=args.get("sheet"),
                formulas=args.get("formulas"),
                styles=args.get("styles"),
                autofit=args.get("autofit", False),
            )
        elif cmd == "set-cells":
            return self.set_cells(
                cells_map=args["cells"],
                sheet=args.get("sheet"),
                autofit=args.get("autofit", False),
            )
        elif cmd == "copy-paste":
            return self.copy_paste_range(
                src=args["src"],
                dst=args["dst"],
                src_sheet=args.get("src_sheet"),
                dst_sheet=args.get("dst_sheet"),
                paste_type=args.get("paste_type", args.get("type", "all")),
            )
        elif cmd == "add-column":
            return self.add_calculated_column(
                range_str=args.get("range", ""),
                header=args["header"],
                formula=args["formula"],
                sheet=args.get("sheet"),
                number_format=args.get("number_format"),
                autofit=args.get("autofit", True),
            )
        elif cmd == "create-pivot":
            return self.create_pivot_table(
                source=args.get("source"),
                dest_sheet=args.get("dest_sheet", "Pivot"),
                dest_cell=args.get("dest_cell", "A3"),
                rows=args.get("rows"),
                cols=args.get("cols"),
                value_field=args.get("value_field"),
                value_func=args.get("value_func", "sum"),
                table_name=args.get("table_name"),
                source_sheet=args.get("source_sheet"),
                value_format=args.get("value_format"),
            )
        elif cmd == "create-chart":
            return self.create_chart(
                source=args.get("source"),
                chart_type=args.get("chart_type", "column_clustered"),
                dest_sheet=args.get("dest_sheet"),
                dest_cell=args.get("anchor", args.get("dest_cell")),
                title=args.get("title"),
                sheet=args.get("source_sheet", args.get("sheet")),
                width=args.get("width"),
                height=args.get("height"),
                placement=args.get("placement", "auto"),
                anchor=args.get("anchor"),
            )
        elif cmd == "extend-chart-series":
            return self.extend_chart_series(
                to_col=args.get("to_col"),
                to_row=args.get("to_row"),
                sheet=args.get("sheet"),
                charts=args.get("charts"),
            )
        elif cmd == "update-chart-series":
            return self.update_chart_series(
                chart_name=args["chart_name"],
                series_index=int(args["series_index"]),
                x_range=args.get("x_range"),
                y_range=args.get("y_range"),
                name_ref=args.get("name_ref"),
                formula=args.get("formula"),
                sheet=args.get("sheet"),
            )
        elif cmd == "apply-preset":
            return self.apply_style_preset(
                range_str=args["range"],
                preset=args.get("preset", args.get("preset_name", "header")),
                sheet=args.get("sheet"),
            )
        elif cmd == "summary-table":
            return self.create_summary_table(
                source_range=args.get("source", ""),
                dest_sheet=args.get("dest_sheet"),
                dest_cell=args.get("dest_cell", args.get("start_cell")),
                sheet=args.get("sheet"),
                title=args.get("title"),
            )
        elif cmd == "dedup":
            return self.remove_duplicates(
                range_or_table=args["source"],
                key_columns=args.get("key_cols"),
                sheet=args.get("sheet"),
            )
        elif cmd == "flag-outliers":
            return self.detect_and_flag_outliers(
                range_or_table=args.get("source"),
                column=args["col"],
                method=args.get("method", "iqr"),
                sheet=args.get("sheet"),
            )
        elif cmd == "clean":
            return self.clean_table_dataset(
                target_range_or_table=args["source"],
                rules=args.get("rules", {}),
                sheet=args.get("sheet"),
            )
        elif cmd == "curate":
            return self.curate_pipeline(
                target_range_or_table=args["source"],
                pipeline_spec=args.get("spec", {}),
                sheet=args.get("sheet"),
            )
        elif cmd == "get-csv":
            return self.get_range_as_csv(
                range_str=args["range"],
                sheet=args.get("sheet"),
                delimiter=args.get("delimiter", ";"),
            )
        elif cmd == "get-ranges":
            return self.get_cell_ranges(
                ranges=args["ranges"],
                sheet=args.get("sheet"),
                include_formulas=args.get("formulas", False),
            )
        elif cmd == "inspect-workbook":
            return self.inspect_workbook(
                sheet=args.get("sheet"),
                include_charts=args.get("include_charts", True),
            )
        elif cmd == "inspect-charts":
            return self.inspect_charts(
                sheet=args.get("sheet"),
            )
        elif cmd in ("create-table", "create_table"):
            return self.create_table(
                range_str=args.get("range_str") or args.get("range", ""),
                table_name=args.get("table_name") or args.get("name"),
                style=args.get("style", "TableStyleMedium9"),
                has_headers=args.get("has_headers", True),
                sheet=args.get("sheet"),
            )
        elif cmd in ("list-tables", "list_tables"):
            return self.list_tables(sheet=args.get("sheet"))
        elif cmd in ("inspect-table", "inspect_table"):
            return self.inspect_table_schema(
                table_or_range=args.get("table_or_range") or args.get("range") or args.get("table", ""),
                sheet=args.get("sheet"),
            )
        elif cmd == "status":
            return self.status()
        else:
            raise ValueError(f"Comando batch no reconocido: '{cmd}'")

    # ─────────────────────────── Compatibilidad ─────────────────────────────

    def __getattr__(self, name: str) -> Any:
        """Fallback al backend para métodos avanzados no expuestos directamente."""
        return getattr(self.backend, name)

    def __repr__(self) -> str:
        return (
            f"AntigravityExcelEngine("
            f"mode={self.mode!r}, "
            f"backend={type(self.backend).__name__}, "
            f"file={self.file_path!r})"
        )


# Alias de compatibilidad
ExcelEngine = AntigravityExcelEngine

__all__ = ["AntigravityExcelEngine", "ExcelEngine"]
