# Creado por Francisco Grandón Vergara
"""
Daemon IPC residente para Antigravity Excel Engine.

Mantiene la sesión de Excel en memoria (warm session) y expone
un canal de comunicación Named Pipe (Windows) de baja latencia.

Responsabilidades:
  - Inicializar y conservar la instancia AntigravityExcelEngine en RAM.
  - Procesar solicitudes JSON entrantes con respuestas estructuradas.
  - Registrar todos los errores con logger.exception (nunca silenciarlos).
  - Devolver siempre { status, result | error, [warnings], [latency_ms] }.
"""

import json
import threading
import time
from typing import Any, Dict, Optional

from antigravity_excel_engine.ipc.protocol import (
    ACTION_PING,
    ACTION_SHUTDOWN,
    BUF_SIZE,
    PIPE_NAME,
    build_error_response,
    build_success_response,
    validate_request,
)
from antigravity_excel_engine.models.errors import (
    ExcelEngineError,
    ExcelLockError,
    FormulaValidationError,
    PipeConnectionError,
    UnsupportedOperationError,
    WorkbookNotFoundError,
)
from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("ipc.daemon")

try:
    import pythoncom  # type: ignore
    import pywintypes  # type: ignore
    import win32file  # type: ignore
    import win32pipe  # type: ignore
except ImportError:
    win32pipe = None
    win32file = None
    pywintypes = None
    pythoncom = None

_DEFAULT_CURRENCY_FORMAT = '$ #,##0.00;($ #,##0.00);"-"'


class AntigravityExcelDaemon:
    """
    Servidor Daemon IPC multihilo para Antigravity Excel Engine.

    Mantiene una sesión Excel caliente en memoria y atiende comandos
    vía Named Pipe de Windows con latencia < 5ms en operaciones de lectura.
    """

    def __init__(
        self,
        mode: str = "auto",
        file_path: Optional[str] = None,
        pipe_name: str = PIPE_NAME,
    ) -> None:
        self.pipe_name = pipe_name
        self.mode = mode
        self.file_path = file_path
        self._shutdown_event = threading.Event()
        self.engine = None
        self._lock = threading.Lock()
        self._init_error: Optional[str] = None

    # ──────────────────────────── Engine Init ────────────────────────────────

    def _init_engine(self) -> None:
        """Inicializa la sesión de AntigravityExcelEngine en memoria."""
        try:
            from antigravity_excel_engine.engine import AntigravityExcelEngine
            self.engine = AntigravityExcelEngine(
                mode=self.mode, file_path=self.file_path
            )
            self._init_error = None
            logger.info(
                "Engine iniciado | mode=%s file=%s backend=%s",
                self.mode, self.file_path, type(self.engine.backend).__name__
            )
        except Exception as exc:
            logger.exception("Error al inicializar el engine en el daemon")
            self.engine = None
            self._init_error = str(exc)

    def _reload_engine_if_needed(self, req_file: Optional[str], req_mode: Optional[str]) -> None:
        """Recarga el engine si se solicita un archivo o modo diferente al actual."""
        if not req_file:
            return
        current_file = getattr(self.engine, "file_path", None) if self.engine else None
        if current_file == req_file:
            return
        if self.engine:
            try:
                self.engine.close()
            except Exception:
                pass
        try:
            from antigravity_excel_engine.engine import AntigravityExcelEngine
            self.engine = AntigravityExcelEngine(
                mode=req_mode or "auto", file_path=req_file
            )
            self._init_error = None
            logger.info("Engine recargado | file=%s mode=%s", req_file, req_mode or "auto")
        except Exception as exc:
            logger.exception("Error al recargar el engine")
            self.engine = None
            self._init_error = str(exc)

    # ──────────────────────────── Request Handler ─────────────────────────────

    def handle_request(self, req: Dict[str, Any]) -> Dict[str, Any]:
        """
        Procesa una petición JSON entrante de forma atómica y segura.

        Siempre devuelve un diccionario con al menos { status, result | error }.
        Nunca lanza excepciones al llamante.
        """
        t0 = time.perf_counter()

        # Validar estructura básica
        validation_error = validate_request(req)
        if validation_error:
            return build_error_response(error=validation_error, error_type="InvalidRequest")

        with self._lock:
            try:
                return self._dispatch(req)
            except ExcelLockError as exc:
                logger.warning("ExcelLockError en handle_request: %s", exc)
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type="ExcelLockError",
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )
            except WorkbookNotFoundError as exc:
                logger.warning("WorkbookNotFoundError: %s", exc)
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type="WorkbookNotFoundError",
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )
            except FormulaValidationError as exc:
                logger.warning("FormulaValidationError: %s", exc)
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type="FormulaValidationError",
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )
            except UnsupportedOperationError as exc:
                logger.warning("UnsupportedOperationError: %s", exc)
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type="UnsupportedOperationError",
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )
            except ExcelEngineError as exc:
                logger.exception("ExcelEngineError en handle_request")
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type=type(exc).__name__,
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )
            except Exception as exc:
                logger.exception("Error inesperado en handle_request | cmd=%s", req.get("command"))
                return build_error_response(
                    error=str(exc),
                    command=req.get("command"),
                    error_type="UnexpectedError",
                    latency_ms=(time.perf_counter() - t0) * 1000,
                )

    def _dispatch(self, req: Dict[str, Any]) -> Dict[str, Any]:
        """Despacha la solicitud validada al handler correspondiente."""
        t0 = time.perf_counter()
        action = req.get("action")
        command = req.get("command")
        args = req.get("args", {})

        # ── Acciones de control ────────────────────────────────────────────
        if action == ACTION_PING:
            eng_status = None
            if self.engine is not None:
                try:
                    eng_status = self.engine.status()
                except Exception as exc:
                    eng_status = {"error": str(exc)}
            return build_success_response(
                result={
                    "action": "pong",
                    "time": time.time(),
                    "engine_status": eng_status,
                    "init_error": self._init_error,
                },
                latency_ms=(time.perf_counter() - t0) * 1000,
            )

        if action == ACTION_SHUTDOWN:
            self._shutdown_event.set()
            return build_success_response(
                result={"message": "Daemon deteniéndose correctamente"},
                latency_ms=(time.perf_counter() - t0) * 1000,
            )

        # ── Recarga de libro/modo bajo demanda ────────────────────────────
        self._reload_engine_if_needed(args.get("file"), args.get("mode"))

        # ── Inicialización diferida ────────────────────────────────────────
        if self.engine is None:
            self._init_engine()
            if self.engine is None:
                return build_error_response(
                    error=f"Motor no disponible: {self._init_error}",
                    error_type="EngineInitError",
                )

        # ── Despacho de comandos del motor ─────────────────────────────────
        result = self._execute_command(command, args)
        return build_success_response(
            result=result,
            latency_ms=(time.perf_counter() - t0) * 1000,
        )

    def _execute_command(self, command: Optional[str], args: Dict[str, Any]) -> Any:
        """Ejecuta el comando dado contra el engine activo."""
        engine = self.engine

        if command == "status":
            return engine.status()

        elif command == "get-csv":
            return {
                "csv": engine.get_range_as_csv(
                    args["range"],
                    sheet=args.get("sheet"),
                    delimiter=args.get("delimiter", ";"),
                )
            }

        elif command == "get-ranges":
            return {
                "ranges": engine.get_cell_ranges(
                    args["ranges"],
                    sheet=args.get("sheet"),
                    include_formulas=args.get("formulas", False),
                )
            }

        elif command == "set-range":
            dry_run = bool(args.get("dry_run", False))
            res = engine.set_cell_range(
                args["start_cell"],
                args["values"],
                sheet=args.get("sheet"),
                autofit=args.get("autofit", False),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            return res

        elif command == "set-cells":
            dry_run = bool(args.get("dry_run", False))
            res = engine.set_cells(
                args["cells"],
                sheet=args.get("sheet"),
                autofit=args.get("autofit", False),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            return res

        elif command == "copy-paste":
            res = engine.copy_paste_range(
                args["src"],
                args["dst"],
                source_sheet=args.get("src_sheet"),
                dest_sheet=args.get("dst_sheet"),
                paste_values_only=args.get("values_only", False),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "check-errors":
            errs = engine.check_formula_errors(
                range_str=args.get("range"), sheet=args.get("sheet")
            )
            return {"errors_count": len(errs) if isinstance(errs, list) else 0, "errors": errs}

        elif command == "audit-quality":
            return {
                "audit_report": engine.audit_data_quality(
                    range_str=args.get("range"), sheet=args.get("sheet")
                )
            }

        elif command == "aggregate":
            return {
                "aggregation": engine.aggregate(
                    range_str=args.get("range", ""),
                    group_by_col=args["group_by"],
                    metric_col=args["metric"],
                    agg_func=args.get("func", "sum"),
                    sheet=args.get("sheet"),
                    top_n=args.get("top", 10),
                )
            }

        elif command == "add-column":
            dry_run = bool(args.get("dry_run", False))
            res = engine.add_calculated_column(
                range_str=args.get("range", ""),
                new_col_name=args["header"],
                formula_template=args["formula"],
                sheet=args.get("sheet"),
                col_format=args.get("number_format"),
                dry_run=dry_run,
            )
            if not dry_run:
                res["saved_to"] = engine.save()
            return res

        elif command == "create-pivot":
            res = engine.create_pivot_table(
                source_range=args["source"],
                dest_sheet=args.get("dest_sheet", "Pivot"),
                dest_cell=args.get("dest_cell", "A3"),
                rows=args.get("rows"),
                cols=args.get("cols"),
                value_field=args.get("value_field"),
                value_func=args.get("value_func", "sum"),
                table_name=args.get("table_name"),
                source_sheet=args.get("source_sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "create-chart":
            res = engine.create_chart(
                source_range=args["source"],
                chart_type=args.get("chart_type", "bar"),
                dest_sheet=args.get("dest_sheet"),
                dest_cell=args.get("anchor"),
                title=args.get("title"),
                sheet=args.get("source_sheet"),
                width=args.get("width"),
                height=args.get("height"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "apply-preset":
            res = engine.apply_style_preset(
                range_str=args["range"],
                preset_name=args["preset"],
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "summary-table":
            res = engine.create_summary_table(
                source_range=args.get("source", ""),
                dest_sheet=args.get("dest_sheet"),
                dest_cell=args.get("start_cell"),
                sheet=args.get("sheet"),
                title=args.get("title"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "dedup":
            res = engine.remove_duplicates(
                range_str=args["source"],
                key_columns=args.get("key_cols"),
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "flag-outliers":
            res = engine.detect_and_flag_outliers(
                range_str=args.get("source", ""),
                column=args.get("col"),
                method=args.get("method", "iqr"),
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "clean":
            res = engine.clean_table_dataset(
                range_str=args["source"],
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "curate":
            res = engine.curate_pipeline(
                range_str=args["source"],
                sheet=args.get("sheet"),
                steps=args.get("steps"),
                key_columns=args.get("key_cols"),
                date_columns=args.get("date_cols"),
                numeric_columns=args.get("numeric_cols"),
                outlier_column=args.get("outlier_col"),
                outlier_method=args.get("outlier_method", "iqr"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command in ("wait-until-ready", "wait_until_ready"):
            return engine.wait_until_ready(
                timeout_sec=float(args.get("timeout", args.get("timeout_sec", 10.0))),
                retry_interval=float(args.get("interval", args.get("retry_interval", 0.25))),
            )

        elif command in ("dispatch-com", "dispatch_com"):
            return engine.dispatch_com(
                target=args.get("target", "app"),
                member=args.get("member", ""),
                args=args.get("args"),
                kwargs=args.get("kwargs"),
            )

        elif command == "inspect-workbook":
            return engine.inspect_workbook(
                sheet=args.get("sheet"),
                include_charts=args.get("include_charts", True),
            )

        elif command == "inspect-charts":
            return engine.inspect_charts(
                sheet=args.get("sheet"),
            )

        elif command == "extend-chart-series":
            res = engine.extend_chart_series(
                to_col=args.get("to_col"),
                to_row=args.get("to_row"),
                sheet=args.get("sheet"),
                charts=args.get("charts"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "update-chart-series":
            res = engine.update_chart_series(
                chart_name=args["chart_name"],
                series_index=int(args["series_index"]),
                x_range=args.get("x_range"),
                y_range=args.get("y_range"),
                name_ref=args.get("name_ref"),
                formula=args.get("formula"),
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command == "batch":
            ops = args.get("operations", [])
            return engine.batch(ops)

        elif command in ("create-table", "create_table"):
            res = engine.create_table(
                range_str=args["range"],
                table_name=args.get("table_name") or args.get("name"),
                style=args.get("style", "TableStyleMedium9"),
                has_headers=args.get("has_headers", True),
                sheet=args.get("sheet"),
            )
            res["saved_to"] = engine.save()
            return res

        elif command in ("inspect-table", "inspect_table"):
            fmt = args.get("format", "json")
            if fmt.lower() == "markdown":
                return {
                    "markdown": engine.get_ai_grounding_context(
                        table_or_range=args.get("table_or_range") or args.get("range", ""),
                        sheet=args.get("sheet"),
                        format="markdown",
                    )
                }
            return engine.inspect_table_schema(
                table_or_range=args.get("table_or_range") or args.get("range", ""),
                sheet=args.get("sheet"),
            )

        else:
            raise UnsupportedOperationError(
                f"Comando no reconocido: '{command}'. "
                f"Comandos disponibles: status, get-csv, get-ranges, set-range, set-cells, "
                f"copy-paste, check-errors, audit-quality, aggregate, add-column, "
                f"create-pivot, create-chart, apply-preset, summary-table, dedup, "
                f"flag-outliers, clean, curate, wait-until-ready, dispatch-com, "
                f"inspect-workbook, inspect-charts, extend-chart-series, update-chart-series, batch."
            )

    # ──────────────────────────── Run Loop ───────────────────────────────────

    def run(self) -> None:
        """Bucle principal de escucha del servidor Named Pipe persistente."""
        if win32pipe is None or win32file is None:
            raise RuntimeError(
                "win32pipe y win32file no están disponibles. "
                "Instale pywin32 para usar el daemon IPC."
            )

        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        self._init_engine()
        logger.info("Daemon iniciado | pipe=%s", self.pipe_name)

        h_pipe = None
        try:
            while not self._shutdown_event.is_set():
                # Crear el pipe si no existe
                if h_pipe is None:
                    try:
                        h_pipe = win32pipe.CreateNamedPipe(
                            self.pipe_name,
                            win32pipe.PIPE_ACCESS_DUPLEX,
                            (
                                win32pipe.PIPE_TYPE_MESSAGE
                                | win32pipe.PIPE_READMODE_MESSAGE
                                | win32pipe.PIPE_WAIT
                            ),
                            win32pipe.PIPE_UNLIMITED_INSTANCES,
                            BUF_SIZE,
                            BUF_SIZE,
                            1000,
                            None,
                        )
                    except Exception as exc:
                        logger.error("Error al crear Named Pipe: %s", exc)
                        time.sleep(0.05)
                        continue

                # Esperar conexión de cliente
                try:
                    win32pipe.ConnectNamedPipe(h_pipe, None)
                except pywintypes.error as e:
                    if e.winerror == 535:  # ERROR_PIPE_CONNECTED — ya conectado
                        pass
                    else:
                        logger.warning("ConnectNamedPipe error %d: %s", e.winerror, e)
                        try:
                            win32file.CloseHandle(h_pipe)
                        except Exception:
                            pass
                        h_pipe = None
                        continue

                # Leer mensaje del cliente
                try:
                    _, data = win32file.ReadFile(h_pipe, BUF_SIZE)
                    req = json.loads(data.decode("utf-8"))
                except Exception as exc:
                    logger.warning("Error leyendo mensaje del pipe: %s", exc)
                    try:
                        win32pipe.DisconnectNamedPipe(h_pipe)
                    except Exception:
                        pass
                    continue

                # Procesar y responder
                response = self.handle_request(req)
                resp_bytes = json.dumps(response, ensure_ascii=False, default=str).encode("utf-8")
                try:
                    win32file.WriteFile(h_pipe, resp_bytes)
                    win32file.FlushFileBuffers(h_pipe)
                except Exception as exc:
                    logger.warning("Error escribiendo respuesta al pipe: %s", exc)

                try:
                    win32pipe.DisconnectNamedPipe(h_pipe)
                except Exception:
                    try:
                        win32file.CloseHandle(h_pipe)
                    except Exception:
                        pass
                    h_pipe = None

                if req.get("action") == ACTION_SHUTDOWN:
                    logger.info("Daemon recibió shutdown. Saliendo del bucle.")
                    break

        finally:
            if h_pipe is not None:
                try:
                    win32file.CloseHandle(h_pipe)
                except Exception:
                    pass
            if self.engine:
                try:
                    self.engine.close()
                except Exception:
                    pass
            if pythoncom is not None:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            logger.info("Daemon detenido | pipe=%s", self.pipe_name)
