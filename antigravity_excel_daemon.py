# Creado por Francisco Grandón Vergara
"""
Servidor Daemon Residente de Antigravity Excel Engine.
Mantiene la sesión de Excel en memoria (RAM) y provee un canal de comunicación
ultra-rápido IPC mediante Named Pipe de Windows (\\\\.\\pipe\\antigravity_excel).
"""

import argparse
import json
import os
import subprocess
import sys
import threading
import time
from typing import Any, Dict, Optional, Tuple

try:
    import win32pipe
    import win32file
    import pywintypes
    import pythoncom
except ImportError:
    win32pipe = None
    win32file = None
    pywintypes = None
    pythoncom = None

PIPE_NAME = r"\\.\pipe\antigravity_excel"
BUF_SIZE = 16 * 1024 * 1024  # 16 MB para transferencias masivas de CSV/datos
_DEFAULT_CURRENCY_FORMAT = "$ #,##0.00;($ #,##0.00);\"-\""


class AntigravityExcelDaemon:
    """Servidor Daemon IPC multihilo para Antigravity Excel Engine."""

    def __init__(self, mode: str = "auto", file_path: Optional[str] = None, pipe_name: str = PIPE_NAME):
        self.pipe_name = pipe_name
        self.mode = mode
        self.file_path = file_path
        self._shutdown_event = threading.Event()
        self.engine = None
        self._lock = threading.Lock()
        self._init_error = None

    def _init_engine(self):
        """Inicializa la sesión de AntigravityExcelEngine en memoria."""
        try:
            from antigravity_excel_core import AntigravityExcelEngine
            self.engine = AntigravityExcelEngine(mode=self.mode, file_path=self.file_path)
            self._init_error = None
        except Exception as e:
            self.engine = None
            self._init_error = str(e)

    def handle_request(self, req: Dict[str, Any]) -> Dict[str, Any]:
        """Procesa una petición JSON entrante de forma atómica y segura."""
        with self._lock:
            try:
                # Comandos de control administrativo
                action = req.get("action")
                if action == "ping":
                    eng_status = None
                    if self.engine is not None:
                        try:
                            eng_status = self.engine.status()
                        except Exception as e:
                            eng_status = {"error": str(e)}
                    return {
                        "status": "success",
                        "result": {
                            "action": "pong",
                            "time": time.time(),
                            "engine_status": eng_status,
                            "init_error": self._init_error
                        }
                    }

                if action == "shutdown":
                    self._shutdown_event.set()
                    return {
                        "status": "success",
                        "result": {
                            "message": "Daemon deteniéndose correctamente"
                        }
                    }

                command = req.get("command")
                args = req.get("args", {})

                # Cambio o recarga bajo demanda de libro/modo si se especifica explícitamente
                req_file = args.get("file")
                req_mode = args.get("mode")
                if req_file and (not self.engine or getattr(self.engine, "file_path", None) != req_file):
                    if self.engine:
                        try:
                            self.engine.close()
                        except Exception:
                            pass
                    from antigravity_excel_core import AntigravityExcelEngine
                    self.engine = AntigravityExcelEngine(mode=req_mode or "auto", file_path=req_file)

                if self.engine is None:
                    self._init_engine()
                    if self.engine is None:
                        return {"status": "error", "error": f"Motor no disponible: {self._init_error}"}

                # Despacho de operaciones del motor
                if command == "status":
                    return {"status": "success", "result": self.engine.status()}

                elif command == "get-csv":
                    csv_data = self.engine.get_range_as_csv(
                        args["range"],
                        sheet=args.get("sheet"),
                        delimiter=args.get("delimiter", ";")
                    )
                    return {"status": "success", "result": {"status": "success", "csv": csv_data}}

                elif command == "get-ranges":
                    res = self.engine.get_cell_ranges(
                        args["ranges"],
                        sheet=args.get("sheet"),
                        include_formulas=args.get("formulas", False)
                    )
                    return {"status": "success", "result": {"status": "success", "ranges": res}}

                elif command == "set-range":
                    res = self.engine.set_cell_range(
                        args["start_cell"],
                        args["values"],
                        sheet=args.get("sheet"),
                        autofit=args.get("autofit", False)
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "set-cells":
                    res = self.engine.set_cells(
                        args["cells"],
                        sheet=args.get("sheet"),
                        autofit=args.get("autofit", False)
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "copy-paste":
                    res = self.engine.copy_paste_range(
                        args["src"],
                        args["dst"],
                        args.get("src_sheet"),
                        args.get("dst_sheet"),
                        args.get("type", "all")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "check-errors":
                    errs = self.engine.check_formula_errors(args["range"], sheet=args.get("sheet"))
                    return {
                        "status": "success",
                        "result": {
                            "status": "success",
                            "errors_count": len(errs),
                            "errors": errs
                        }
                    }

                elif command == "audit-quality":
                    report = self.engine.audit_data_quality(range_str=args.get("range"), sheet=args.get("sheet"))
                    return {"status": "success", "result": {"status": "success", "audit_report": report}}

                elif command == "aggregate":
                    res = self.engine.aggregate(
                        group_by_col=args["group_by"],
                        metric_col=args["metric"],
                        agg_func=args.get("func", "sum"),
                        sheet=args.get("sheet"),
                        top_n=args.get("top", 10),
                        ascending=args.get("ascending", False)
                    )
                    return {"status": "success", "result": {"status": "success", "aggregation": res}}

                elif command == "add-column":
                    res = self.engine.add_calculated_column(
                        header=args["header"],
                        formula_template=args["formula"],
                        number_format=args.get("number_format"),
                        sheet=args.get("sheet"),
                        autofit=args.get("autofit", True)
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "create-pivot":
                    res = self.engine.create_pivot_table(
                        source=args["source"],
                        rows=args["rows"],
                        cols=args.get("cols"),
                        value_field=args["value_field"],
                        value_func=args.get("value_func", "sum"),
                        value_format=args.get("value_format", _DEFAULT_CURRENCY_FORMAT),
                        dest_sheet=args.get("dest_sheet"),
                        dest_cell=args.get("dest_cell", "A3"),
                        table_name=args.get("table_name"),
                        source_sheet=args.get("source_sheet")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "create-chart":
                    res = self.engine.create_chart(
                        source=args["source"],
                        chart_type=args.get("chart_type", "column_clustered"),
                        title=args.get("title"),
                        dest_sheet=args.get("dest_sheet"),
                        placement=args.get("placement", "auto"),
                        anchor=args.get("anchor"),
                        width=args.get("width", 480),
                        height=args.get("height", 300),
                        source_sheet=args.get("source_sheet")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "apply-preset":
                    res = self.engine.apply_style_preset(args["range"], preset=args["preset"], sheet=args.get("sheet"))
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "summary-table":
                    res = self.engine.create_summary_table(
                        title=args.get("title", "Resumen"),
                        headers=args.get("headers", []),
                        rows_spec=args.get("rows_spec", []),
                        dest_sheet=args.get("dest_sheet", "Resumen"),
                        start_cell=args.get("start_cell", "B2"),
                        parameters=args.get("parameters")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "dedup":
                    res = self.engine.remove_duplicates(
                        range_or_table=args["source"],
                        key_columns=args.get("key_cols"),
                        sheet=args.get("sheet")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "flag-outliers":
                    res = self.engine.detect_and_flag_outliers(
                        column=args["col"],
                        method=args.get("method", "iqr"),
                        range_or_table=args.get("source"),
                        sheet=args.get("sheet"),
                        add_comment=args.get("add_comment", True)
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "clean":
                    res = self.engine.clean_table_dataset(
                        target_range_or_table=args["source"],
                        rules=args["rules"],
                        sheet=args.get("sheet")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                elif command == "curate":
                    res = self.engine.curate_pipeline(
                        target_range_or_table=args["source"],
                        pipeline_spec=args["spec"],
                        sheet=args.get("sheet")
                    )
                    saved_to = self.engine.save()
                    res["saved_to"] = saved_to
                    return {"status": "success", "result": res}

                else:
                    return {"status": "error", "error": f"Comando no reconocido: {command}"}

            except Exception as exc:
                return {"status": "error", "error": str(exc)}

    def run(self):
        """Bucle principal de escucha del servidor Named Pipe persistente."""
        if win32pipe is None or win32file is None:
            raise RuntimeError("win32pipe y win32file no están disponibles en este entorno.")

        if pythoncom is not None:
            try:
                pythoncom.CoInitialize()
            except Exception:
                pass

        self._init_engine()

        h_pipe = None
        try:
            while not self._shutdown_event.is_set():
                if h_pipe is None:
                    try:
                        h_pipe = win32pipe.CreateNamedPipe(
                            self.pipe_name,
                            win32pipe.PIPE_ACCESS_DUPLEX,
                            win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                            win32pipe.PIPE_UNLIMITED_INSTANCES,
                            BUF_SIZE,
                            BUF_SIZE,
                            1000,
                            None
                        )
                    except Exception:
                        time.sleep(0.05)
                        continue

                try:
                    win32pipe.ConnectNamedPipe(h_pipe, None)
                except pywintypes.error as e:
                    if e.winerror == 535:  # ERROR_PIPE_CONNECTED
                        pass
                    else:
                        try:
                            win32file.CloseHandle(h_pipe)
                        except Exception:
                            pass
                        h_pipe = None
                        continue

                # Lectura del mensaje del cliente
                try:
                    hr, data = win32file.ReadFile(h_pipe, BUF_SIZE)
                    req = json.loads(data.decode("utf-8"))
                except Exception:
                    try:
                        win32pipe.DisconnectNamedPipe(h_pipe)
                    except Exception:
                        pass
                    continue

                # Procesar petición
                response = self.handle_request(req)
                resp_bytes = json.dumps(response, ensure_ascii=False).encode("utf-8")

                # Escribir respuesta de retorno y desconectar cliente
                try:
                    win32file.WriteFile(h_pipe, resp_bytes)
                    win32file.FlushFileBuffers(h_pipe)
                except Exception:
                    pass

                try:
                    win32pipe.DisconnectNamedPipe(h_pipe)
                except Exception:
                    try:
                        win32file.CloseHandle(h_pipe)
                    except Exception:
                        pass
                    h_pipe = None

                if req.get("action") == "shutdown":
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


# ===========================================================================
# Funciones Cliente de Named Pipe (Ultrarrápidas, < 30ms)
# ===========================================================================

def send_pipe_message(payload: Dict[str, Any], timeout_ms: int = 1000, pipe_name: str = PIPE_NAME) -> Optional[Dict[str, Any]]:
    """
    Envía un mensaje JSON al Named Pipe del daemon y devuelve la respuesta parseada.
    Retorna None si el pipe no existe o la conexión falla.
    """
    if win32pipe is None:
        return None
    try:
        req_bytes = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        resp_bytes = win32pipe.CallNamedPipe(pipe_name, req_bytes, BUF_SIZE, timeout_ms)
        return json.loads(resp_bytes.decode("utf-8"))
    except Exception:
        return None


def ping_daemon(timeout_ms: int = 1000, pipe_name: str = PIPE_NAME) -> Tuple[bool, float, Optional[Dict[str, Any]]]:
    """
    Realiza un ping al daemon midiendo la latencia exacta en milisegundos.
    Retorna (is_alive, latency_ms, response_data).
    """
    t0 = time.perf_counter()
    resp = send_pipe_message({"action": "ping"}, timeout_ms=timeout_ms, pipe_name=pipe_name)
    latency_ms = (time.perf_counter() - t0) * 1000.0
    if resp and resp.get("status") == "success":
        return True, latency_ms, resp.get("result", {})
    return False, latency_ms, None


def start_daemon_process(mode: str = "auto", file_path: Optional[str] = None, pipe_name: str = PIPE_NAME) -> Dict[str, Any]:
    """Inicia el servidor daemon en segundo plano (background) si no está activo."""
    alive, lat, details = ping_daemon(timeout_ms=500, pipe_name=pipe_name)
    if alive:
        return {
            "status": "info",
            "message": "El daemon ya está en ejecución",
            "pipe": pipe_name,
            "latency_ms": round(lat, 2),
            "details": details
        }

    daemon_script = os.path.abspath(__file__)
    cmd = [sys.executable, daemon_script, "--run", "--mode", mode]
    if file_path:
        cmd.extend(["--file", file_path])

    creationflags = 0
    if sys.platform == "win32":
        DETACHED_PROCESS = 0x00000008
        CREATE_NO_WINDOW = 0x08000000
        creationflags = DETACHED_PROCESS | CREATE_NO_WINDOW

    proc = subprocess.Popen(
        cmd,
        creationflags=creationflags,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True
    )

    # Esperar hasta 5 segundos a que el pipe esté operativo
    t_end = time.perf_counter() + 5.0
    while time.perf_counter() < t_end:
        alive, lat, details = ping_daemon(timeout_ms=100, pipe_name=pipe_name)
        if alive:
            return {
                "status": "success",
                "message": "Daemon iniciado correctamente",
                "pid": proc.pid,
                "pipe": pipe_name,
                "latency_ms": round(lat, 2),
                "details": details
            }
        time.sleep(0.05)

    return {
        "status": "error",
        "error": "Timeout esperando que el daemon inicialice el Named Pipe",
        "pid": proc.pid
    }


def stop_daemon_process(pipe_name: str = PIPE_NAME) -> Dict[str, Any]:
    """Envía la señal de apagado al daemon vía Named Pipe."""
    alive, _, _ = ping_daemon(timeout_ms=500, pipe_name=pipe_name)
    if not alive:
        return {
            "status": "info",
            "message": "El daemon no estaba en ejecución",
            "pipe": pipe_name
        }

    resp = send_pipe_message({"action": "shutdown"}, timeout_ms=2000, pipe_name=pipe_name)
    return {
        "status": "success",
        "message": "Daemon detenido exitosamente",
        "pipe": pipe_name,
        "response": resp
    }


def status_daemon(pipe_name: str = PIPE_NAME) -> Dict[str, Any]:
    """Verifica si el daemon está vivo y reporta latencia y estado del libro."""
    alive, lat, details = ping_daemon(timeout_ms=1000, pipe_name=pipe_name)
    if alive:
        return {
            "status": "success",
            "daemon_running": True,
            "latency_ms": round(lat, 2),
            "pipe": pipe_name,
            "engine_status": details.get("engine_status") if details else None
        }
    return {
        "status": "info",
        "daemon_running": False,
        "message": "Daemon no está en ejecución",
        "pipe": pipe_name
    }


def main():
    """Punto de entrada CLI para administración del daemon."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Antigravity Excel Engine Daemon")
    parser.add_argument("--start", action="store_true", help="Iniciar daemon en segundo plano")
    parser.add_argument("--stop", action="store_true", help="Detener daemon vía señal Named Pipe")
    parser.add_argument("--status", action="store_true", help="Consultar estado y latencia del daemon")
    parser.add_argument("--run", action="store_true", help="Ejecutar bucle del servidor en primer plano")
    parser.add_argument("--mode", default="auto", choices=["auto", "live", "file"], help="Modo de ejecución del motor")
    parser.add_argument("--file", help="Ruta al archivo .xlsx inicial")
    parser.add_argument("--json", action="store_true", help="Salida en formato JSON")

    args = parser.parse_args()

    def out(data):
        print(json.dumps(data, indent=2, ensure_ascii=False))

    if args.start:
        res = start_daemon_process(mode=args.mode, file_path=args.file)
        out(res)
    elif args.stop:
        res = stop_daemon_process()
        out(res)
    elif args.status:
        res = status_daemon()
        out(res)
    elif args.run:
        server = AntigravityExcelDaemon(mode=args.mode, file_path=args.file)
        server.run()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
