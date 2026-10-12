# Creado por Francisco Grandón Vergara
"""
Shim de compatibilidad retroactiva para antigravity_excel_daemon.
Re-exporta todas las clases y funciones del nuevo subpaquete ipc/.

Los scripts que importen directamente desde este archivo seguirán funcionando.
"""

import argparse
import json
import sys
from typing import Any, Dict, Optional

# Re-exports del subpaquete ipc/
from antigravity_excel_engine.ipc.protocol import (
    PIPE_NAME,
    BUF_SIZE,
)
from antigravity_excel_engine.ipc.client import (
    send_pipe_message,
    ping_daemon,
    call_command,
    start_daemon_process,
    stop_daemon_process,
    status_daemon,
)
from antigravity_excel_engine.ipc.daemon import AntigravityExcelDaemon


def main() -> None:
    """Punto de entrada CLI para administración del daemon."""
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")

    parser = argparse.ArgumentParser(description="Antigravity Excel Engine Daemon")
    parser.add_argument("--start", action="store_true", help="Iniciar daemon en segundo plano")
    parser.add_argument("--stop", action="store_true", help="Detener daemon vía señal Named Pipe")
    parser.add_argument("--status", action="store_true", help="Consultar estado y latencia")
    parser.add_argument("--run", action="store_true", help="Ejecutar servidor en primer plano")
    parser.add_argument("--mode", default="auto", choices=["auto", "live", "file"])
    parser.add_argument("--file", help="Ruta al archivo .xlsx inicial")
    parser.add_argument("--pipe", default=PIPE_NAME, help="Nombre del Named Pipe")
    parser.add_argument("--json", action="store_true", help="Salida en formato JSON")

    args = parser.parse_args()

    def out(data: Any) -> None:
        print(json.dumps(data, indent=2, ensure_ascii=False))

    if args.start:
        out(start_daemon_process(mode=args.mode, file_path=args.file, pipe_name=args.pipe))
    elif args.stop:
        out(stop_daemon_process(pipe_name=args.pipe))
    elif args.status:
        out(status_daemon(pipe_name=args.pipe))
    elif args.run:
        server = AntigravityExcelDaemon(
            mode=args.mode, file_path=args.file, pipe_name=args.pipe
        )
        server.run()
    else:
        parser.print_help()


__all__ = [
    "PIPE_NAME",
    "BUF_SIZE",
    "AntigravityExcelDaemon",
    "send_pipe_message",
    "ping_daemon",
    "call_command",
    "start_daemon_process",
    "stop_daemon_process",
    "status_daemon",
]

if __name__ == "__main__":
    main()
