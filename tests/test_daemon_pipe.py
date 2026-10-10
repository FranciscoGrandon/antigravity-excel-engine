# Creado por Francisco Grandón Vergara
"""
Tests unitarios y de integración para el Servidor Daemon Residente
y comunicación ultra-rápida por Named Pipe (\\\\.\\pipe\\antigravity_excel).
"""

import json
import os
import sys
import threading
import time
import pytest

# Agregar ruta padre
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_daemon import (
    AntigravityExcelDaemon,
    PIPE_NAME,
    send_pipe_message,
    ping_daemon,
    status_daemon,
    stop_daemon_process
)
from antigravity_excel_cli import _try_call_daemon


@pytest.fixture(scope="module")
def active_daemon(tmp_path_factory):
    """Inicia un daemon en un hilo en segundo plano con un pipe de prueba."""
    tmp_dir = tmp_path_factory.mktemp("daemon_test")
    wb_path = os.path.join(str(tmp_dir), "daemon_test_wb.xlsx")
    pipe_test = r"\\.\pipe\antigravity_excel_test_suite"

    daemon = AntigravityExcelDaemon(mode="file", file_path=wb_path, pipe_name=pipe_test)
    t = threading.Thread(target=daemon.run, daemon=True)
    t.start()

    # Esperar activamente hasta que el Named Pipe responda al ping
    t_end = time.time() + 5.0
    alive = False
    details = None
    while time.time() < t_end:
        alive, lat, details = ping_daemon(timeout_ms=150, pipe_name=pipe_test)
        if alive:
            break
        time.sleep(0.05)

    assert alive is True, f"El daemon de prueba no pudo inicializarse a tiempo: {daemon._init_error}"

    yield {"daemon": daemon, "thread": t, "pipe": pipe_test, "file_path": wb_path}

    # Apagado ordenado
    send_pipe_message({"action": "shutdown"}, timeout_ms=2000, pipe_name=pipe_test)
    t.join(timeout=3.0)


class TestDaemonNamedPipe:
    """Suite de pruebas para el daemon y Named Pipe IPC."""

    def test_daemon_ping_latency(self, active_daemon):
        """Valida que la latencia de ping a través del Named Pipe sea < 30ms (SLA estricto)."""
        pipe = active_daemon["pipe"]
        latencies = []

        # Ejecutar ráfaga de 10 pings consecutivos
        for _ in range(10):
            alive, lat, details = ping_daemon(timeout_ms=500, pipe_name=pipe)
            assert alive is True
            assert details.get("action") == "pong"
            assert details.get("time") is not None
            latencies.append(lat)

        avg_lat = sum(latencies) / len(latencies)
        min_lat = min(latencies)
        max_lat = max(latencies)

        print(f"\n[TELEMETRIA PIPE] Ping Promedio: {avg_lat:.2f} ms | Min: {min_lat:.2f} ms | Max: {max_lat:.2f} ms")
        assert avg_lat < 30.0, f"Latencia promedio {avg_lat:.2f}ms excede el umbral crítico de 30ms"

    def test_daemon_status_command(self, active_daemon):
        """Valida la ejecución del comando 'status' sobre el Named Pipe."""
        pipe = active_daemon["pipe"]
        resp = send_pipe_message({"command": "status", "args": {}}, timeout_ms=1000, pipe_name=pipe)
        assert resp is not None
        assert resp.get("status") == "success"
        res = resp.get("result")
        assert "backend" in res
        assert "sheets" in res

    def test_daemon_set_range_and_get_csv(self, active_daemon):
        """Valida escritura atómica set-range y lectura get-csv sobre el Named Pipe."""
        pipe = active_daemon["pipe"]

        # 1. Escribir datos
        data = [
            ["ID", "Producto", "Monto"],
            [1, "Servicio A", 1500.50],
            [2, "Servicio B", 2300.00]
        ]
        set_req = {
            "command": "set-range",
            "args": {
                "start_cell": "A1",
                "values": data,
                "sheet": "Sheet",
                "autofit": True
            }
        }
        set_resp = send_pipe_message(set_req, timeout_ms=1000, pipe_name=pipe)
        assert set_resp is not None
        assert set_resp.get("status") == "success"
        assert set_resp["result"]["rows"] == 3
        assert set_resp["result"]["cols"] == 3

        # 2. Leer como CSV
        csv_req = {
            "command": "get-csv",
            "args": {
                "range": "A1:C3",
                "sheet": "Sheet",
                "delimiter": ";"
            }
        }
        csv_resp = send_pipe_message(csv_req, timeout_ms=1000, pipe_name=pipe)
        assert csv_resp is not None
        assert csv_resp.get("status") == "success"
        csv_text = csv_resp["result"]["csv"]
        assert "Servicio A" in csv_text
        assert "1500.5" in csv_text

    def test_daemon_get_ranges(self, active_daemon):
        """Valida lectura de rangos dispersos get-ranges sobre Named Pipe."""
        pipe = active_daemon["pipe"]
        req = {
            "command": "get-ranges",
            "args": {
                "ranges": ["A1:C1", "B2"],
                "sheet": "Sheet"
            }
        }
        resp = send_pipe_message(req, timeout_ms=1000, pipe_name=pipe)
        assert resp is not None
        assert resp.get("status") == "success"
        ranges = resp["result"]["ranges"]
        assert "A1:C1" in ranges
        assert "B2" in ranges

    def test_daemon_audit_quality(self, active_daemon):
        """Valida auditoría de calidad audit-quality sobre Named Pipe."""
        pipe = active_daemon["pipe"]
        req = {
            "command": "audit-quality",
            "args": {
                "range": "A1:C3",
                "sheet": "Sheet"
            }
        }
        resp = send_pipe_message(req, timeout_ms=1000, pipe_name=pipe)
        assert resp is not None
        assert resp.get("status") == "success"
        report = resp["result"]["audit_report"]
        assert "total_cells" in report
        assert report["total_cells"] == 9

    def test_daemon_unknown_command(self, active_daemon):
        """Valida manejo resiliente de errores ante comando desconocido."""
        pipe = active_daemon["pipe"]
        resp = send_pipe_message({"command": "comando_inexistente", "args": {}}, timeout_ms=1000, pipe_name=pipe)
        assert resp is not None
        assert resp.get("status") == "error"
        assert "no reconocido" in resp.get("error", "").lower()

    def test_cli_fallback_when_daemon_offline(self):
        """Valida que el cliente CLI caiga a standalone en < 5ms si el daemon está inactivo."""
        t0 = time.perf_counter()
        res = send_pipe_message({"command": "status", "args": {}}, timeout_ms=50, pipe_name=r"\\.\pipe\pipe_no_existe_xyz")
        dur_ms = (time.perf_counter() - t0) * 1000.0

        assert res is None, "Debe retornar None ante pipe inexistente"
        assert dur_ms < 10.0, f"El chequeo de fallback tardó demasiado: {dur_ms:.2f}ms"

    def test_daemon_status_and_stop_offline(self):
        """Valida que los helpers status_daemon y stop_daemon respondan limpiamente si el daemon no corre."""
        pipe = r"\\.\pipe\pipe_no_existe_offline"
        st = status_daemon(pipe_name=pipe)
        assert st["daemon_running"] is False
        assert st["status"] == "info"

        stop_res = stop_daemon_process(pipe_name=pipe)
        assert stop_res["status"] == "info"
        assert "no estaba en ejecución" in stop_res["message"]
