# Creado por Francisco Grandón Vergara
"""
Cliente IPC de Named Pipe para Antigravity Excel Engine.

Encapsula todas las funciones de comunicación con el daemon residente:
  - send_pipe_message: envío atómico de mensajes JSON
  - ping_daemon: verificación de disponibilidad con latencia
  - call_command: llamada de alto nivel a un comando del motor
  - start_daemon_process / stop_daemon_process / status_daemon: ciclo de vida
"""

import json
import os
import subprocess
import sys
import time
from typing import Any, Dict, Optional, Tuple

from antigravity_excel_engine.ipc.protocol import (
    PIPE_NAME,
    BUF_SIZE,
    DEFAULT_TIMEOUT_MS,
    START_TIMEOUT_SEC,
    ACTION_PING,
    ACTION_SHUTDOWN,
    build_command_request,
    build_action_request,
    build_error_response,
)
from antigravity_excel_engine.utils.logging import get_logger

logger = get_logger("ipc.client")

try:
    import win32pipe  # type: ignore
    import win32file  # type: ignore
    import pywintypes  # type: ignore
except ImportError:
    win32pipe = None
    win32file = None
    pywintypes = None


# ──────────────────────────────────────────────────────────────────────────────
# Primitiva de transporte
# ──────────────────────────────────────────────────────────────────────────────

def send_pipe_message(
    payload: Dict[str, Any],
    timeout_ms: int = DEFAULT_TIMEOUT_MS,
    pipe_name: str = PIPE_NAME,
) -> Optional[Dict[str, Any]]:
    """
    Envía un mensaje JSON al Named Pipe del daemon y devuelve la respuesta.

    Args:
        payload: Diccionario a serializar y enviar.
        timeout_ms: Timeout de espera al Named Pipe en milisegundos.
        pipe_name: Nombre del Named Pipe (por defecto el canónico).

    Returns:
        Diccionario de respuesta parseado, o None si el pipe no está disponible
        o la conexión falla.
    """
    if win32pipe is None:
        logger.debug("win32pipe no disponible; send_pipe_message retorna None")
        return None
    try:
        req_bytes = json.dumps(payload, ensure_ascii=False, default=str).encode("utf-8")
        resp_bytes = win32pipe.CallNamedPipe(pipe_name, req_bytes, BUF_SIZE, timeout_ms)
        return json.loads(resp_bytes.decode("utf-8"))
    except Exception as exc:
        logger.debug("send_pipe_message error: %s", exc)
        return None


# ──────────────────────────────────────────────────────────────────────────────
# Funciones de Alto Nivel
# ──────────────────────────────────────────────────────────────────────────────

def ping_daemon(
    timeout_ms: int = DEFAULT_TIMEOUT_MS,
    pipe_name: str = PIPE_NAME,
) -> Tuple[bool, float, Optional[Dict[str, Any]]]:
    """
    Verifica si el daemon está activo y mide la latencia de respuesta.

    Args:
        timeout_ms: Tiempo máximo de espera en milisegundos.
        pipe_name: Nombre del Named Pipe.

    Returns:
        Tupla (is_alive: bool, latency_ms: float, response_data: dict | None).
    """
    t0 = time.perf_counter()
    resp = send_pipe_message(
        build_action_request(ACTION_PING),
        timeout_ms=timeout_ms,
        pipe_name=pipe_name,
    )
    latency_ms = (time.perf_counter() - t0) * 1000.0
    if resp and resp.get("status") == "success":
        logger.debug("ping OK | latency=%.2fms", latency_ms)
        return True, latency_ms, resp.get("result", {})
    return False, latency_ms, None


def call_command(
    command: str,
    args: Optional[Dict[str, Any]] = None,
    timeout_ms: int = DEFAULT_TIMEOUT_MS,
    pipe_name: str = PIPE_NAME,
) -> Optional[Dict[str, Any]]:
    """
    Envía un comando de motor al daemon y retorna la respuesta.

    Esta es la función principal para interaccionar con el daemon desde el CLI
    o desde scripts externos.

    Args:
        command: Nombre del comando (ej. "get-csv", "aggregate").
        args: Argumentos del comando como diccionario.
        timeout_ms: Timeout en milisegundos.
        pipe_name: Nombre del Named Pipe.

    Returns:
        Diccionario de respuesta, o None si el daemon no está disponible.
    """
    logger.debug("call_command | cmd=%s args=%s", command, args)
    return send_pipe_message(
        build_command_request(command, args),
        timeout_ms=timeout_ms,
        pipe_name=pipe_name,
    )


def start_daemon_process(
    mode: str = "auto",
    file_path: Optional[str] = None,
    pipe_name: str = PIPE_NAME,
) -> Dict[str, Any]:
    """
    Inicia el proceso daemon en segundo plano si no está activo.

    Espera hasta START_TIMEOUT_SEC segundos a que el Named Pipe esté operativo.

    Args:
        mode: Modo del motor ("auto", "live", "file").
        file_path: Ruta opcional al archivo Excel inicial.
        pipe_name: Nombre del Named Pipe.

    Returns:
        Diccionario con estado del inicio ("success", "info" o "error").
    """
    alive, lat, details = ping_daemon(timeout_ms=500, pipe_name=pipe_name)
    if alive:
        return {
            "status": "info",
            "message": "El daemon ya está en ejecución",
            "pipe": pipe_name,
            "latency_ms": round(lat, 2),
            "details": details,
        }

    # Ruta al script del daemon
    daemon_script = os.path.join(
        os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))),
        "antigravity_excel_daemon.py",
    )
    cmd = [sys.executable, daemon_script, "--run", "--mode", mode]
    if file_path:
        cmd.extend(["--file", file_path])
    if pipe_name != PIPE_NAME:
        cmd.extend(["--pipe", pipe_name])

    creationflags = 0
    if sys.platform == "win32":
        DETACHED_PROCESS = 0x00000008
        CREATE_NO_WINDOW = 0x08000000
        creationflags = DETACHED_PROCESS | CREATE_NO_WINDOW

    try:
        proc = subprocess.Popen(
            cmd,
            creationflags=creationflags,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
    except Exception as exc:
        logger.exception("Error al lanzar el proceso daemon")
        return build_error_response(
            error=f"No se pudo iniciar el proceso daemon: {exc}",
            error_type="ProcessLaunchError",
        )

    # Esperar que el pipe esté operativo
    t_end = time.perf_counter() + START_TIMEOUT_SEC
    while time.perf_counter() < t_end:
        alive, lat, details = ping_daemon(timeout_ms=100, pipe_name=pipe_name)
        if alive:
            logger.info("Daemon iniciado | pid=%d latency=%.2fms", proc.pid, lat)
            return {
                "status": "success",
                "message": "Daemon iniciado correctamente",
                "pid": proc.pid,
                "pipe": pipe_name,
                "latency_ms": round(lat, 2),
                "details": details,
            }
        time.sleep(0.05)

    return build_error_response(
        error="Timeout esperando que el daemon inicialice el Named Pipe",
        error_type="TimeoutError",
    )


def stop_daemon_process(pipe_name: str = PIPE_NAME) -> Dict[str, Any]:
    """
    Envía la señal de apagado ordenado al daemon vía Named Pipe.

    Args:
        pipe_name: Nombre del Named Pipe.

    Returns:
        Diccionario con estado de la operación.
    """
    alive, _, _ = ping_daemon(timeout_ms=500, pipe_name=pipe_name)
    if not alive:
        return {
            "status": "info",
            "message": "El daemon no estaba en ejecución",
            "pipe": pipe_name,
        }

    resp = send_pipe_message(
        build_action_request(ACTION_SHUTDOWN),
        timeout_ms=2000,
        pipe_name=pipe_name,
    )
    logger.info("Daemon detenido vía Named Pipe | pipe=%s", pipe_name)
    return {
        "status": "success",
        "message": "Daemon detenido exitosamente",
        "pipe": pipe_name,
        "response": resp,
    }


def status_daemon(pipe_name: str = PIPE_NAME) -> Dict[str, Any]:
    """
    Consulta si el daemon está activo y reporta latencia y estado del libro.

    Args:
        pipe_name: Nombre del Named Pipe.

    Returns:
        Diccionario con daemon_running, latency_ms y engine_status.
    """
    alive, lat, details = ping_daemon(timeout_ms=1000, pipe_name=pipe_name)
    if alive:
        return {
            "status": "success",
            "daemon_running": True,
            "latency_ms": round(lat, 2),
            "pipe": pipe_name,
            "engine_status": details.get("engine_status") if details else None,
        }
    return {
        "status": "info",
        "daemon_running": False,
        "message": "Daemon no está en ejecución",
        "pipe": pipe_name,
    }
