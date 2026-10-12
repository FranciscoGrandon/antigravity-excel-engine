# Creado por Francisco Grandón Vergara
"""
Subpaquete IPC de Antigravity Excel Engine.
Contiene el protocolo de mensajes Named Pipe, el daemon residente
y el cliente IPC con funciones de alto nivel.
"""

from .protocol import (
    PIPE_NAME,
    BUF_SIZE,
    DEFAULT_TIMEOUT_MS,
    build_command_request,
    build_action_request,
    build_success_response,
    build_error_response,
    validate_request,
)
from .client import (
    send_pipe_message,
    ping_daemon,
    call_command,
    start_daemon_process,
    stop_daemon_process,
    status_daemon,
)
from .daemon import AntigravityExcelDaemon

__all__ = [
    # Protocol
    "PIPE_NAME",
    "BUF_SIZE",
    "DEFAULT_TIMEOUT_MS",
    "build_command_request",
    "build_action_request",
    "build_success_response",
    "build_error_response",
    "validate_request",
    # Client
    "send_pipe_message",
    "ping_daemon",
    "call_command",
    "start_daemon_process",
    "stop_daemon_process",
    "status_daemon",
    # Daemon
    "AntigravityExcelDaemon",
]
