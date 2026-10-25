# Creado por Francisco Grandón Vergara
"""
Protocolo canónico del canal IPC de Antigravity Excel Engine.

Define:
  - Constantes del Named Pipe (nombre, buffer, timeouts)
  - Constructores de mensajes de solicitud y respuesta
  - Validación básica de mensajes entrantes
"""

from typing import Any, Dict, List, Optional

# ──────────────────────────────────────────────────────────────────────────────
# Constantes del Named Pipe
# ──────────────────────────────────────────────────────────────────────────────

PIPE_NAME: str = r"\\.\pipe\antigravity_excel"
"""Nombre del Named Pipe del daemon residente en Windows."""

BUF_SIZE: int = 16 * 1024 * 1024  # 16 MB — soporta transferencias CSV masivas
"""Tamaño de buffer para lectura/escritura del Named Pipe (bytes)."""

DEFAULT_TIMEOUT_MS: int = 1000
"""Timeout por defecto en milisegundos para llamadas al Named Pipe."""

START_TIMEOUT_SEC: float = 5.0
"""Tiempo máximo en segundos para esperar que el daemon levante el pipe."""

# Comandos de control administrativo del daemon
ACTION_PING = "ping"
ACTION_SHUTDOWN = "shutdown"
ACTION_STATUS = "status"

# Comandos de operación del motor Excel
COMMANDS: List[str] = [
    "status",
    "get-csv",
    "get-ranges",
    "set-range",
    "set-cells",
    "copy-paste",
    "check-errors",
    "audit-quality",
    "aggregate",
    "add-column",
    "create-pivot",
    "create-chart",
    "apply-preset",
    "summary-table",
    "dedup",
    "flag-outliers",
    "clean",
    "curate",
    "wait-until-ready",
    "dispatch-com",
    "inspect-workbook",
    "inspect-charts",
    "extend-chart-series",
    "update-chart-series",
    "batch",
    "create-table",
    "inspect-table",
]

VALID_COMMANDS: List[str] = COMMANDS

# ──────────────────────────────────────────────────────────────────────────────
# Constructores de mensajes
# ──────────────────────────────────────────────────────────────────────────────


def build_command_request(
    command: str,
    args: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Construye un mensaje de solicitud de comando estándar.

    Args:
        command: Nombre del comando a ejecutar (ej. "get-csv", "aggregate").
        args: Argumentos del comando como diccionario.

    Returns:
        Diccionario JSON-serializable listo para enviar al daemon.
    """
    return {
        "command": command,
        "args": args or {},
    }


def build_action_request(
    action: str,
    payload: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Construye un mensaje de acción de control administrativo.

    Args:
        action: Nombre de la acción ("ping", "shutdown").
        payload: Datos adicionales opcionales.

    Returns:
        Diccionario JSON-serializable.
    """
    msg: Dict[str, Any] = {"action": action}
    if payload:
        msg.update(payload)
    return msg


def build_success_response(
    result: Any,
    warnings: Optional[List[str]] = None,
    latency_ms: Optional[float] = None,
) -> Dict[str, Any]:
    """
    Construye una respuesta de éxito estandarizada.

    Args:
        result: Resultado de la operación (dict, list, str, etc.).
        warnings: Lista opcional de advertencias no fatales.
        latency_ms: Latencia de procesamiento en ms (opcional).

    Returns:
        Diccionario de respuesta estándar.
    """
    resp: Dict[str, Any] = {
        "status": "success",
        "result": result,
    }
    if warnings:
        resp["warnings"] = warnings
    if latency_ms is not None:
        resp["latency_ms"] = round(latency_ms, 3)
    return resp


def build_error_response(
    error: str,
    command: Optional[str] = None,
    error_type: Optional[str] = None,
    latency_ms: Optional[float] = None,
) -> Dict[str, Any]:
    """
    Construye una respuesta de error estandarizada.

    Args:
        error: Mensaje de error descriptivo.
        command: Comando que produjo el error (opcional).
        error_type: Tipo de error ("ExcelLockError", "FormulaValidationError", etc.).
        latency_ms: Latencia de procesamiento en ms (opcional).

    Returns:
        Diccionario de respuesta de error estándar.
    """
    resp: Dict[str, Any] = {
        "status": "error",
        "error": error,
    }
    if command:
        resp["command"] = command
    if error_type:
        resp["error_type"] = error_type
    if latency_ms is not None:
        resp["latency_ms"] = round(latency_ms, 3)
    return resp


def validate_request(req: Dict[str, Any]) -> Optional[str]:
    """
    Valida la estructura básica de un mensaje de solicitud entrante.

    Args:
        req: Diccionario de mensaje a validar.

    Returns:
        None si el mensaje es válido; descripción del error si es inválido.
    """
    if not isinstance(req, dict):
        return "La solicitud debe ser un objeto JSON."

    has_action = "action" in req
    has_command = "command" in req

    if not has_action and not has_command:
        return "La solicitud debe contener 'action' o 'command'."

    if has_command:
        cmd = req["command"]
        if not isinstance(cmd, str) or not cmd.strip():
            return "El campo 'command' debe ser una cadena no vacía."
        args = req.get("args")
        if args is not None and not isinstance(args, dict):
            return "El campo 'args' debe ser un objeto JSON (dict)."

    return None
