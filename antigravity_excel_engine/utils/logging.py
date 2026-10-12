# Creado por Francisco Grandón Vergara
"""
Configuración centralizada de logging y utilidades de registro para Antigravity Excel Engine.
"""

import logging
import sys
import time
from contextlib import contextmanager
from typing import Any, Dict, Optional

LOGGER_NAME = "antigravity_excel"


def get_logger(name: str = LOGGER_NAME) -> logging.Logger:
    """
    Obtiene o inicializa el logger centralizado 'antigravity_excel'.
    Configura un StreamHandler por defecto si el logger no tiene manejadores asociados.
    """
    logger = logging.getLogger(name)
    if not logger.handlers:
        logger.setLevel(logging.INFO)
        formatter = logging.Formatter(
            fmt="%(asctime)s [%(levelname)s] [%(name)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.propagate = False
    return logger


_default_logger = get_logger()


def log_operation(operation: str, details: Optional[Dict[str, Any]] = None, level: int = logging.INFO) -> None:
    """Registra una operación ejecutada en el motor con detalles opcionales."""
    detail_str = f" | {details}" if details else ""
    _default_logger.log(level, f"[OP] {operation}{detail_str}")


def log_exception(exc: Exception, context: Optional[str] = None) -> None:
    """Registra una excepción capturada con contexto opcional."""
    ctx_str = f" en contexto: {context}" if context else ""
    _default_logger.error(f"[ERROR] {type(exc).__name__}: {exc}{ctx_str}", exc_info=True)


@contextmanager
def log_duration(operation_name: str, logger: Optional[logging.Logger] = None):
    """Context manager para medir y loguear la duración de una operación."""
    lg = logger or _default_logger
    start = time.perf_counter()
    try:
        yield
    finally:
        elapsed_ms = (time.perf_counter() - start) * 1000.0
        lg.debug(f"[DURATION] {operation_name}: {elapsed_ms:.2f} ms")
