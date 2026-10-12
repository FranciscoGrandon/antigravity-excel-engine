# Creado por Francisco Grandón Vergara
"""
Subpaquete de backends de ejecución para Antigravity Excel Engine.
Proporciona la abstracción base y las implementaciones para Live COM (interactivo)
y Headless OpenPyXL (procesamiento en disco y sin dependencias de GUI).
"""

from .base import BaseExcelBackend
from .live_com import LiveExcelCOMBackend, retry_on_excel_busy, excel_fast_mode, get_com_clean_address
from .headless import HeadlessOpenPyXLBackend

__all__ = [
    "BaseExcelBackend",
    "LiveExcelCOMBackend",
    "HeadlessOpenPyXLBackend",
    "retry_on_excel_busy",
    "excel_fast_mode",
    "get_com_clean_address",
]
