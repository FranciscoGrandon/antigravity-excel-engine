# Creado por Francisco Grandón Vergara
"""
Antigravity Excel Engine Package.
The Enterprise-Grade, AI-Native Excel Automation Engine for Autonomous Agents.
"""

from .antigravity_excel_core import (
    AntigravityExcelEngine,
    LiveExcelCOMBackend,
    HeadlessOpenPyXLBackend,
    retry_on_excel_busy
)

# Aliases de conveniencia
ExcelEngine = AntigravityExcelEngine

__all__ = [
    "AntigravityExcelEngine",
    "ExcelEngine",
    "LiveExcelCOMBackend",
    "HeadlessOpenPyXLBackend",
    "retry_on_excel_busy"
]

__version__ = "1.0.0"
__author__ = "Francisco Grandón Vergara"
