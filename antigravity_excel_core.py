# Creado por Francisco Grandón Vergara
"""
Motor de automatización y manipulación inteligente de Excel para Agentes de IA y pipelines de datos Python.
Soporte dual: Live COM (Excel activo en pantalla) y Headless (OpenPyXL en disco).
"""

import os
import sys
import time
import csv
import io
import re
import functools
import datetime
import statistics
from typing import List, Dict, Any, Optional, Union

# Re-exports desde antigravity_excel_engine.backends
from antigravity_excel_engine.backends import (
    BaseExcelBackend,
    LiveExcelCOMBackend,
    HeadlessOpenPyXLBackend,
    retry_on_excel_busy,
    excel_fast_mode,
    get_com_clean_address,
)

# Re-export del engine orquestador
from antigravity_excel_engine.engine import AntigravityExcelEngine, ExcelEngine

# Re-exports desde antigravity_excel_engine.formatting
from antigravity_excel_engine.formatting import (
    FinancialStyleGuide,
    hex_to_rgb_int,
    apply_style_preset_openpyxl,
    apply_style_preset_com,
    apply_style_preset,
)

# Re-exports desde antigravity_excel_engine.formulas
from antigravity_excel_engine.formulas import (
    sanitize_formula_compatibility,
    propagate_formula_row,
    audit_openpyxl_formula_errors,
    audit_com_formula_errors,
)

# Re-exports desde antigravity_excel_engine.utils
from antigravity_excel_engine.utils.ranges import (
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
    range_boundaries,
    chunk_cell_addresses,
)
_chunk_cell_addresses = chunk_cell_addresses

from antigravity_excel_engine.utils.dates import (
    parse_excel_date_serial,
)

try:
    import win32com.client
    import pythoncom
    import pywintypes
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None

# Módulos COM/Win32 secundarios diferidos
win32file = None
win32con = None
win32gui = None
win32service = None
ctypes = None
uuid = None
wintypes = None

# Lazy Loaders para openpyxl, estilos, comentarios y módulos secundarios
_openpyxl_module = None
_openpyxl_styles = None
_openpyxl_comments = None

Font = None
PatternFill = None
Border = None
Side = None
Alignment = None
Comment = None


def _get_openpyxl():
    """Carga perezosa del módulo openpyxl con cacheo singleton."""
    global _openpyxl_module
    if _openpyxl_module is None:
        try:
            import openpyxl as _op
            _openpyxl_module = _op
        except ImportError:
            _openpyxl_module = False
    return _openpyxl_module if _openpyxl_module is not False else None


def _is_openpyxl_available() -> bool:
    """Verifica si openpyxl está instalado sin forzar su import si aún no ha sido cargado."""
    global _openpyxl_module
    if _openpyxl_module is not None:
        return _openpyxl_module is not False
    try:
        import importlib.util
        return importlib.util.find_spec("openpyxl") is not None
    except Exception:
        return _get_openpyxl() is not None


def _get_styles():
    global _openpyxl_styles, Font, PatternFill, Border, Side, Alignment
    if _openpyxl_styles is None:
        import openpyxl.styles as _st
        _openpyxl_styles = _st
        Font = _st.Font
        PatternFill = _st.PatternFill
        Border = _st.Border
        Side = _st.Side
        Alignment = _st.Alignment
    return _openpyxl_styles


def _get_comments():
    global _openpyxl_comments, Comment
    if _openpyxl_comments is None:
        import openpyxl.comments as _cm
        _openpyxl_comments = _cm
        Comment = _cm.Comment
    return _openpyxl_comments


def _get_win32gui():
    global win32gui
    if win32gui is None:
        try:
            import win32gui as _wgui
            win32gui = _wgui
        except ImportError:
            pass
    return win32gui


def _get_win32file():
    global win32file
    if win32file is None:
        try:
            import win32file as _wfile
            win32file = _wfile
        except ImportError:
            pass
    return win32file


def _get_win32con():
    global win32con
    if win32con is None:
        try:
            import win32con as _wcon
            win32con = _wcon
        except ImportError:
            pass
    return win32con


def _get_win32service():
    global win32service
    if win32service is None:
        try:
            import win32service as _wsvc
            win32service = _wsvc
        except ImportError:
            pass
    return win32service


def _get_ctypes():
    global ctypes, wintypes
    if ctypes is None:
        try:
            import ctypes as _ct
            from ctypes import wintypes as _wt
            ctypes = _ct
            wintypes = _wt
        except ImportError:
            pass
    return ctypes


def _get_uuid():
    global uuid
    if uuid is None:
        import uuid as _u
        uuid = _u
    return uuid


def __getattr__(name: str):
    """Acceso perezoso a símbolos públicos de openpyxl y win32 para máxima retrocompatibilidad."""
    if name == "openpyxl":
        val = _get_openpyxl()
        globals()["openpyxl"] = val
        return val
    elif name in ("Font", "PatternFill", "Border", "Side", "Alignment"):
        styles = _get_styles()
        val = getattr(styles, name, None)
        globals()[name] = val
        return val
    elif name == "Comment":
        comments = _get_comments()
        val = getattr(comments, "Comment", None)
        globals()["Comment"] = val
        return val
    elif name == "win32gui":
        val = _get_win32gui()
        globals()["win32gui"] = val
        return val
    elif name == "win32service":
        val = _get_win32service()
        globals()["win32service"] = val
        return val
    elif name == "win32file":
        val = _get_win32file()
        globals()["win32file"] = val
        return val
    elif name == "win32con":
        val = _get_win32con()
        globals()["win32con"] = val
        return val
    elif name == "ctypes":
        val = _get_ctypes()
        globals()["ctypes"] = val
        return val
    elif name == "wintypes":
        _get_ctypes()
        return wintypes
    elif name == "uuid":
        val = _get_uuid()
        globals()["uuid"] = val
        return val
    raise AttributeError(f"module '{__name__}' has no attribute '{name}'")


class AntigravityExcelEngine:
    """Fachada unificada con selector inteligente de backend."""
    def __init__(self, mode: str = "auto", file_path: Optional[str] = None):
        self.mode = mode.lower()
        self.file_path = file_path
        self.backend = self._init_backend()

    def _init_backend(self) -> BaseExcelBackend:
        if self.mode == "live":
            return LiveExcelCOMBackend(self.file_path)
        elif self.mode == "file":
            return HeadlessOpenPyXLBackend(self.file_path)
        elif self.mode == "auto":
            # 1. Check if Excel is active via standard COM ROT
            is_excel_active = False
            if win32com:
                try:
                    win32com.client.GetActiveObject("Excel.Application")
                    is_excel_active = True
                except Exception:
                    is_excel_active = False

            if is_excel_active:
                return LiveExcelCOMBackend(self.file_path)

            # 2. Check if the target file (or an Excel window) is active on desktop via oleacc
            if win32com and _get_win32gui():
                try:
                    tester = LiveExcelCOMBackend.__new__(LiveExcelCOMBackend)
                    desktop_app = tester._try_get_live_desktop_excel(self.file_path)
                    if desktop_app:
                        return LiveExcelCOMBackend(self.file_path)
                except Exception:
                    pass

            # 3. Fallback to Headless openpyxl for fast file processing
            has_openpyxl = _is_openpyxl_available()
            if self.file_path and has_openpyxl:
                return HeadlessOpenPyXLBackend(self.file_path)
            elif win32com:
                return LiveExcelCOMBackend(self.file_path)
            elif has_openpyxl:
                return HeadlessOpenPyXLBackend(self.file_path)
            else:
                raise ImportError("No se encontró backend disponible (pywin32 u openpyxl).")
        else:
            raise ValueError(f"Modo no reconocido: '{self.mode}'. Use 'auto', 'live' o 'file'.")

    def wait_until_ready(self, timeout_sec: float = 10.0, retry_interval: float = 0.25) -> Dict[str, Any]:
        return self.backend.wait_until_ready(timeout_sec=timeout_sec, retry_interval=retry_interval)

    def dispatch_com(
        self,
        target: str,
        member: str,
        args: Optional[List[Any]] = None,
        kwargs: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        return self.backend.dispatch_com(target=target, member=member, args=args, kwargs=kwargs)

    def __getattr__(self, name):
        return getattr(self.backend, name)


__all__ = [
    "AntigravityExcelEngine",
    "ExcelEngine",
    "BaseExcelBackend",
    "LiveExcelCOMBackend",
    "HeadlessOpenPyXLBackend",
    "retry_on_excel_busy",
    "excel_fast_mode",
    "get_com_clean_address",
    "FinancialStyleGuide",
    "hex_to_rgb_int",
    "apply_style_preset_openpyxl",
    "apply_style_preset_com",
    "apply_style_preset",
    "sanitize_formula_compatibility",
    "propagate_formula_row",
    "audit_openpyxl_formula_errors",
    "audit_com_formula_errors",
    "coordinate_from_string",
    "column_index_from_string",
    "get_column_letter",
    "range_boundaries",
    "chunk_cell_addresses",
    "parse_excel_date_serial",
]
