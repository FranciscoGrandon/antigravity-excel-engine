# Creado por Francisco Grandón Vergara
"""
Motor de automatización y manipulación inteligente de Excel para Agentes de IA y pipelines de datos Python.
Soporte dual: Live COM (Excel activo en pantalla) y Headless (OpenPyXL en disco).
"""

import os
import time
import csv
import io
import re
import functools
import datetime
import statistics
from contextlib import contextmanager
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Union


@contextmanager
def excel_fast_mode(app):
    """
    Context manager de alto rendimiento para Live Excel COM.
    Desactiva temporalmente el refresco de pantalla, eventos de aplicación y
    el recálculo automático para acelerar drásticamente escrituras masivas.
    Restaura el estado previo y fuerza un recálculo final al salir.
    """
    if not app:
        yield
        return

    prev_screen_updating = True
    prev_display_alerts = True
    prev_enable_events = True
    prev_calc = -4105 # xlCalculationAutomatic

    try:
        prev_screen_updating = getattr(app, "ScreenUpdating", True)
        prev_display_alerts = getattr(app, "DisplayAlerts", True)
        prev_enable_events = getattr(app, "EnableEvents", True)
        prev_calc = getattr(app, "Calculation", -4105)

        app.ScreenUpdating = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.Calculation = -4135 # xlCalculationManual
    except Exception:
        pass

    try:
        yield
    finally:
        try:
            app.Calculation = prev_calc
            app.EnableEvents = prev_enable_events
            app.DisplayAlerts = prev_display_alerts
            app.ScreenUpdating = prev_screen_updating
            app.Calculate()
        except Exception:
            pass


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

# ---------------------------------------------------------------------------
# Utilidades de celda y coordenadas puras (100% compatibles con openpyxl, sin sobrecarga de import)
# ---------------------------------------------------------------------------
_COORD_RE = re.compile(r"^[$]?([A-Za-z]{1,3})[$]?(\d+)$")
_ABSOLUTE_RE = re.compile(r"^[$]?(?P<min_col>[A-Za-z]{1,3})?[$]?(?P<min_row>\d+)?(?::[$]?(?P<max_col>[A-Za-z]{1,3})?[$]?(?P<max_row>\d+)?)?$", re.IGNORECASE)


def coordinate_from_string(coord_string: str):
    """Convierte una coordenada como 'B12' o '$B$12' en una tupla ('B', 12)."""
    m = _COORD_RE.match(coord_string)
    if not m:
        raise ValueError(f"Invalid cell coordinates ({coord_string})")
    col, row = m.groups()
    r = int(row)
    if not r:
        raise ValueError(f"There is no row 0 ({coord_string})")
    return col, r


@functools.lru_cache(maxsize=None)
def column_index_from_string(col: str) -> int:
    """Convierte nombre de columna alfanumérico (ej. 'A', 'Z', 'AA') en índice decimal 1-based."""
    col_clean = str(col).strip().upper()
    if len(col_clean) > 3:
        raise ValueError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
    idx = 0
    for char in col_clean:
        if not ("A" <= char <= "Z"):
            raise ValueError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
        idx = idx * 26 + (ord(char) - ord("A") + 1)
    if not (1 <= idx <= 18278):
        raise ValueError(f"'{col}' is not a valid column name. Column names are from A to ZZZ")
    return idx


@functools.lru_cache(maxsize=None)
def get_column_letter(col_idx: int) -> str:
    """Convierte un índice de columna 1-based en su letra correspondiente ('A', 'AA', etc.)."""
    if not (1 <= col_idx <= 18278):
        raise ValueError(f"Invalid column index {col_idx}")
    result = []
    while col_idx > 0:
        col_idx, remainder = divmod(col_idx - 1, 26)
        result.append(chr(ord("A") + remainder))
    return "".join(reversed(result))


def range_boundaries(range_string: str):
    """Convierte una cadena de rango (ej. 'A1:B10', 'C5') en (min_col, min_row, max_col, max_row)."""
    m = _ABSOLUTE_RE.match(range_string)
    if not m:
        raise ValueError(f"{range_string} is not a valid coordinate or range")
    min_col, min_row, max_col, max_row = m.group("min_col"), m.group("min_row"), m.group("max_col"), m.group("max_row")
    c_min = column_index_from_string(min_col) if min_col else None
    r_min = int(min_row) if min_row else None
    c_max = column_index_from_string(max_col) if max_col else c_min
    r_max = int(max_row) if max_row else r_min
    return c_min, r_min, c_max, r_max


def _chunk_cell_addresses(cell_list: List[str], max_chars: int = 200) -> List[str]:
    """
    Agrupa direcciones de celdas ('A1', 'B2', ...) en cadenas unificadas separadas por comas
    de longitud segura para llamadas multi-rango en COM (evita el límite de ~255 caracteres de Range).
    """
    chunks = []
    current = []
    current_len = 0
    for c in cell_list:
        c_str = str(c).strip()
        if not c_str:
            continue
        c_len = len(c_str) + (1 if current else 0)
        if current and (current_len + c_len > max_chars):
            chunks.append(",".join(current))
            current = [c_str]
            current_len = len(c_str)
        else:
            current.append(c_str)
            current_len += c_len
    if current:
        chunks.append(",".join(current))
    return chunks



# ---------------------------------------------------------------------------
# Lazy Loaders para openpyxl, estilos, comentarios y módulos secundarios
# ---------------------------------------------------------------------------
_openpyxl_module = None
_openpyxl_styles = None
_openpyxl_comments = None


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


def hex_to_rgb_int(hex_str: str) -> int:
    """Convierte '#RRGGBB' a entero BGR para Excel COM (0x00BBGGRR)."""
    clean_hex = hex_str.lstrip('#')
    if len(clean_hex) == 6:
        r = int(clean_hex[0:2], 16)
        g = int(clean_hex[2:4], 16)
        b = int(clean_hex[4:6], 16)
        return r + (g << 8) + (b << 16)
    return 0


def retry_on_excel_busy(max_retries: int = 10, delay_sec: float = 0.5):
    """
    Mecanismo adaptativo con backoff para tolerancia a fallos por edición interactiva (RPC_E_SERVERCALL_RETRYLATER).
    """
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            retries = 0
            while retries < max_retries:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    is_busy = False
                    if pywintypes and isinstance(e, pywintypes.com_error):
                        hresult = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                        if hresult in (-2147417846, -2147418111, -2146777998, 0x8001010A, 0x80010001, 0x800AC472):
                            is_busy = True
                    elif pythoncom and isinstance(e, pythoncom.com_error):
                        is_busy = True

                    if is_busy:
                        retries += 1
                        time.sleep(delay_sec)
                        continue
                    raise
            raise RuntimeError(
                f"Excel permaneció bloqueado en modo edición tras {max_retries * delay_sec:.1f}s. "
                f"Por favor presiona [ENTER] o [ESC] en Excel para continuar."
            )
        return wrapper
    return decorator


class FinancialStyleGuide:
    """
    Constantes y directrices de modelado financiero estándar.
    Proporciona formatos numéricos y paletas visuales corporativas para
    modelado financiero profesional en Excel.
    """
    CURRENCY_FORMAT = "$ #,##0.00;($ #,##0.00);\"-\""
    CURRENCY_FORMAT_LOCAL = "$ #.##0,00;($ #.##0,00);\"-\""
    PERCENT_FORMAT = "0.0%"
    PERCENT_FORMAT_LOCAL = "0,0%"
    COLOR_INPUT = "#1F4E79"            # Azul corporativo para celdas de entrada
    COLOR_FORMULA = "#000000"          # Negro para fórmulas
    COLOR_ASSUMPTION_FILL = "#FFF2CC"  # Amarillo marfil para supuestos / escenarios

    @staticmethod
    def apply_number_format_resilient(excel_cell_or_range, fmt_type_or_mask: str):
        """Aplica formato numérico probando NumberFormatLocal y NumberFormat para máxima compatibilidad regional."""
        if not fmt_type_or_mask:
            return
        t = fmt_type_or_mask.lower()
        if t == "currency":
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.CURRENCY_FORMAT_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.CURRENCY_FORMAT
        elif t == "percent":
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.PERCENT_FORMAT_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.PERCENT_FORMAT
        else:
            try:
                excel_cell_or_range.NumberFormatLocal = fmt_type_or_mask
            except Exception:
                excel_cell_or_range.NumberFormat = fmt_type_or_mask


def sanitize_formula_compatibility(formula: str) -> str:
    """
    Detecta y normaliza nombres de funciones Excel modernas (post-2007) que
    requieren el prefijo de compatibilidad XML '_xlfn.' en bibliotecas como openpyxl.
    Funciones tratadas: TEXTJOIN, CONCAT, IFS, SWITCH, MAXIFS, MINIFS.
    Si ya traen el prefijo, se preservan sin duplicarlo.
    """
    if not formula or not isinstance(formula, str):
        return formula

    post_2007_funcs = ("TEXTJOIN", "CONCAT", "IFS", "SWITCH", "MAXIFS", "MINIFS")

    pattern = re.compile(
        r'(?<!_xlfn\.)\b(' + '|'.join(post_2007_funcs) + r')\b(?=\s*\()',
        flags=re.IGNORECASE
    )

    def _replace_match(m):
        func_name = m.group(1).upper()
        return f"_xlfn.{func_name}"

    return pattern.sub(_replace_match, formula)


def get_com_clean_address(excel_range_or_cell) -> str:
    """Obtiene la dirección en formato limpio (A1 o A1:B2) de un objeto Range o Cell de win32com."""
    try:
        addr = excel_range_or_cell.Address
        if callable(addr):
            return str(addr(False, False)).replace("$", "")
        return str(addr).replace("$", "")
    except Exception:
        return str(getattr(excel_range_or_cell, "Address", "")).replace("$", "")


def propagate_formula_row(formula_template: str, row_idx: int, base_row: int = 2) -> str:
    """
    Propaga una plantilla de fórmula para una fila destino `row_idx`, respetando referencias
    absolutas y relativas.
    - Soporta placeholders explícitos `{r}` y `{row}`.
    - Si la fórmula contiene referencias relativas a la fila base (por defecto fila 2, ej. `A2`, `$B2`, `Sheet!C2`),
      sustituye dinámicamente el número de fila base por `row_idx`.
    - Preserva estrictamente intactas las referencias absolutas en fila (ej. `$C$4`, `A$2`, `$A$2`, `'Params'!$C$4`).
    """
    if not formula_template or not isinstance(formula_template, str):
        return formula_template

    if "{r}" in formula_template:
        return formula_template.replace("{r}", str(row_idx))
    if "{row}" in formula_template:
        return formula_template.replace("{row}", str(row_idx))

    # Identificar celdas relativas en la fila base_row:
    # 1. Columna con opcional '$' inicial y 1-3 letras: (\$?[A-Za-z]{1,3})
    # 2. Número de fila base_row exacto sin '$' entre columna y número
    # 3. No seguido de dígitos adicionales ni de paréntesis abierto (evitar funciones como LOG2)
    pattern = rf"(?<![\$A-Za-z0-9_])(\$?[A-Za-z]{{1,3}}){base_row}(?!\d|\s*\()"
    return re.sub(pattern, rf"\g<1>{row_idx}", formula_template)


def parse_excel_date_serial(value: Any) -> Optional[int]:
    """
    Convierte un valor de fecha a su número serial exacto de Excel (días transcurridos desde 1899-12-30).
    Soporta:
      - Cadenas con formato ISO 'YYYY-MM-DD' o 'YYYY/MM/DD'.
      - Formatos 'DD/MM/YYYY' o 'MM/DD/YYYY' (con o sin marcas de hora como '3:00:00 AM' o '15:30:00').
      - Formatos ISO con hora ('YYYY-MM-DDTHH:MM:SS').
      - Objetos nativos datetime.datetime y datetime.date.
      - Enteros o flotantes que ya correspondan a seriales válidos de Excel.
    Retorna None si no es convertible o es nulo.
    """
    if value is None:
        return None

    excel_base = datetime.date(1899, 12, 30)

    if isinstance(value, (datetime.datetime, datetime.date)):
        d = value.date() if isinstance(value, datetime.datetime) else value
        return (d - excel_base).days

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 2958465:
            return int(value)
        return None

    val_str = str(value).strip()
    if not val_str:
        return None

    # 1. ISO YYYY-MM-DD o YYYY/MM/DD con o sin hora
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$", val_str)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return (datetime.date(y, m, d) - excel_base).days
        except Exception:
            pass

    # 2. DD/MM/YYYY o MM/DD/YYYY con 4 dígitos en el año con o sin hora
    m_slash = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[ T].*)?$", val_str)
    if m_slash:
        n1, n2, y = int(m_slash.group(1)), int(m_slash.group(2)), int(m_slash.group(3))
        if n1 > 12 and n2 <= 12:
            d, m = n1, n2
        elif n2 > 12 and n1 <= 12:
            m, d = n1, n2
        else:
            d, m = n1, n2
        try:
            return (datetime.date(y, m, d) - excel_base).days
        except Exception:
            pass

    # 3. Formato con 2 dígitos en el año (DD/MM/YY o MM/DD/YY)
    m_slash2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2})(?:[ T].*)?$", val_str)
    if m_slash2:
        n1, n2, y_short = int(m_slash2.group(1)), int(m_slash2.group(2)), int(m_slash2.group(3))
        y = y_short + 2000 if y_short < 50 else y_short + 1900
        if n1 > 12 and n2 <= 12:
            d, m = n1, n2
        elif n2 > 12 and n1 <= 12:
            m, d = n1, n2
        else:
            d, m = n1, n2
        try:
            return (datetime.date(y, m, d) - excel_base).days
        except Exception:
            pass

    # 4. Fallback exhaustivo con strptime
    formats = [
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %I:%M:%S %p", "%d/%m/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"
    ]
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(val_str, fmt)
            return (dt.date() - excel_base).days
        except Exception:
            continue

    return None


class BaseExcelBackend(ABC):
    @abstractmethod
    def get_cell_ranges(self, ranges: List[str], sheet: Optional[str] = None, include_formulas: bool = False, include_styles: bool = False) -> Dict[str, Any]: pass
    
    @abstractmethod
    def set_cell_range(self, start_cell: str, values: List[List[Any]], sheet: Optional[str] = None, formulas: Optional[List[List[str]]] = None, styles: Optional[Dict[str, Any]] = None, autofit: bool = False) -> Dict[str, Any]: pass
    
    @abstractmethod
    def set_cells(self, cells_map: Dict[str, Dict[str, Any]], sheet: Optional[str] = None, autofit: bool = False) -> Dict[str, Any]: pass

    @abstractmethod
    def get_range_as_csv(self, range_str: str, sheet: Optional[str] = None, delimiter: str = ";") -> str: pass
    
    @abstractmethod
    def copy_paste_range(self, src_range: str, dst_start_cell: str, src_sheet: Optional[str] = None, dst_sheet: Optional[str] = None, paste_type: str = "all") -> Dict[str, Any]: pass
    
    @abstractmethod
    def check_formula_errors(self, range_str: str, sheet: Optional[str] = None) -> List[Dict[str, Any]]: pass

    @abstractmethod
    def audit_data_quality(self, range_str: Optional[str] = None, sheet: Optional[str] = None) -> Dict[str, Any]: pass

    @staticmethod
    def _calculate_placement_coordinates(
        start_col: int,
        start_row: int,
        num_cols: int,
        num_rows: int,
        placement: str = "auto",
        anchor: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Helper de cálculo de posición geométrica no invasiva anti-solapamiento.
        Si anchor está especificado (ej. 'I3'), utiliza esa celda.
        Si placement='auto':
          - si num_cols <= 8: ubica a la derecha (dejando 2 columnas libres de margen, misma fila inicial).
          - si num_cols > 8: ubica abajo (dejando 2 filas libres de margen, misma columna inicial).
        Si placement='right': a la derecha dejando 2 columnas libres.
        Si placement='bottom': abajo dejando 2 filas libres.
        """
        if anchor:
            # Parse anchor string (e.g. 'I3')
            m = re.match(r"^([A-Za-z]+)(\d+)$", anchor.strip())
            if m:
                col_letters = m.group(1).upper()
                r_num = int(m.group(2))
                c_num = 0
                for char in col_letters:
                    c_num = c_num * 26 + (ord(char) - ord('A') + 1)
                return {
                    "col": c_num,
                    "row": r_num,
                    "cell": f"{col_letters}{r_num}",
                    "placement": "custom"
                }

        plc = placement.lower() if placement else "auto"
        if plc == "auto":
            plc = "right" if num_cols <= 8 else "bottom"

        if plc == "right":
            target_col = start_col + num_cols + 2 # 2 columnas libres de margen
            target_row = start_row
        elif plc == "bottom":
            target_col = start_col
            target_row = start_row + num_rows + 2 # 2 filas libres de margen
        else:
            target_col = start_col + num_cols + 2
            target_row = start_row

        # Convert target_col to letters
        col_temp = target_col
        letters = []
        while col_temp > 0:
            col_temp, rem = divmod(col_temp - 1, 26)
            letters.append(chr(ord('A') + rem))
        col_str = "".join(reversed(letters))

        return {
            "col": target_col,
            "row": target_row,
            "cell": f"{col_str}{target_row}",
            "placement": plc
        }

    @abstractmethod
    def create_pivot_table(
        self,
        source: Union[str, Dict[str, Any]],
        rows: List[str],
        cols: Optional[List[str]] = None,
        value_field: Optional[str] = None,
        value_func: str = "sum",
        value_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT,
        dest_sheet: Optional[str] = None,
        dest_cell: str = "A3",
        table_name: Optional[str] = None,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]: pass

    @abstractmethod
    def create_chart(
        self,
        source: Union[str, Dict[str, Any]],
        chart_type: str = "column_clustered",
        title: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        placement: str = "auto",
        anchor: Optional[str] = None,
        width: int = 480,
        height: int = 300,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]: pass

    @abstractmethod
    def aggregate(self, group_by_col: str, metric_col: str, agg_func: str = "sum", sheet: Optional[str] = None, top_n: int = 10, ascending: bool = False) -> Dict[str, Any]: pass

    @abstractmethod
    def add_calculated_column(self, header: str, formula_template: str, number_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT, sheet: Optional[str] = None, autofit: bool = True) -> Dict[str, Any]: pass
    
    @abstractmethod
    def apply_style_preset(self, range_str: str, preset: str, sheet: Optional[str] = None) -> Dict[str, Any]: pass

    @abstractmethod
    def create_summary_table(self, title: str, headers: List[str], rows_spec: List[Dict[str, Any]], dest_sheet: str, start_cell: str = "B2", parameters: Optional[List[Dict[str, Any]]] = None) -> Dict[str, Any]: pass

    @abstractmethod
    def save(self, file_path: Optional[str] = None) -> str: pass
    
    @abstractmethod
    def close(self): pass
    
    @abstractmethod
    def remove_duplicates(self, range_or_table: str, key_columns: Optional[List[str]] = None, sheet: Optional[str] = None) -> Dict[str, Any]: pass

    @abstractmethod
    def detect_and_flag_outliers(self, column: str, method: str = "iqr", range_or_table: Optional[str] = None, sheet: Optional[str] = None, flag_style: str = "warning", add_comment: bool = True) -> Dict[str, Any]: pass

    @abstractmethod
    def clean_table_dataset(self, target_range_or_table: str, rules: Dict[str, Any], sheet: Optional[str] = None) -> Dict[str, Any]: pass

    @abstractmethod
    def curate_pipeline(self, target_range_or_table: str, pipeline_spec: Dict[str, Any], sheet: Optional[str] = None) -> Dict[str, Any]: pass

    @abstractmethod
    def status(self) -> Dict[str, Any]: pass



class LiveExcelCOMBackend(BaseExcelBackend):
    def __init__(self, file_path: Optional[str] = None):
        if not win32com:
            raise ImportError("win32com no está instalado en este entorno.")
        self.file_path = file_path
        self.app = self._init_app()
        self.wb = self._init_workbook(file_path)

    @retry_on_excel_busy()
    def _init_app(self):
        # 1. Try standard GetActiveObject
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            app.Visible = True
            return app
        except Exception:
            pass

        # 2. Try native accessibility hook to interactive desktop Excel window (oleacc)
        app = self._try_get_live_desktop_excel(self.file_path)
        if app:
            return app

        # 3. Fallback to Dispatch
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = True
        return app

    def _try_get_live_desktop_excel(self, target_file: Optional[str] = None):
        """Intenta engancharse a la instancia visual de Excel abierta por el usuario en WinSta0\\Default."""
        _wgui = _get_win32gui()
        _wsvc = _get_win32service()
        _ct = _get_ctypes()
        _u = _get_uuid()
        if not (_wgui and _wsvc and _ct and _u and pythoncom):
            return None
        try:
            import threading
            result = []
            target_name = os.path.basename(target_file).lower() if target_file else None

            def worker():
                try:
                    pythoncom.CoInitialize()
                    hwinsta = win32service.OpenWindowStation('WinSta0', False, 0x10000000)
                    hwinsta.SetProcessWindowStation()
                    hdesk = win32service.OpenDesktop('Default', 0, False, 0x10000000)
                    hdesk.SetThreadDesktop()

                    excel_hwnds = []
                    def enum_windows(hwnd, _):
                        cls = win32gui.GetClassName(hwnd)
                        if cls == 'XLMAIN':
                            title = win32gui.GetWindowText(hwnd).lower()
                            # Prioritize window matching filename
                            if target_name and target_name in title:
                                excel_hwnds.insert(0, hwnd)
                            else:
                                excel_hwnds.append(hwnd)

                    win32gui.EnumDesktopWindows(hdesk, enum_windows, None)

                    oleacc = ctypes.windll.oleacc
                    oleacc.ObjectFromLresult.restype = ctypes.c_long
                    oleacc.ObjectFromLresult.argtypes = [
                        wintypes.LPARAM,
                        ctypes.c_char_p,
                        wintypes.WPARAM,
                        ctypes.POINTER(ctypes.c_void_p)
                    ]

                    for x_hwnd in excel_hwnds:
                        excel7_hwnd = None
                        def enum_child(h, _):
                            nonlocal excel7_hwnd
                            if win32gui.GetClassName(h) == 'EXCEL7':
                                excel7_hwnd = h
                        win32gui.EnumChildWindows(x_hwnd, enum_child, None)

                        if excel7_hwnd:
                            lres = win32gui.SendMessage(excel7_hwnd, 0x003D, 0, -16 & 0xFFFFFFFF)
                            iid_bytes = uuid.UUID('{00020400-0000-0000-C000-000000000046}').bytes_le
                            p_disp = ctypes.c_void_p()
                            hr = oleacc.ObjectFromLresult(lres, iid_bytes, 0, ctypes.byref(p_disp))
                            if hr == 0 and p_disp.value:
                                unk = pythoncom.ObjectFromAddress(p_disp.value)
                                disp = unk.QueryInterface(pythoncom.IID_IDispatch)
                                # Marshal to stream for main thread
                                stream = pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, disp)
                                result.append(stream)
                                return
                except Exception:
                    pass

            t = threading.Thread(target=worker)
            t.start()
            t.join(timeout=3.0)

            if result:
                pythoncom.CoInitialize()
                stream = result[0]
                disp = pythoncom.CoGetInterfaceAndReleaseStream(stream, pythoncom.IID_IDispatch)
                window = win32com.client.Dispatch(disp)
                if window and hasattr(window, 'Application'):
                    return window.Application
        except Exception:
            pass
        return None

    @retry_on_excel_busy()
    def _init_workbook(self, file_path: Optional[str]):
        if file_path:
            abs_p = os.path.abspath(file_path).lower()
            base_name = os.path.basename(file_path).lower()

            for w in self.app.Workbooks:
                try:
                    w_full = str(w.FullName).lower()
                    w_name = str(w.Name).lower()
                    if w_full == abs_p or w_name == base_name or base_name in w_full:
                        return w
                except Exception:
                    pass

            if os.path.exists(abs_p):
                return self.app.Workbooks.Open(abs_p)
            else:
                wb = self.app.Workbooks.Add()
                wb.SaveAs(abs_p)
                return wb

        if self.app.ActiveWorkbook:
            return self.app.ActiveWorkbook
        return self.app.Workbooks.Add()

    def _get_sheet(self, sheet_name: Optional[str]):
        if sheet_name:
            for s in self.wb.Sheets:
                if str(s.Name).lower() == str(sheet_name).lower():
                    return s
            new_ws = self.wb.Sheets.Add(After=self.wb.Sheets(self.wb.Sheets.Count))
            new_ws.Name = sheet_name
            return new_ws
        return self.wb.ActiveSheet

    @retry_on_excel_busy()
    def status(self) -> Dict[str, Any]:
        return {
            "backend": "live_com",
            "active_workbook": self.wb.Name if self.wb else None,
            "active_sheet": self.wb.ActiveSheet.Name if (self.wb and self.wb.ActiveSheet) else None,
            "sheets": [s.Name for s in self.wb.Sheets] if self.wb else []
        }

    @retry_on_excel_busy()
    def get_cell_ranges(self, ranges: List[str], sheet: Optional[str] = None, include_formulas: bool = False, include_styles: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        res = {}
        for r in ranges:
            target_ws = ws
            clean_range = r
            if "!" in r:
                s_name, clean_range = r.split("!", 1)
                target_ws = self.wb.Sheets(s_name)

            rng = target_ws.Range(clean_range)
            vals = rng.Value
            if not isinstance(vals, tuple):
                vals = [[vals]]
            elif not isinstance(vals[0], tuple):
                vals = [[x for x in vals]]
            else:
                vals = [list(row) for row in vals]

            res[r] = {"values": vals}
            if include_formulas:
                forms = rng.Formula
                if not isinstance(forms, tuple):
                    forms = [[forms]]
                elif not isinstance(forms[0], tuple):
                    forms = [[x for x in forms]]
                else:
                    forms = [list(row) for row in forms]
                res[r]["formulas"] = forms

        return res

    @retry_on_excel_busy()
    def set_cell_range(self, start_cell: str, values: List[List[Any]], sheet: Optional[str] = None, formulas: Optional[List[List[str]]] = None, styles: Optional[Dict[str, Any]] = None, autofit: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        start_rng = ws.Range(start_cell)
        row_count = len(values) if values else (len(formulas) if formulas else 0)
        col_count = len(values[0]) if (values and len(values) > 0) else (len(formulas[0]) if (formulas and len(formulas) > 0) else 0)
        
        if row_count == 0 or col_count == 0:
            return {"status": "error", "message": "No data provided"}

        end_rng = ws.Cells(start_rng.Row + row_count - 1, start_rng.Column + col_count - 1)
        target_rng = ws.Range(start_rng, end_rng)

        if values:
            target_rng.Value = [tuple(row) for row in values]
        if formulas:
            target_rng.Formula = [tuple(row) for row in formulas]

        if autofit:
            target_rng.Columns.AutoFit()

        self.app.Calculate()
        return {"status": "success", "range": target_rng.Address(False, False), "rows": row_count, "cols": col_count}

    @retry_on_excel_busy()
    def set_cells(self, cells_map: Dict[str, Dict[str, Any]], sheet: Optional[str] = None, autofit: bool = False) -> Dict[str, Any]:
        """Aplica un mapa declarativo de celdas A1 -> {value, formula, note, cellStyles} con recálculo atómico."""
        ws = self._get_sheet(sheet)
        formula_results = {}
        for coord, cell_info in cells_map.items():
            cell = ws.Range(coord)
            if "formula" in cell_info and cell_info["formula"]:
                cell.Formula = cell_info["formula"]
            elif "value" in cell_info:
                cell.Value = cell_info["value"]

            if "note" in cell_info and cell_info["note"]:
                try:
                    cell.ClearComments()
                    cell.AddComment(cell_info["note"])
                except Exception:
                    pass

            styles = cell_info.get("cellStyles", {})
            if styles:
                if "fontWeight" in styles:
                    cell.Font.Bold = (styles["fontWeight"] == "bold")
                if "fontStyle" in styles:
                    cell.Font.Italic = (styles["fontStyle"] == "italic")
                if "fontSize" in styles:
                    cell.Font.Size = int(styles["fontSize"])
                if "fontColor" in styles:
                    cell.Font.Color = hex_to_rgb_int(styles["fontColor"])
                if "backgroundColor" in styles:
                    cell.Interior.Color = hex_to_rgb_int(styles["backgroundColor"])
                if "numberFormat" in styles:
                    cell.NumberFormat = styles["numberFormat"]
                if "horizontalAlignment" in styles:
                    align_map = {"left": -4131, "center": -4108, "right": -4152}
                    val = align_map.get(styles["horizontalAlignment"].lower())
                    if val:
                        cell.HorizontalAlignment = val

        self.app.Calculate()
        for coord, cell_info in cells_map.items():
            if "formula" in cell_info:
                formula_results[coord] = ws.Range(coord).Value

        if autofit:
            ws.Columns.AutoFit()

        return {"status": "success", "cells_updated": len(cells_map), "formula_results": formula_results}

    @retry_on_excel_busy()
    def get_range_as_csv(self, range_str: str, sheet: Optional[str] = None, delimiter: str = ";") -> str:
        data = self.get_cell_ranges([range_str], sheet=sheet)[range_str]["values"]
        output = io.StringIO()
        writer = csv.writer(output, delimiter=delimiter, lineterminator="\n")
        for row in data:
            writer.writerow(["" if x is None else str(x) for x in row])
        return output.getvalue().strip()

    @retry_on_excel_busy()
    def copy_paste_range(self, src_range: str, dst_start_cell: str, src_sheet: Optional[str] = None, dst_sheet: Optional[str] = None, paste_type: str = "all") -> Dict[str, Any]:
        ws_src = self._get_sheet(src_sheet)
        ws_dst = self._get_sheet(dst_sheet)
        ws_src.Range(src_range).Copy()
        dst_rng = ws_dst.Range(dst_start_cell)

        paste_map = {
            "all": -4104,       # xlPasteAll
            "values": -4163,    # xlPasteValues
            "formulas": -4123,  # xlPasteFormulas
            "formats": -4122    # xlPasteFormats
        }
        dst_rng.PasteSpecial(Paste=paste_map.get(paste_type, -4104))
        self.app.CutCopyMode = False
        return {"status": "success", "source": src_range, "destination": dst_start_cell}

    @retry_on_excel_busy()
    def check_formula_errors(self, range_str: str, sheet: Optional[str] = None) -> List[Dict[str, Any]]:
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_str)
        errors = []
        err_map = {
            -2146826281: "#DIV/0!",
            -2146826246: "#N/A",
            -2146826259: "#NAME?",
            -2146826288: "#NULL!",
            -2146826252: "#NUM!",
            -2146826265: "#REF!",
            -2146826273: "#VALUE!"
        }
        for cell in rng:
            val = cell.Value
            if isinstance(val, int) and val in err_map:
                errors.append({
                    "cell": cell.Address(False, False),
                    "error": err_map[val],
                    "formula": cell.Formula
                })
        return errors

    @retry_on_excel_busy()
    def audit_data_quality(self, range_str: Optional[str] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        if range_str:
            rng = ws.Range(range_str)
        else:
            rng = ws.UsedRange

        header_whitespace = []
        numbers_as_text = []
        inconsistent_types = []
        total_formulas = 0
        total_cells = rng.Count
        
        # Check first row (headers)
        rows_cnt = rng.Rows.Count
        cols_cnt = rng.Columns.Count
        
        for c_idx in range(1, cols_cnt + 1):
            h_cell = rng.Cells(1, c_idx)
            val = str(h_cell.Value or "")
            if val and (val.startswith(" ") or val.endswith(" ") or "  " in val):
                header_whitespace.append({
                    "cell": h_cell.Address(False, False),
                    "header": val,
                    "trimmed": val.strip()
                })

        # Check column types and text numbers
        for c_idx in range(1, cols_cnt + 1):
            h_name = str(rng.Cells(1, c_idx).Value or f"Col_{c_idx}")
            col_types = set()
            for r_idx in range(2, min(rows_cnt + 1, 1000)):
                cell = rng.Cells(r_idx, c_idx)
                val = cell.Value
                formula = str(cell.Formula or "")
                if formula.startswith("="):
                    total_formulas += 1
                if val is not None:
                    col_types.add(type(val).__name__)
                    if isinstance(val, str) and val.strip().replace(".", "", 1).replace("-", "", 1).isdigit():
                        if len(numbers_as_text) < 50:
                            numbers_as_text.append({
                                "cell": cell.Address(False, False),
                                "column": h_name,
                                "value": val
                            })
            if len(col_types) > 1:
                inconsistent_types.append({
                    "column_index": c_idx,
                    "header": h_name,
                    "types_detected": list(col_types)
                })

        return {
            "sheet": ws.Name,
            "range_audited": rng.Address(False, False),
            "total_cells": total_cells,
            "has_dynamic_formulas": (total_formulas > 0),
            "formula_count": total_formulas,
            "header_whitespace_issues": header_whitespace,
            "numbers_stored_as_text": numbers_as_text,
            "inconsistent_column_types": inconsistent_types
        }

    @retry_on_excel_busy()
    def aggregate(self, group_by_col: str, metric_col: str, agg_func: str = "sum", sheet: Optional[str] = None, top_n: int = 10, ascending: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        rng = ws.UsedRange
        rows_cnt = rng.Rows.Count
        cols_cnt = rng.Columns.Count

        # Find column indices (flexible case/whitespace)
        norm_group = group_by_col.strip().lower()
        norm_metric = metric_col.strip().lower()

        group_idx = None
        metric_idx = None

        for c in range(1, cols_cnt + 1):
            h_val = str(rng.Cells(1, c).Value or "").strip().lower()
            if h_val == norm_group and group_idx is None:
                group_idx = c
            if h_val == norm_metric and metric_idx is None:
                metric_idx = c

        if group_idx is None or metric_idx is None:
            raise ValueError(f"No se encontraron las columnas especificadas: group_by='{group_by_col}', metric='{metric_col}'")

        groups = {}
        for r in range(2, rows_cnt + 1):
            g_val = rng.Cells(r, group_idx).Value
            m_val = rng.Cells(r, metric_idx).Value
            if g_val is None or m_val is None:
                continue

            try:
                num = float(m_val)
            except (ValueError, TypeError):
                continue

            g_key = str(g_val)
            addr_prop = rng.Cells(r, metric_idx).Address
            cell_coord = addr_prop(False, False) if callable(addr_prop) else str(addr_prop).replace("$", "")

            if g_key not in groups:
                groups[g_key] = {
                    "count": 0,
                    "values": [],
                    "cells": [],
                    "min_val": num,
                    "min_cell": cell_coord,
                    "max_val": num,
                    "max_cell": cell_coord
                }

            entry = groups[g_key]
            entry["count"] += 1
            entry["values"].append(num)
            entry["cells"].append(cell_coord)

            if num < entry["min_val"]:
                entry["min_val"] = num
                entry["min_cell"] = cell_coord
            if num > entry["max_val"]:
                entry["max_val"] = num
                entry["max_cell"] = cell_coord

        results = []
        for g_key, data in groups.items():
            vals = data["values"]
            total = sum(vals)
            avg = total / len(vals) if vals else 0.0

            if agg_func.lower() == "sum":
                agg_val = total
            elif agg_func.lower() in ("avg", "mean"):
                agg_val = avg
            elif agg_func.lower() == "min":
                agg_val = data["min_val"]
            elif agg_func.lower() == "max":
                agg_val = data["max_val"]
            elif agg_func.lower() == "count":
                agg_val = float(data["count"])
            else:
                agg_val = total

            results.append({
                "group": g_key,
                "aggregated_value": round(agg_val, 2),
                "total_sum": round(total, 2),
                "average": round(avg, 2),
                "count": data["count"],
                "max_cell": data["max_cell"],
                "max_value": round(data["max_val"], 2),
                "min_cell": data["min_cell"],
                "min_value": round(data["min_val"], 2),
                "sample_citation_cells": data["cells"][:3] + data["cells"][-3:] if len(data["cells"]) > 6 else data["cells"]
            })

        results.sort(key=lambda x: x["aggregated_value"], reverse=not ascending)
        if top_n > 0:
            results = results[:top_n]

        return {
            "sheet": ws.Name,
            "group_by": group_by_col,
            "metric": metric_col,
            "agg_func": agg_func,
            "groups_count": len(results),
            "ranking": results
        }

    @retry_on_excel_busy()
    def add_calculated_column(self, header: str, formula_template: str, number_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT, sheet: Optional[str] = None, autofit: bool = True) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        rng = ws.UsedRange
        last_col = rng.Columns.Count
        new_col = last_col + 1
        last_row = rng.Rows.Count

        # Set header
        ws.Cells(1, new_col).Value = header

        # Set formula for data rows (vectorized in Excel COM, with row substitution if {r}/{row} are used)
        data_rng = ws.Range(ws.Cells(2, new_col), ws.Cells(last_row, new_col))
        if "{r}" in formula_template or "{row}" in formula_template:
            formulas = [[formula_template.replace("{r}", str(r)).replace("{row}", str(r))] for r in range(2, last_row + 1)]
            data_rng.Formula = formulas
        else:
            data_rng.Formula = formula_template

        if number_format:
            data_rng.NumberFormat = number_format

        # Auto-resize ListObject table if sheet contains structured tables
        table_resized = None
        if ws.ListObjects.Count > 0:
            tbl = ws.ListObjects(1)
            tbl.Resize(ws.Range(ws.Cells(1, 1), ws.Cells(last_row, new_col)))
            table_resized = tbl.Name

        if autofit:
            ws.Columns(new_col).AutoFit()

        return {
            "status": "success",
            "sheet": ws.Name,
            "header": header,
            "column_index": new_col,
            "formula_applied": formula_template,
            "rows_affected": last_row - 1,
            "table_resized": table_resized
        }

    @retry_on_excel_busy()
    def apply_style_preset(self, range_str: str, preset: str, sheet: Optional[str] = None) -> Dict[str, Any]:
        """Aplica un preset de estilo universal a un rango en Live Excel COM."""
        ws = self._get_sheet(sheet)
        rng = ws.Range(range_str)
        p = preset.lower()

        if p == "header":
            rng.Interior.Color = hex_to_rgb_int("#0E2E63")
            rng.Font.Color = hex_to_rgb_int("#FFFFFF")
            rng.Font.Bold = True
            rng.VerticalAlignment = -4108  # xlCenter
        elif p == "input_cell":
            rng.Interior.Color = hex_to_rgb_int("#FFF2CC")
            rng.Font.Color = hex_to_rgb_int("#1F4E79")
            rng.HorizontalAlignment = -4108  # xlCenter
            rng.Borders.LineStyle = 1        # xlContinuous
            rng.Borders.Weight = -4138       # xlMedium
            rng.Borders.Color = hex_to_rgb_int("#1F4E79")
        elif p == "total_row":
            rng.Font.Bold = True
            rng.Borders(8).LineStyle = 1     # xlEdgeTop = 8, xlContinuous = 1
            rng.Borders(8).Weight = 2        # xlThin = 2
            rng.Borders(9).LineStyle = -4119 # xlEdgeBottom = 9, xlDouble = -4119
        elif p == "currency":
            FinancialStyleGuide.apply_number_format_resilient(rng, "currency")
        elif p == "percent":
            FinancialStyleGuide.apply_number_format_resilient(rng, "percent")
        elif p == "delta_positive":
            rng.Font.Bold = True
            rng.Font.Color = hex_to_rgb_int("#008000")
        elif p == "delta_negative":
            rng.Font.Bold = True
            rng.Font.Color = hex_to_rgb_int("#C00000")
        else:
            raise ValueError(f"Preset no reconocido: '{preset}'. Presets válidos: header, input_cell, total_row, currency, percent, delta_positive, delta_negative")

        return {
            "status": "success",
            "range": get_com_clean_address(rng),
            "preset": preset,
            "sheet": ws.Name,
            "cells_affected": rng.Count
        }

    @retry_on_excel_busy()
    def create_summary_table(
        self,
        title: str,
        headers: List[str],
        rows_spec: List[Dict[str, Any]],
        dest_sheet: str,
        start_cell: str = "B2",
        parameters: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """
        Crea una tabla resumen / comparativa profesional en Live Excel COM con parámetros vinculados,
        encabezados corporativos, filas calculadas/estáticas, totales formateados y cuadrícula activa.
        """
        ws = self._get_sheet(dest_sheet)
        ws.Activate()
        try:
            self.app.ActiveWindow.DisplayGridlines = True
        except Exception:
            pass

        start_rng = ws.Range(start_cell)
        start_col_idx = start_rng.Column
        start_r_idx = start_rng.Row
        curr_r = start_r_idx

        with excel_fast_mode(self.app):
            # 1. Bloque de Parámetros de Entrada (si se especifican)
            param_cells: Dict[str, str] = {}
            if parameters:
                ws.Cells(curr_r, start_col_idx).Value = "Supuestos y Parámetros"
                ws.Cells(curr_r, start_col_idx).Font.Name = "Segoe UI"
                ws.Cells(curr_r, start_col_idx).Font.Size = 11
                ws.Cells(curr_r, start_col_idx).Font.Bold = True
                ws.Cells(curr_r, start_col_idx).Font.Color = hex_to_rgb_int("#0E2E63")
                curr_r += 1

                for param in parameters:
                    param_name = param.get("name") or param.get("label", "Parámetro")
                    param_val = param.get("value")
                    param_format = param.get("format")
                    param_note = param.get("note")

                    if param.get("cell"):
                        val_coord = param["cell"]
                        val_cell = ws.Range(val_coord)
                        val_cell.Value = param_val
                    else:
                        lbl_cell = ws.Cells(curr_r, start_col_idx)
                        lbl_cell.Value = param_name
                        lbl_cell.Font.Name = "Segoe UI"
                        lbl_cell.Font.Size = 10
                        lbl_cell.Font.Bold = True

                        val_cell = ws.Cells(curr_r, start_col_idx + 1)
                        val_cell.Value = param_val
                        val_coord = get_com_clean_address(val_cell)
                        curr_r += 1

                    self.apply_style_preset(val_coord, "input_cell", sheet=dest_sheet)
                    if param_format:
                        FinancialStyleGuide.apply_number_format_resilient(ws.Range(val_coord), param_format)

                    if param_note:
                        try:
                            ws.Range(val_coord).ClearComments()
                            ws.Range(val_coord).AddComment(param_note)
                        except Exception:
                            pass

                    param_cells[param_name] = val_coord

                curr_r += 1

            # 2. Título general de la tabla
            tbl_title_row = curr_r
            title_cell = ws.Cells(tbl_title_row, start_col_idx)
            title_cell.Value = title
            title_cell.Font.Name = "Segoe UI"
            title_cell.Font.Size = 13
            title_cell.Font.Bold = True
            title_cell.Font.Color = hex_to_rgb_int("#0E2E63")
            curr_r += 1

            # 3. Fila de encabezados
            header_row = curr_r
            end_col_idx = start_col_idx + len(headers) - 1
            for idx, h in enumerate(headers):
                ws.Cells(header_row, start_col_idx + idx).Value = h

            hdr_start_addr = get_com_clean_address(ws.Cells(header_row, start_col_idx))
            hdr_end_addr = get_com_clean_address(ws.Cells(header_row, end_col_idx))
            hdr_range = f"{hdr_start_addr}:{hdr_end_addr}"
            self.apply_style_preset(hdr_range, "header", sheet=dest_sheet)
            curr_r += 1

            # 4. Filas de datos y totales
            row_end_addr = hdr_end_addr
            for row_spec in rows_spec:
                is_total = row_spec.get("is_total", False) or row_spec.get("total", False)
                row_start_addr = get_com_clean_address(ws.Cells(curr_r, start_col_idx))
                row_end_addr = get_com_clean_address(ws.Cells(curr_r, end_col_idx))
                row_range = f"{row_start_addr}:{row_end_addr}"

                label = row_spec.get("label") or row_spec.get("name") or row_spec.get("title")
                row_vals = []
                if "values" in row_spec:
                    vals = row_spec["values"]
                    if len(vals) == len(headers):
                        row_vals = vals
                    elif len(vals) == len(headers) - 1:
                        row_vals = [label] + vals
                    else:
                        row_vals = vals
                elif "formulas" in row_spec:
                    forms = row_spec["formulas"]
                    if len(forms) == len(headers):
                        row_vals = forms
                    elif len(forms) == len(headers) - 1:
                        row_vals = [label] + forms
                    else:
                        row_vals = forms
                else:
                    row_vals = [row_spec.get(h, "") for h in headers]
                    if label and not row_vals[0]:
                        row_vals[0] = label

                formats = row_spec.get("formats")
                row_fmt = row_spec.get("format") or row_spec.get("number_format")
                presets = row_spec.get("presets") or row_spec.get("styles")

                for c_idx, val in enumerate(row_vals):
                    if c_idx >= len(headers):
                        break
                    cell = ws.Cells(curr_r, start_col_idx + c_idx)
                    if isinstance(val, str) and val.startswith("="):
                        cell.Formula = val
                    else:
                        cell.Value = val

                    if not is_total:
                        cell.Font.Name = "Segoe UI"
                        cell.Font.Size = 10
                        cell.Borders.LineStyle = 1  # xlContinuous
                        cell.Borders.Weight = 2     # xlThin
                        cell.Borders.Color = hex_to_rgb_int("#D9D9D9")

                    col_fmt = formats[c_idx] if (formats and c_idx < len(formats)) else row_fmt
                    if col_fmt:
                        FinancialStyleGuide.apply_number_format_resilient(cell, col_fmt)

                    # Aplicar presets directos de columna si se especificaron
                    if presets and c_idx < len(presets) and presets[c_idx]:
                        self.apply_style_preset(get_com_clean_address(cell), presets[c_idx], sheet=dest_sheet)

                if is_total:
                    self.apply_style_preset(row_range, "total_row", sheet=dest_sheet)

                curr_r += 1

            # 5. AutoFit en bloque
            try:
                ws.Range(ws.Cells(start_r_idx, start_col_idx), ws.Cells(curr_r - 1, end_col_idx)).Columns.AutoFit()
            except Exception:
                pass

        return {
            "status": "success",
            "sheet": ws.Name,
            "title": title,
            "table_range": f"{hdr_start_addr}:{row_end_addr}",
            "rows_count": len(rows_spec),
            "headers": headers,
            "parameters": param_cells,
            "start_cell": start_cell
        }

    @retry_on_excel_busy()
    def create_pivot_table(
        self,
        source: Union[str, Dict[str, Any]],
        rows: List[str],
        cols: Optional[List[str]] = None,
        value_field: Optional[str] = None,
        value_func: str = "sum",
        value_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT,
        dest_sheet: Optional[str] = None,
        dest_cell: str = "A3",
        table_name: Optional[str] = None,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Crea PivotCache y PivotTable en Live Excel COM sobre la hoja destino.
        Configura campos de filas, columnas y valor con función de agregación y formato numérico.
        Aplica AutoFit y activa visualización de cuadrícula.
        """
        # 1. Obtener rango origen
        src_ws = self._get_sheet(source_sheet) if source_sheet else self.wb.ActiveSheet
        if isinstance(source, dict) and "range" in source:
            source_range_str = source["range"]
        else:
            source_range_str = str(source)

        if ":" in source_range_str:
            src_range = src_ws.Range(source_range_str)
        elif source_range_str in [tbl.Name for tbl in src_ws.ListObjects]:
            src_range = src_ws.ListObjects(source_range_str).Range
        else:
            # Fallback a UsedRange o ListObject si existe
            if src_ws.ListObjects.Count > 0:
                src_range = src_ws.ListObjects(1).Range
            else:
                src_range = src_ws.UsedRange

        # 2. Obtener o crear hoja destino
        target_sheet_name = dest_sheet or (src_ws.Name + "_Pivot")
        dest_ws = None
        for s in self.wb.Sheets:
            if s.Name.lower() == target_sheet_name.lower():
                dest_ws = s
                break
        if not dest_ws:
            dest_ws = self.wb.Sheets.Add(After=self.wb.Sheets(self.wb.Sheets.Count))
            dest_ws.Name = target_sheet_name

        # Asegurar cuadrícula activa en hoja destino
        try:
            self.app.ActiveWindow.DisplayGridlines = True
        except Exception:
            pass

        # 3. Crear PivotCache
        # SourceType: 1 = xlDatabase
        pc = self.wb.PivotCaches().Create(SourceType=1, SourceData=src_range)

        # 4. Crear PivotTable
        pt_name = table_name or f"PT_{uuid.uuid4().hex[:6].upper() if uuid else 'PivotTable1'}"
        dest_rng = dest_ws.Range(dest_cell)
        pt = pc.CreatePivotTable(TableDestination=dest_rng, TableName=pt_name)

        # 5. Configurar campos de filas
        # xlRowField = 1
        for idx, r_field in enumerate(rows, start=1):
            pf = pt.PivotFields(r_field)
            pf.Orientation = 1
            pf.Position = idx

        # 6. Configurar campos de columnas si se especifican
        # xlColumnField = 2
        if cols:
            for idx, c_field in enumerate(cols, start=1):
                pf = pt.PivotFields(c_field)
                pf.Orientation = 2
                pf.Position = idx

        # 7. Configurar métrica de valores
        # xlSum = -4157, xlCount = -4112, xlAverage = -4106, xlMin = -4139, xlMax = -4136
        func_map = {
            "sum": -4157,
            "count": -4112,
            "average": -4106,
            "avg": -4106,
            "min": -4139,
            "max": -4136
        }
        xl_func = func_map.get(value_func.lower(), -4157)
        metric_name = value_field or "Value"
        caption = f"{value_func.capitalize()} de {metric_name}"

        data_field = pt.AddDataField(pt.PivotFields(metric_name), caption, xl_func)
        if value_format:
            data_field.NumberFormat = value_format

        # 8. AutoFit en columnas de la hoja destino
        try:
            dest_ws.Columns.AutoFit()
        except Exception:
            pass

        # Dimensiones del PivotTable
        pt_range_address = pt.TableRange1.Address if hasattr(pt, "TableRange1") and pt.TableRange1 else dest_cell

        return {
            "status": "success",
            "pivot_table_name": pt.Name,
            "dest_sheet": dest_ws.Name,
            "dest_range": pt_range_address,
            "rows": rows,
            "cols": cols or [],
            "value_field": metric_name,
            "value_func": value_func,
            "value_format": value_format
        }

    @retry_on_excel_busy()
    def create_chart(
        self,
        source: Union[str, Dict[str, Any]],
        chart_type: str = "column_clustered",
        title: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        placement: str = "auto",
        anchor: Optional[str] = None,
        width: int = 480,
        height: int = 300,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Crea ChartObject en la posición calculada (anti-solapamiento) en Live Excel COM.
        Configura tipo de gráfico, título y asocia datos de tabla dinámica o rango fuente.
        """
        # 1. Determinar hoja y rango de datos
        target_ws = self._get_sheet(dest_sheet) if dest_sheet else self.wb.ActiveSheet

        # Si source es nombre de PivotTable o diccionario de create_pivot_table
        source_range = None
        if isinstance(source, dict):
            if "dest_sheet" in source and not dest_sheet:
                target_ws = self._get_sheet(source["dest_sheet"])
            if "dest_range" in source:
                source_range = target_ws.Range(source["dest_range"])
            elif "pivot_table_name" in source:
                pt_obj = target_ws.PivotTables(source["pivot_table_name"])
                source_range = pt_obj.TableRange1
        elif isinstance(source, str):
            # Verificar si coincide con una tabla dinámica existente
            try:
                pt_obj = target_ws.PivotTables(source)
                source_range = pt_obj.TableRange1
            except Exception:
                pass
            if not source_range:
                if ":" in source:
                    src_ws = self._get_sheet(source_sheet) if source_sheet else target_ws
                    source_range = src_ws.Range(source)
                else:
                    try:
                        source_range = target_ws.Range(source)
                    except Exception:
                        source_range = target_ws.UsedRange

        # 2. Calcular posición geométrica anti-solapamiento
        # Obtener coordenadas relativas de la tabla/rango fuente
        start_col = source_range.Column
        start_row = source_range.Row
        num_cols = source_range.Columns.Count
        num_rows = source_range.Rows.Count

        pos = self._calculate_placement_coordinates(
            start_col=start_col,
            start_row=start_row,
            num_cols=num_cols,
            num_rows=num_rows,
            placement=placement,
            anchor=anchor
        )

        target_cell_obj = target_ws.Cells(pos["row"], pos["col"])
        left = float(target_cell_obj.Left)
        top = float(target_cell_obj.Top)

        # 3. Crear ChartObject
        ch_obj = target_ws.ChartObjects().Add(left, top, float(width), float(height))
        chart = ch_obj.Chart

        # 4. Mapear tipo de gráfico COM
        # xlColumnClustered = 51, xlBarClustered = 57, xlLine = 4, xlPie = 5, xlArea = 1
        type_map = {
            "column_clustered": 51,
            "column": 51,
            "bar_clustered": 57,
            "bar": 57,
            "line": 4,
            "pie": 5,
            "area": 1
        }
        xl_chart_type = type_map.get(chart_type.lower(), 51)
        chart.ChartType = xl_chart_type
        chart.SetSourceData(source_range)

        # 5. Título
        if title:
            chart.HasTitle = True
            chart.ChartTitle.Text = title

        return {
            "status": "success",
            "sheet": target_ws.Name,
            "chart_type": chart_type,
            "title": title,
            "placement": pos["placement"],
            "anchor_cell": pos["cell"],
            "bounds": {"left": left, "top": top, "width": width, "height": height}
        }

    @retry_on_excel_busy()
    def save(self, file_path: Optional[str] = None) -> str:
        if file_path:
            abs_p = os.path.abspath(file_path)
            self.wb.SaveAs(abs_p)
            return abs_p
        self.wb.Save()
        return self.wb.FullName

    @retry_on_excel_busy()
    def remove_duplicates(self, range_or_table: str, key_columns: Optional[List[str]] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
        """
        Elimina filas duplicadas dentro del rango o tabla estructurada (ListObject) on-demand.
        En tablas estructuradas de Live COM, redimensiona inmediatamente la tabla (ListObject.Resize)
        sin dejar filas vacías.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if ws.ListObjects.Count > 0:
            if range_or_table:
                for t in ws.ListObjects:
                    if str(t.Name).lower() == str(range_or_table).lower():
                        tbl = t
                        break
            if not tbl and not range_or_table:
                tbl = ws.ListObjects(1)

        if tbl:
            rng = tbl.Range
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count
        else:
            if range_or_table:
                rng = ws.Range(range_or_table)
            else:
                rng = ws.UsedRange
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count

        if num_rows <= 1:
            return {
                "status": "success",
                "range_or_table": range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
                "sheet": ws.Name,
                "original_rows": 0,
                "duplicates_removed": 0,
                "remaining_rows": 0,
                "key_columns": key_columns or "all"
            }

        headers = [str(ws.Cells(start_row, start_col + c - 1).Value or "").strip() for c in range(1, num_cols + 1)]
        
        compare_col_indices = []
        if key_columns:
            headers_lower = [h.lower() for h in headers]
            for k in key_columns:
                k_clean = str(k).strip()
                k_lower = k_clean.lower()
                if k_lower in headers_lower:
                    idx = headers_lower.index(k_lower) + 1
                    compare_col_indices.append(idx)
                elif k_clean.isalpha() and len(k_clean) <= 3:
                    c_idx = column_index_from_string(k_clean) if column_index_from_string else 1
                    rel_idx = c_idx - start_col + 1
                    if 1 <= rel_idx <= num_cols:
                        compare_col_indices.append(rel_idx)
            if not compare_col_indices:
                compare_col_indices = list(range(1, num_cols + 1))
        else:
            compare_col_indices = list(range(1, num_cols + 1))

        data_rows_count = num_rows - 1
        data_range = ws.Range(
            ws.Cells(start_row + 1, start_col),
            ws.Cells(start_row + data_rows_count, start_col + num_cols - 1)
        )
        data_values = data_range.Value

        if data_rows_count == 1:
            data_values = [data_values]
        elif not isinstance(data_values, (list, tuple)):
            data_values = [[data_values]]

        seen = set()
        dup_row_indices = []
        for i, row in enumerate(data_values):
            sheet_r = start_row + 1 + i
            row_items = list(row) if isinstance(row, (list, tuple)) else [row]
            key = tuple(row_items[ci - 1] if ci - 1 < len(row_items) else None for ci in compare_col_indices)
            if key in seen:
                dup_row_indices.append(sheet_r)
            else:
                seen.add(key)

        duplicates_count = len(dup_row_indices)
        if duplicates_count > 0:
            with excel_fast_mode(self.app):
                for r in sorted(dup_row_indices, reverse=True):
                    if tbl:
                        tbl_row_idx = r - start_row
                        try:
                            tbl.ListRows(tbl_row_idx).Delete()
                        except Exception:
                            ws.Range(ws.Cells(r, start_col), ws.Cells(r, start_col + num_cols - 1)).Delete(-4162)
                    else:
                        ws.Range(ws.Cells(r, start_col), ws.Cells(r, start_col + num_cols - 1)).Delete(-4162)

                if tbl:
                    new_end_row = start_row + data_rows_count - duplicates_count
                    try:
                        new_tbl_rng = ws.Range(ws.Cells(start_row, start_col), ws.Cells(new_end_row, start_col + num_cols - 1))
                        tbl.Resize(new_tbl_rng)
                    except Exception:
                        pass

        remaining_data_rows = data_rows_count - duplicates_count
        return {
            "status": "success",
            "range_or_table": range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
            "sheet": ws.Name,
            "original_rows": data_rows_count,
            "duplicates_removed": duplicates_count,
            "remaining_rows": remaining_data_rows,
            "key_columns": key_columns or "all"
        }

    @retry_on_excel_busy()
    def detect_and_flag_outliers(
        self,
        column: str,
        method: str = "iqr",
        range_or_table: Optional[str] = None,
        sheet: Optional[str] = None,
        flag_style: str = "warning",
        add_comment: bool = True
    ) -> Dict[str, Any]:
        """
        Detecta y señala visualmente valores atípicos (outliers) mediante IQR o Z-score on-demand.
        Aplica formato no destructivo (fondo suave #FCE4D6 y texto rojo oscuro bold) e inyecta comentarios.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if ws.ListObjects.Count > 0:
            if range_or_table:
                for t in ws.ListObjects:
                    if str(t.Name).lower() == str(range_or_table).lower():
                        tbl = t
                        break
            if not tbl and not range_or_table:
                tbl = ws.ListObjects(1)

        if tbl:
            rng = tbl.Range
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count
        else:
            if range_or_table:
                rng = ws.Range(range_or_table)
            else:
                rng = ws.UsedRange
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count

        headers = [str(ws.Cells(start_row, start_col + c - 1).Value or "").strip() for c in range(1, num_cols + 1)]
        headers_lower = [h.lower() for h in headers]
        col_clean = str(column).strip()
        col_lower = col_clean.lower()

        target_col_rel = None
        if col_lower in headers_lower:
            target_col_rel = headers_lower.index(col_lower) + 1
        elif col_clean.isalpha() and len(col_clean) <= 3:
            c_idx = column_index_from_string(col_clean) if column_index_from_string else 1
            rel_idx = c_idx - start_col + 1
            if 1 <= rel_idx <= num_cols:
                target_col_rel = rel_idx

        if not target_col_rel:
            raise ValueError(f"Columna '{column}' no encontrada en el rango/tabla. Encabezados disponibles: {headers}")

        abs_col_idx = start_col + target_col_rel - 1

        val_list = []
        row_val_map = {}
        for r in range(start_row + 1, start_row + num_rows):
            v = ws.Cells(r, abs_col_idx).Value
            if v is not None and str(v).strip() != "":
                try:
                    fv = float(str(v).replace("$", "").replace(" ", "").replace(",", ""))
                    val_list.append(fv)
                    row_val_map[r] = fv
                except Exception:
                    pass

        if len(val_list) < 4:
            return {
                "status": "success",
                "sheet": ws.Name,
                "column": column,
                "method": method,
                "total_values_evaluated": len(val_list),
                "outliers_count": 0,
                "thresholds": {},
                "flagged_cells": [],
                "outliers": []
            }

        method_clean = method.lower()
        if method_clean == "zscore":
            mean = statistics.mean(val_list)
            std = statistics.stdev(val_list) if len(val_list) > 1 else 0
            z_thresh = 2.5 if len(val_list) < 30 else 3.0
            lower_bound = mean - z_thresh * std
            upper_bound = mean + z_thresh * std
            thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "mean": mean, "std": std, "z_threshold": z_thresh}
        else:
            sorted_vals = sorted(val_list)
            q1, _, q3 = statistics.quantiles(sorted_vals, n=4, method="inclusive")
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "q1": q1, "q3": q3, "iqr": iqr}

        outliers = []
        col_letter = get_column_letter(abs_col_idx) if get_column_letter else f"C{abs_col_idx}"
        flagged_coords = []

        for r, fv in row_val_map.items():
            if fv < lower_bound or fv > upper_bound:
                cell_coord = f"{col_letter}{r}"
                flagged_coords.append(cell_coord)
                outliers.append({
                    "cell": cell_coord,
                    "row": r,
                    "value": fv,
                    "reason": f"Valor {fv} fuera de límites [{lower_bound:.2f}, {upper_bound:.2f}]"
                })

        with excel_fast_mode(self.app):
            if flagged_coords:
                for chunk in _chunk_cell_addresses(flagged_coords):
                    try:
                        chunk_rng = ws.Range(chunk)
                        chunk_rng.Interior.Color = hex_to_rgb_int("#FCE4D6")
                        chunk_rng.Font.Color = hex_to_rgb_int("#C00000")
                        chunk_rng.Font.Bold = True
                    except Exception:
                        for coord in chunk.split(","):
                            try:
                                c_rng = ws.Range(coord)
                                c_rng.Interior.Color = hex_to_rgb_int("#FCE4D6")
                                c_rng.Font.Color = hex_to_rgb_int("#C00000")
                                c_rng.Font.Bold = True
                            except Exception:
                                pass

            if add_comment:
                for o in outliers:
                    try:
                        cell = ws.Cells(o["row"], abs_col_idx)
                        try:
                            cell.ClearComments()
                        except Exception:
                            pass
                        cell.AddComment(f"Outlier detectado ({method_clean.upper()}): valor={o['value']}, límites=[{lower_bound:.2f}, {upper_bound:.2f}]")
                    except Exception:
                        pass


        return {
            "status": "success",
            "sheet": ws.Name,
            "column": headers[target_col_rel - 1],
            "method": method_clean,
            "total_values_evaluated": len(val_list),
            "outliers_count": len(outliers),
            "thresholds": thresholds,
            "flagged_cells": [o["cell"] for o in outliers],
            "outliers": outliers
        }

    @retry_on_excel_busy()
    def clean_table_dataset(self, target_range_or_table: str, rules: Dict[str, Any], sheet: Optional[str] = None) -> Dict[str, Any]:
        """
        Limpia y estandariza en bloque los datos de una tabla o rango on-demand.
        Procesa en una sola pasada en memoria:
          - standardize_text: trim, casing y mapeo de sinónimos/reemplazos.
          - cast_types: conversión segura de tipos (float, int, str).
          - coerce_dates: conversión a seriales de Excel o datetimes con formato uniforme.
        Vuelca los datos limpios en una sola asignación matricial a través de excel_fast_mode.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if ws.ListObjects.Count > 0:
            if target_range_or_table:
                for t in ws.ListObjects:
                    if str(t.Name).lower() == str(target_range_or_table).lower():
                        tbl = t
                        break
            if not tbl and not target_range_or_table:
                tbl = ws.ListObjects(1)

        if tbl:
            rng = tbl.Range
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count
        else:
            if target_range_or_table:
                rng = ws.Range(target_range_or_table)
            else:
                rng = ws.UsedRange
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count

        if num_rows <= 1:
            return {
                "status": "success",
                "sheet": ws.Name,
                "target_range_or_table": target_range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
                "rows_processed": 0,
                "cells_modified": 0,
                "rules_applied": list(rules.keys())
            }

        headers = [str(ws.Cells(start_row, start_col + c - 1).Value or "").strip() for c in range(1, num_cols + 1)]

        std_text_rules = rules.get("standardize_text", {})
        cast_type_rules = rules.get("cast_types", {})
        coerce_date_rules = rules.get("coerce_dates", {})

        col_text_cfg = {}
        col_cast_cfg = {}
        col_date_cfg = {}

        for c_idx, h in enumerate(headers, start=1):
            h_lower = h.lower()
            for rule_col, cfg in std_text_rules.items():
                if rule_col.strip().lower() == h_lower:
                    col_text_cfg[c_idx] = cfg
                    break

            for rule_col, target_t in cast_type_rules.items():
                if rule_col.strip().lower() == h_lower:
                    col_cast_cfg[c_idx] = str(target_t).lower()
                    break

            if isinstance(coerce_date_rules, list):
                if any(dc.strip().lower() == h_lower for dc in coerce_date_rules):
                    col_date_cfg[c_idx] = {"to_serial": True, "format": "yyyy-mm-dd"}
            elif isinstance(coerce_date_rules, dict):
                for dc, cfg in coerce_date_rules.items():
                    if dc.strip().lower() == h_lower:
                        col_date_cfg[c_idx] = cfg if isinstance(cfg, dict) else {"to_serial": True, "format": "yyyy-mm-dd"}
                        break

        data_rows_count = num_rows - 1
        data_range = ws.Range(
            ws.Cells(start_row + 1, start_col),
            ws.Cells(start_row + data_rows_count, start_col + num_cols - 1)
        )
        data_values = data_range.Value
        if data_rows_count == 1:
            data_values = [data_values]
        elif not isinstance(data_values, (list, tuple)):
            data_values = [[data_values]]

        modified_data = []
        cells_modified = 0

        for r_idx, row in enumerate(data_values):
            row_items = list(row) if isinstance(row, (list, tuple)) else [row]
            new_row = []
            for c_idx in range(1, num_cols + 1):
                val = row_items[c_idx - 1] if c_idx - 1 < len(row_items) else None
                orig_val = val

                # 1. Date coercion
                if c_idx in col_date_cfg:
                    serial = parse_excel_date_serial(val)
                    if serial is not None:
                        val = serial

                # 2. Text standardization
                if c_idx in col_text_cfg and val is not None:
                    t_cfg = col_text_cfg[c_idx]
                    s = str(val)
                    if t_cfg.get("trim", True):
                        s = s.strip()
                    reps = t_cfg.get("replacements") or t_cfg.get("mapping")
                    if reps:
                        if s in reps:
                            s = reps[s]
                        elif s.lower() in {k.lower(): v for k, v in reps.items()}:
                            for k, v in reps.items():
                                if s.lower() == k.lower():
                                    s = v
                                    break
                    casing = t_cfg.get("casing") or t_cfg.get("case")
                    if casing == "upper":
                        s = s.upper()
                    elif casing == "lower":
                        s = s.lower()
                    elif casing == "title":
                        s = s.title()
                    elif casing == "capitalize":
                        s = s.capitalize()
                    val = s

                # 3. Type casting
                if c_idx in col_cast_cfg and val is not None and str(val).strip() != "":
                    tgt = col_cast_cfg[c_idx]
                    try:
                        clean_str = str(val).replace("$", "").replace(" ", "").replace(",", "")
                        if tgt in ("float", "double", "decimal"):
                            val = float(clean_str)
                        elif tgt in ("int", "integer"):
                            val = int(round(float(clean_str)))
                        elif tgt in ("str", "string", "text"):
                            val = str(val)
                    except Exception:
                        pass

                if val != orig_val:
                    cells_modified += 1
                new_row.append(val)
            modified_data.append(tuple(new_row))

        with excel_fast_mode(self.app):
            if rules.get("trim_headers", True):
                for c_idx in range(1, num_cols + 1):
                    h_val = ws.Cells(start_row, start_col + c_idx - 1).Value
                    if isinstance(h_val, str) and h_val != h_val.strip():
                        ws.Cells(start_row, start_col + c_idx - 1).Value = h_val.strip()
                        cells_modified += 1

            data_range.Value = tuple(modified_data)

            for c_idx, d_cfg in col_date_cfg.items():
                fmt = d_cfg.get("format", "yyyy-mm-dd") if isinstance(d_cfg, dict) else "yyyy-mm-dd"
                col_rng = ws.Range(
                    ws.Cells(start_row + 1, start_col + c_idx - 1),
                    ws.Cells(start_row + data_rows_count, start_col + c_idx - 1)
                )
                try:
                    col_rng.NumberFormat = fmt
                except Exception:
                    pass

        return {
            "status": "success",
            "sheet": ws.Name,
            "target_range_or_table": target_range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
            "rows_processed": data_rows_count,
            "cells_modified": cells_modified,
            "rules_applied": list(rules.keys())
        }

    @retry_on_excel_busy()
    def curate_pipeline(
        self,
        target_range_or_table: str,
        pipeline_spec: Dict[str, Any],
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Pipeline unificado de curaduría y vectorización multi-rango de alto rendimiento.
        Ejecuta deduplicación, limpieza/estandarización de datos, casteo de tipos,
        coerción de fechas a números seriales y detección/marcado de outliers en UNA SOLA PASADA
        en memoria Python y UN SOLO VIAJE COM de lectura/escritura bajo excel_fast_mode.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if ws.ListObjects.Count > 0:
            if target_range_or_table:
                for t in ws.ListObjects:
                    if str(t.Name).lower() == str(target_range_or_table).lower():
                        tbl = t
                        break
            if not tbl and not target_range_or_table:
                tbl = ws.ListObjects(1)

        if tbl:
            rng = tbl.Range
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count
        else:
            if target_range_or_table:
                rng = ws.Range(target_range_or_table)
            else:
                rng = ws.UsedRange
            start_row = rng.Row
            start_col = rng.Column
            num_rows = rng.Rows.Count
            num_cols = rng.Columns.Count

        if num_rows <= 1:
            return {
                "status": "success",
                "sheet": ws.Name,
                "target_range_or_table": target_range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
                "original_rows": 0,
                "remaining_rows": 0,
                "duplicates_removed": 0,
                "cells_cleaned": 0,
                "rules_applied": [],
                "outliers_count": 0,
                "outliers_flagged": [],
                "outliers_summary": []
            }

        # 1. Lectura única a memoria bajo excel_fast_mode
        with excel_fast_mode(self.app):
            raw_val = rng.Value

        if not isinstance(raw_val, (list, tuple)):
            matrix = [[raw_val]]
        elif len(raw_val) > 0 and not isinstance(raw_val[0], (list, tuple)):
            matrix = [list(raw_val)]
        else:
            matrix = [list(r) for r in raw_val]

        header_row = matrix[0] if matrix else []
        headers = [str(header_row[c] or "").strip() for c in range(min(num_cols, len(header_row)))]
        while len(headers) < num_cols:
            headers.append("")

        data_rows = matrix[1:] if len(matrix) > 1 else []
        for r in data_rows:
            while len(r) < num_cols:
                r.append(None)
        orig_data_rows_count = len(data_rows)

        # 2. Deduplicación en memoria
        dedup_spec = pipeline_spec.get("dedup")
        duplicates_count = 0
        if dedup_spec:
            if isinstance(dedup_spec, dict):
                key_cols = dedup_spec.get("key_columns")
            elif isinstance(dedup_spec, (list, tuple)):
                key_cols = list(dedup_spec)
            else:
                key_cols = None

            headers_lower = [h.lower() for h in headers]
            compare_cols = []
            if key_cols:
                for k in key_cols:
                    k_clean = str(k).strip()
                    k_lower = k_clean.lower()
                    if k_lower in headers_lower:
                        compare_cols.append(headers_lower.index(k_lower) + 1)
                    elif k_clean.isalpha() and len(k_clean) <= 3:
                        c_idx = column_index_from_string(k_clean) if column_index_from_string else 1
                        rel_idx = c_idx - start_col + 1
                        if 1 <= rel_idx <= num_cols:
                            compare_cols.append(rel_idx)
                if not compare_cols:
                    compare_cols = list(range(1, num_cols + 1))
            else:
                compare_cols = list(range(1, num_cols + 1))

            seen = set()
            unique_rows = []
            for r in data_rows:
                key = tuple(
                    str(r[ci - 1]).strip() if isinstance(r[ci - 1], str) else r[ci - 1]
                    if ci - 1 < len(r) else None
                    for ci in compare_cols
                )
                if key not in seen:
                    seen.add(key)
                    unique_rows.append(r)
            duplicates_count = len(data_rows) - len(unique_rows)
            data_rows = unique_rows

        # 3. Limpieza y estandarización en memoria
        clean_spec = pipeline_spec.get("clean")
        cells_modified = 0
        col_date_cfg = {}
        if clean_spec:
            std_text_rules = clean_spec.get("standardize_text", {})
            cast_type_rules = clean_spec.get("cast_types", {})
            coerce_date_rules = clean_spec.get("coerce_dates", {})
            trim_headers = clean_spec.get("trim_headers", True)

            col_text_cfg = {}
            col_cast_cfg = {}

            for c_idx, h in enumerate(headers, start=1):
                h_lower = h.lower()
                for rule_col, cfg in std_text_rules.items():
                    if rule_col.strip().lower() == h_lower:
                        col_text_cfg[c_idx] = cfg
                        break
                for rule_col, tgt_t in cast_type_rules.items():
                    if rule_col.strip().lower() == h_lower:
                        col_cast_cfg[c_idx] = str(tgt_t).lower()
                        break
                if isinstance(coerce_date_rules, (list, tuple)):
                    if any(dc.strip().lower() == h_lower for dc in coerce_date_rules):
                        col_date_cfg[c_idx] = {"to_serial": True, "format": "yyyy-mm-dd"}
                elif isinstance(coerce_date_rules, dict):
                    for dc, cfg in coerce_date_rules.items():
                        if dc.strip().lower() == h_lower:
                            col_date_cfg[c_idx] = cfg if isinstance(cfg, dict) else {"to_serial": True, "format": "yyyy-mm-dd"}
                            break

            if trim_headers:
                new_headers = []
                for h in headers:
                    th = h.strip() if isinstance(h, str) else h
                    if th != h:
                        cells_modified += 1
                    new_headers.append(th)
                headers = new_headers

            cleaned_data_rows = []
            for r in data_rows:
                new_row = []
                for c_idx in range(1, num_cols + 1):
                    val = r[c_idx - 1] if c_idx - 1 < len(r) else None
                    orig_val = val

                    # 1. Coerción de fechas a números seriales
                    if c_idx in col_date_cfg:
                        serial = parse_excel_date_serial(val)
                        if serial is not None:
                            val = serial

                    # 2. Estandarización de texto
                    if c_idx in col_text_cfg and val is not None:
                        t_cfg = col_text_cfg[c_idx]
                        s = str(val)
                        if t_cfg.get("trim", True):
                            s = s.strip()
                        reps = t_cfg.get("replacements") or t_cfg.get("mapping")
                        if reps:
                            if s in reps:
                                s = reps[s]
                            elif s.lower() in {k.lower(): v for k, v in reps.items()}:
                                for k, v in reps.items():
                                    if s.lower() == k.lower():
                                        s = v
                                        break
                        casing = t_cfg.get("casing") or t_cfg.get("case")
                        if casing == "upper":
                            s = s.upper()
                        elif casing == "lower":
                            s = s.lower()
                        elif casing == "title":
                            s = s.title()
                        elif casing == "capitalize":
                            s = s.capitalize()
                        val = s

                    # 3. Casteo seguro de tipos
                    if c_idx in col_cast_cfg and val is not None and str(val).strip() != "":
                        tgt = col_cast_cfg[c_idx]
                        try:
                            clean_str = str(val).replace("$", "").replace(" ", "").replace(",", "")
                            if tgt in ("float", "double", "decimal"):
                                val = float(clean_str)
                            elif tgt in ("int", "integer"):
                                val = int(round(float(clean_str)))
                            elif tgt in ("str", "string", "text"):
                                val = str(val)
                        except Exception:
                            pass

                    if val != orig_val:
                        cells_modified += 1
                    new_row.append(val)
                cleaned_data_rows.append(new_row)
            data_rows = cleaned_data_rows

        # 4. Cálculo estadístico de outliers en memoria
        outliers_spec = pipeline_spec.get("outliers")
        outliers_results = []
        all_flagged_cells = []
        outlier_cells_to_format = []

        if outliers_spec:
            if isinstance(outliers_spec, dict):
                if "column" in outliers_spec or "columns" in outliers_spec:
                    spec_list = [outliers_spec]
                else:
                    spec_list = [{"column": k, **(v if isinstance(v, dict) else {})} for k, v in outliers_spec.items()]
            elif isinstance(outliers_spec, (list, tuple)):
                spec_list = []
                for it in outliers_spec:
                    if isinstance(it, str):
                        spec_list.append({"column": it})
                    elif isinstance(it, dict):
                        spec_list.append(it)
            else:
                spec_list = []

            headers_lower = [h.lower() for h in headers]
            for s_item in spec_list:
                target_col_name = str(s_item.get("column") or "").strip()
                method = str(s_item.get("method") or "iqr").lower()
                add_comment = bool(s_item.get("add_comment", True))

                target_col_rel = None
                if target_col_name.lower() in headers_lower:
                    target_col_rel = headers_lower.index(target_col_name.lower()) + 1
                elif target_col_name.isalpha() and len(target_col_name) <= 3:
                    c_idx = column_index_from_string(target_col_name) if column_index_from_string else 1
                    rel_idx = c_idx - start_col + 1
                    if 1 <= rel_idx <= num_cols:
                        target_col_rel = rel_idx

                if not target_col_rel:
                    continue

                abs_col_idx = start_col + target_col_rel - 1
                col_letter = get_column_letter(abs_col_idx) if get_column_letter else f"C{abs_col_idx}"

                val_list = []
                row_val_map = {}
                for r_offset, r_data in enumerate(data_rows):
                    sheet_r = start_row + 1 + r_offset
                    v = r_data[target_col_rel - 1] if target_col_rel - 1 < len(r_data) else None
                    if v is not None and str(v).strip() != "":
                        try:
                            fv = float(str(v).replace("$", "").replace(" ", "").replace(",", ""))
                            val_list.append(fv)
                            row_val_map[sheet_r] = fv
                        except Exception:
                            pass

                if len(val_list) < 4:
                    outliers_results.append({
                        "column": headers[target_col_rel - 1],
                        "method": method,
                        "total_values_evaluated": len(val_list),
                        "outliers_count": 0,
                        "thresholds": {},
                        "flagged_cells": [],
                        "outliers": []
                    })
                    continue

                if method == "zscore":
                    mean = statistics.mean(val_list)
                    std = statistics.stdev(val_list) if len(val_list) > 1 else 0
                    z_thresh = 2.5 if len(val_list) < 30 else 3.0
                    lower_bound = mean - z_thresh * std
                    upper_bound = mean + z_thresh * std
                    thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "mean": mean, "std": std, "z_threshold": z_thresh}
                else:
                    sorted_vals = sorted(val_list)
                    q1, _, q3 = statistics.quantiles(sorted_vals, n=4, method="inclusive")
                    iqr = q3 - q1
                    lower_bound = q1 - 1.5 * iqr
                    upper_bound = q3 + 1.5 * iqr
                    thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "q1": q1, "q3": q3, "iqr": iqr}

                col_outliers = []
                for sheet_r, fv in row_val_map.items():
                    if fv < lower_bound or fv > upper_bound:
                        coord = f"{col_letter}{sheet_r}"
                        reason = f"Valor {fv} fuera de límites [{lower_bound:.2f}, {upper_bound:.2f}]"
                        col_outliers.append({
                            "cell": coord,
                            "row": sheet_r,
                            "value": fv,
                            "reason": reason
                        })
                        all_flagged_cells.append(coord)
                        outlier_cells_to_format.append((coord, sheet_r, abs_col_idx, fv, reason, add_comment, method))

                outliers_results.append({
                    "column": headers[target_col_rel - 1],
                    "method": method,
                    "total_values_evaluated": len(val_list),
                    "outliers_count": len(col_outliers),
                    "thresholds": thresholds,
                    "flagged_cells": [o["cell"] for o in col_outliers],
                    "outliers": col_outliers
                })

        # 5. Volcado único a Excel COM bajo excel_fast_mode
        with excel_fast_mode(self.app):
            full_matrix = [headers] + data_rows
            new_total_rows = len(full_matrix)
            new_data_rows_count = len(data_rows)

            write_rng = ws.Range(
                ws.Cells(start_row, start_col),
                ws.Cells(start_row + new_total_rows - 1, start_col + num_cols - 1)
            )
            write_rng.Value = tuple(tuple(r) for r in full_matrix)

            # Redimensionamiento de tabla y eliminación de filas sobrantes si hubo duplicados
            if duplicates_count > 0:
                leftover_start = start_row + new_total_rows
                leftover_end = start_row + num_rows - 1
                if tbl:
                    new_tbl_rng = ws.Range(
                        ws.Cells(start_row, start_col),
                        ws.Cells(start_row + new_total_rows - 1, start_col + num_cols - 1)
                    )
                    try:
                        tbl.Resize(new_tbl_rng)
                    except Exception:
                        pass
                if leftover_start <= leftover_end:
                    try:
                        ws.Range(
                            ws.Cells(leftover_start, start_col),
                            ws.Cells(leftover_end, start_col + num_cols - 1)
                        ).Delete(-4162)
                    except Exception:
                        try:
                            ws.Range(
                                ws.Cells(leftover_start, start_col),
                                ws.Cells(leftover_end, start_col + num_cols - 1)
                            ).Clear()
                        except Exception:
                            pass

            # Formato de fechas en columnas correspondientes
            if col_date_cfg and new_data_rows_count > 0:
                for c_idx, d_cfg in col_date_cfg.items():
                    fmt = d_cfg.get("format", "yyyy-mm-dd") if isinstance(d_cfg, dict) else "yyyy-mm-dd"
                    c_rng = ws.Range(
                        ws.Cells(start_row + 1, start_col + c_idx - 1),
                        ws.Cells(start_row + new_data_rows_count, start_col + c_idx - 1)
                    )
                    try:
                        c_rng.NumberFormat = fmt
                    except Exception:
                        pass

            # Formateo agrupado multi-rango de outliers en UNA SOLA llamada COM por bloque
            if outlier_cells_to_format:
                flagged_coords = [item[0] for item in outlier_cells_to_format]
                for chunk in _chunk_cell_addresses(flagged_coords):
                    try:
                        chunk_rng = ws.Range(chunk)
                        chunk_rng.Interior.Color = hex_to_rgb_int("#FCE4D6")
                        chunk_rng.Font.Color = hex_to_rgb_int("#C00000")
                        chunk_rng.Font.Bold = True
                    except Exception:
                        for coord in chunk.split(","):
                            try:
                                c_rng = ws.Range(coord)
                                c_rng.Interior.Color = hex_to_rgb_int("#FCE4D6")
                                c_rng.Font.Color = hex_to_rgb_int("#C00000")
                                c_rng.Font.Bold = True
                            except Exception:
                                pass

                for coord, sheet_r, abs_col_idx, fv, reason, add_com, m_name in outlier_cells_to_format:
                    if add_com:
                        try:
                            c_obj = ws.Cells(sheet_r, abs_col_idx)
                            try:
                                c_obj.ClearComments()
                            except Exception:
                                pass
                            c_obj.AddComment(f"Outlier detectado ({m_name.upper()}): {reason}")
                        except Exception:
                            pass

        return {
            "status": "success",
            "sheet": ws.Name,
            "target_range_or_table": target_range_or_table or (tbl.Name if tbl else get_com_clean_address(rng)),
            "original_rows": orig_data_rows_count,
            "remaining_rows": len(data_rows),
            "duplicates_removed": duplicates_count,
            "cells_cleaned": cells_modified,
            "rules_applied": list(clean_spec.keys()) if clean_spec else [],
            "outliers_count": len(all_flagged_cells),
            "outliers_flagged": all_flagged_cells,
            "outliers_summary": outliers_results
        }

    def close(self):
        pass



class HeadlessOpenPyXLBackend(BaseExcelBackend):
    def __init__(self, file_path: Optional[str] = None):
        op = _get_openpyxl()
        if not op:
            raise ImportError("openpyxl no está instalado en este entorno.")
        # Carga perezosa de estilos y comentarios en el namespace global
        _get_styles()
        _get_comments()
        self.file_path = file_path
        if file_path and os.path.exists(file_path):
            self.wb = self._load_workbook_resilient(file_path)
        else:
            self.wb = op.Workbook()

    def _load_workbook_resilient(self, file_path: str):
        """Carga openpyxl soportando lectura concurrente si el archivo está bloqueado por Excel."""
        op = _get_openpyxl()
        try:
            return op.load_workbook(file_path, data_only=False)
        except PermissionError:
            wfile = _get_win32file()
            wcon = _get_win32con()
            if wfile and wcon:
                handle = wfile.CreateFile(
                    os.path.abspath(file_path),
                    wfile.GENERIC_READ,
                    wfile.FILE_SHARE_READ | wfile.FILE_SHARE_WRITE | wfile.FILE_SHARE_DELETE,
                    None,
                    wfile.OPEN_EXISTING,
                    wfile.FILE_ATTRIBUTE_NORMAL,
                    None
                )
                chunks = []
                while True:
                    hr, data = wfile.ReadFile(handle, 64 * 1024)
                    if not data:
                        break
                    chunks.append(data)
                wfile.CloseHandle(handle)
                buffer = io.BytesIO(b"".join(chunks))
                return op.load_workbook(buffer, data_only=False)
            raise

    def _get_sheet(self, sheet_name: Optional[str]):
        if sheet_name:
            if sheet_name in self.wb.sheetnames:
                return self.wb[sheet_name]
            return self.wb.create_sheet(sheet_name)
        return self.wb.active

    def status(self) -> Dict[str, Any]:
        return {
            "backend": "headless_openpyxl",
            "file": self.file_path,
            "active_sheet": self.wb.active.title if self.wb.active else None,
            "sheets": self.wb.sheetnames
        }

    def get_cell_ranges(self, ranges: List[str], sheet: Optional[str] = None, include_formulas: bool = False, include_styles: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        res = {}
        for r in ranges:
            target_ws = ws
            clean_range = r
            if "!" in r:
                s_name, clean_range = r.split("!", 1)
                target_ws = self._get_sheet(s_name)

            min_col, min_row, max_col, max_row = range_boundaries(clean_range)
            vals = []
            forms = []
            for row_idx in range(min_row, max_row + 1):
                r_vals = []
                r_forms = []
                for col_idx in range(min_col, max_col + 1):
                    cell = target_ws.cell(row=row_idx, column=col_idx)
                    r_vals.append(cell.value)
                    r_forms.append(cell.value if (isinstance(cell.value, str) and cell.value.startswith("=")) else "")
                vals.append(r_vals)
                forms.append(r_forms)

            res[r] = {"values": vals}
            if include_formulas:
                res[r]["formulas"] = forms
        return res

    def set_cell_range(self, start_cell: str, values: List[List[Any]], sheet: Optional[str] = None, formulas: Optional[List[List[str]]] = None, styles: Optional[Dict[str, Any]] = None, autofit: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        start_col, start_row = coordinate_from_string(start_cell)
        start_col_idx = column_index_from_string(start_col)

        data = formulas if formulas else values
        if not data:
            return {"status": "error", "message": "No data provided"}

        for r_idx, row in enumerate(data):
            for c_idx, val in enumerate(row):
                cell = ws.cell(row=start_row + r_idx, column=start_col_idx + c_idx)
                if isinstance(val, str) and val.startswith("="):
                    cell.value = sanitize_formula_compatibility(val)
                else:
                    cell.value = val

        if autofit:
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

        return {"status": "success", "rows": len(data), "cols": len(data[0]) if data else 0}

    def set_cells(self, cells_map: Dict[str, Dict[str, Any]], sheet: Optional[str] = None, autofit: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        for coord, cell_info in cells_map.items():
            cell = ws[coord]
            if "formula" in cell_info and cell_info["formula"]:
                cell.value = sanitize_formula_compatibility(cell_info["formula"])
            elif "value" in cell_info:
                val = cell_info["value"]
                if isinstance(val, str) and val.startswith("="):
                    cell.value = sanitize_formula_compatibility(val)
                else:
                    cell.value = val

            if "note" in cell_info and cell_info["note"]:
                cell.comment = Comment(cell_info["note"], "Antigravity")

            styles = cell_info.get("cellStyles", {})
            if styles:
                font_args = {}
                if "fontWeight" in styles:
                    font_args["bold"] = (styles["fontWeight"] == "bold")
                if "fontStyle" in styles:
                    font_args["italic"] = (styles["fontStyle"] == "italic")
                if "fontSize" in styles:
                    font_args["size"] = int(styles["fontSize"])
                if "fontColor" in styles:
                    font_args["color"] = styles["fontColor"].lstrip("#")
                if font_args:
                    cell.font = Font(**font_args)

                if "backgroundColor" in styles:
                    hex_color = styles["backgroundColor"].lstrip("#")
                    cell.fill = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")

                if "numberFormat" in styles:
                    cell.number_format = styles["numberFormat"]

                if "horizontalAlignment" in styles:
                    cell.alignment = Alignment(horizontal=styles["horizontalAlignment"].lower())

        if autofit:
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                col_letter = get_column_letter(col[0].column)
                ws.column_dimensions[col_letter].width = max(max_len + 3, 10)

        return {"status": "success", "cells_updated": len(cells_map)}

    def get_range_as_csv(self, range_str: str, sheet: Optional[str] = None, delimiter: str = ";") -> str:
        data = self.get_cell_ranges([range_str], sheet=sheet)[range_str]["values"]
        output = io.StringIO()
        writer = csv.writer(output, delimiter=delimiter, lineterminator="\n")
        for row in data:
            writer.writerow(["" if x is None else str(x) for x in row])
        return output.getvalue().strip()

    def copy_paste_range(self, src_range: str, dst_start_cell: str, src_sheet: Optional[str] = None, dst_sheet: Optional[str] = None, paste_type: str = "all") -> Dict[str, Any]:
        ws_src = self._get_sheet(src_sheet)
        ws_dst = self._get_sheet(dst_sheet)

        min_col, min_row, max_col, max_row = range_boundaries(src_range)
        dst_col_letter, dst_row = coordinate_from_string(dst_start_cell)
        dst_col = column_index_from_string(dst_col_letter)

        for r_offset, r in enumerate(range(min_row, max_row + 1)):
            for c_offset, c in enumerate(range(min_col, max_col + 1)):
                src_cell = ws_src.cell(row=r, column=c)
                target_cell = ws_dst.cell(row=dst_row + r_offset, column=dst_col + c_offset)
                if paste_type in ("all", "values", "formulas"):
                    target_cell.value = src_cell.value
                if paste_type in ("all", "formats") and src_cell.has_style:
                    target_cell._style = src_cell._style

        return {"status": "success", "source": src_range, "destination": dst_start_cell}

    def check_formula_errors(self, range_str: str, sheet: Optional[str] = None) -> List[Dict[str, Any]]:
        ws = self._get_sheet(sheet)
        min_col, min_row, max_col, max_row = range_boundaries(range_str)
        errors = []
        known_errors = {"#VALUE!", "#REF!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!"}
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if isinstance(cell.value, str) and cell.value in known_errors:
                    errors.append({
                        "cell": cell.coordinate,
                        "error": cell.value,
                        "formula": cell.value
                    })
        return errors

    def audit_data_quality(self, range_str: Optional[str] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        if range_str:
            min_col, min_row, max_col, max_row = range_boundaries(range_str)
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        header_whitespace = []
        numbers_as_text = []
        inconsistent_types = []
        total_formulas = 0
        total_cells = (max_row - min_row + 1) * (max_col - min_col + 1)

        # Check headers (first row)
        for col_idx in range(min_col, max_col + 1):
            cell = ws.cell(row=min_row, column=col_idx)
            val = str(cell.value or "")
            if val and (val.startswith(" ") or val.endswith(" ") or "  " in val):
                header_whitespace.append({
                    "cell": cell.coordinate,
                    "header": val,
                    "trimmed": val.strip()
                })

        # Check column types, formulas, and numbers formatted as text
        for col_idx in range(min_col, max_col + 1):
            header_cell = ws.cell(row=min_row, column=col_idx)
            h_name = str(header_cell.value or f"Col_{col_idx}")
            col_types = set()
            for r_idx in range(min_row + 1, min(max_row + 1, min_row + 1000)):
                cell = ws.cell(row=r_idx, column=col_idx)
                val = cell.value
                if isinstance(val, str) and val.startswith("="):
                    total_formulas += 1
                if val is not None:
                    col_types.add(type(val).__name__)
                    if isinstance(val, str) and val.strip().replace(".", "", 1).replace("-", "", 1).isdigit():
                        if len(numbers_as_text) < 50:
                            numbers_as_text.append({
                                "cell": cell.coordinate,
                                "column": h_name,
                                "value": val
                            })
            if len(col_types) > 1:
                inconsistent_types.append({
                    "column_index": col_idx,
                    "header": h_name,
                    "types_detected": list(col_types)
                })

        range_name = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"
        return {
            "sheet": ws.title,
            "range_audited": range_name,
            "total_cells": total_cells,
            "has_dynamic_formulas": (total_formulas > 0),
            "formula_count": total_formulas,
            "header_whitespace_issues": header_whitespace,
            "numbers_stored_as_text": numbers_as_text,
            "inconsistent_column_types": inconsistent_types
        }

    def aggregate(self, group_by_col: str, metric_col: str, agg_func: str = "sum", sheet: Optional[str] = None, top_n: int = 10, ascending: bool = False) -> Dict[str, Any]:
        ws = self._get_sheet(sheet)
        norm_group = group_by_col.strip().lower()
        norm_metric = metric_col.strip().lower()

        group_idx = None
        metric_idx = None

        for col_idx in range(1, ws.max_column + 1):
            h_val = str(ws.cell(row=1, column=col_idx).value or "").strip().lower()
            if h_val == norm_group and group_idx is None:
                group_idx = col_idx
            if h_val == norm_metric and metric_idx is None:
                metric_idx = col_idx

        if group_idx is None or metric_idx is None:
            raise ValueError(f"No se encontraron las columnas especificadas: group_by='{group_by_col}', metric='{metric_col}'")

        groups = {}
        for r_idx in range(2, ws.max_row + 1):
            g_val = ws.cell(row=r_idx, column=group_idx).value
            m_val = ws.cell(row=r_idx, column=metric_idx).value
            if g_val is None or m_val is None:
                continue

            try:
                num = float(m_val)
            except (ValueError, TypeError):
                continue

            g_key = str(g_val)
            cell_coord = ws.cell(row=r_idx, column=metric_idx).coordinate

            if g_key not in groups:
                groups[g_key] = {
                    "count": 0,
                    "values": [],
                    "cells": [],
                    "min_val": num,
                    "min_cell": cell_coord,
                    "max_val": num,
                    "max_cell": cell_coord
                }

            entry = groups[g_key]
            entry["count"] += 1
            entry["values"].append(num)
            entry["cells"].append(cell_coord)

            if num < entry["min_val"]:
                entry["min_val"] = num
                entry["min_cell"] = cell_coord
            if num > entry["max_val"]:
                entry["max_val"] = num
                entry["max_cell"] = cell_coord

        results = []
        for g_key, data in groups.items():
            vals = data["values"]
            total = sum(vals)
            avg = total / len(vals) if vals else 0.0

            if agg_func.lower() == "sum":
                agg_val = total
            elif agg_func.lower() in ("avg", "mean"):
                agg_val = avg
            elif agg_func.lower() == "min":
                agg_val = data["min_val"]
            elif agg_func.lower() == "max":
                agg_val = data["max_val"]
            elif agg_func.lower() == "count":
                agg_val = float(data["count"])
            else:
                agg_val = total

            results.append({
                "group": g_key,
                "aggregated_value": round(agg_val, 2),
                "total_sum": round(total, 2),
                "average": round(avg, 2),
                "count": data["count"],
                "max_cell": data["max_cell"],
                "max_value": round(data["max_val"], 2),
                "min_cell": data["min_cell"],
                "min_value": round(data["min_val"], 2),
                "sample_citation_cells": data["cells"][:3] + data["cells"][-3:] if len(data["cells"]) > 6 else data["cells"]
            })

        results.sort(key=lambda x: x["aggregated_value"], reverse=not ascending)
        if top_n > 0:
            results = results[:top_n]

        return {
            "sheet": ws.title,
            "group_by": group_by_col,
            "metric": metric_col,
            "agg_func": agg_func,
            "groups_count": len(results),
            "ranking": results
        }

    def add_calculated_column(self, header: str, formula_template: str, number_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT, sheet: Optional[str] = None, autofit: bool = True) -> Dict[str, Any]:
        from openpyxl.worksheet.table import TableColumn
        ws = self._get_sheet(sheet)
        new_col = ws.max_column + 1
        max_row = ws.max_row

        # Set header
        ws.cell(row=1, column=new_col, value=header)

        # Parse formula_template to adapt relative row numbers and sanitize compatibility
        sanitized_template = sanitize_formula_compatibility(formula_template)
        # If formula has row 2 (e.g. '=L2/J2') or {r}/{row}, adapt while strictly protecting absolute references ($)
        for r in range(2, max_row + 1):
            cell = ws.cell(row=r, column=new_col)
            row_formula = propagate_formula_row(sanitized_template, r, base_row=2)
            cell.value = row_formula
            if number_format:
                cell.number_format = number_format

        # Auto-resize openpyxl table if present
        table_resized = None
        for tbl in ws.tables.values():
            min_c, min_r, max_c, max_r = range_boundaries(tbl.ref)
            if new_col == max_c + 1:
                new_ref = f"{get_column_letter(min_c)}{min_r}:{get_column_letter(new_col)}{max_r}"
                tbl.ref = new_ref
                new_col_id = len(tbl.tableColumns) + 1
                tbl.tableColumns.append(TableColumn(id=new_col_id, name=header))
                table_resized = tbl.name
                break

        if autofit:
            col_letter = get_column_letter(new_col)
            max_len = max(len(str(ws.cell(row=r, column=new_col).value or '')) for r in range(1, max_row + 1))
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        return {
            "status": "success",
            "sheet": ws.title,
            "header": header,
            "column_index": new_col,
            "formula_applied": sanitized_template,
            "rows_affected": max_row - 1,
            "table_resized": table_resized
        }

    def apply_style_preset(self, range_str: str, preset: str, sheet: Optional[str] = None) -> Dict[str, Any]:
        """Aplica un preset de estilo universal a un rango en Headless OpenPyXL."""
        ws = self._get_sheet(sheet)
        if ":" in range_str:
            cells = [cell for row in ws[range_str] for cell in row]
        else:
            cells = [ws[range_str]]

        p = preset.lower()
        if p == "header":
            fill = PatternFill(start_color="0E2E63", end_color="0E2E63", fill_type="solid")
            for cell in cells:
                cell.fill = fill
                cell.font = Font(name=cell.font.name if cell.font and cell.font.name else "Segoe UI", size=cell.font.size if cell.font and cell.font.size else 11, bold=True, color="FFFFFF")
                cell.alignment = Alignment(vertical="center", horizontal=cell.alignment.horizontal if cell.alignment else None, wrap_text=cell.alignment.wrap_text if cell.alignment else None)
        elif p == "input_cell":
            fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            border = Border(
                left=Side(style='medium', color='1F4E79'),
                right=Side(style='medium', color='1F4E79'),
                top=Side(style='medium', color='1F4E79'),
                bottom=Side(style='medium', color='1F4E79')
            )
            for cell in cells:
                cell.fill = fill
                cell.font = Font(name=cell.font.name if cell.font and cell.font.name else "Segoe UI", size=cell.font.size if cell.font and cell.font.size else 11, bold=cell.font.bold if cell.font else False, color="1F4E79")
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical=cell.alignment.vertical if cell.alignment else "center")
        elif p == "total_row":
            for cell in cells:
                cell.font = Font(name=cell.font.name if cell.font and cell.font.name else "Segoe UI", size=cell.font.size if cell.font and cell.font.size else 10, bold=True, color=cell.font.color if cell.font else None)
                cell.border = Border(
                    top=Side(style='thin', color='000000'),
                    bottom=Side(style='double', color='000000'),
                    left=cell.border.left if cell.border else None,
                    right=cell.border.right if cell.border else None
                )
        elif p == "currency":
            for cell in cells:
                cell.number_format = FinancialStyleGuide.CURRENCY_FORMAT
        elif p == "percent":
            for cell in cells:
                cell.number_format = FinancialStyleGuide.PERCENT_FORMAT
        elif p == "delta_positive":
            for cell in cells:
                cell.font = Font(name=cell.font.name if cell.font and cell.font.name else "Segoe UI", size=cell.font.size if cell.font and cell.font.size else 10, bold=True, color="008000")
        elif p == "delta_negative":
            for cell in cells:
                cell.font = Font(name=cell.font.name if cell.font and cell.font.name else "Segoe UI", size=cell.font.size if cell.font and cell.font.size else 10, bold=True, color="C00000")
        else:
            raise ValueError(f"Preset no reconocido: '{preset}'. Presets válidos: header, input_cell, total_row, currency, percent, delta_positive, delta_negative")

        return {
            "status": "success",
            "range": range_str,
            "preset": preset,
            "sheet": ws.title,
            "cells_affected": len(cells)
        }

    def create_summary_table(
        self,
        title: str,
        headers: List[str],
        rows_spec: List[Dict[str, Any]],
        dest_sheet: str,
        start_cell: str = "B2",
        parameters: Optional[List[Dict[str, Any]]] = None
    ) -> Dict[str, Any]:
        """
        Crea una tabla resumen / comparativa profesional en Headless OpenPyXL con parámetros vinculados,
        encabezados corporativos, filas calculadas/estáticas, totales formateados y cuadrícula activa.
        """
        if dest_sheet in self.wb.sheetnames:
            ws = self.wb[dest_sheet]
        else:
            ws = self.wb.create_sheet(title=dest_sheet)

        # Activar visualización de cuadrícula
        if ws.views and ws.views.sheetView:
            ws.views.sheetView[0].showGridLines = True

        start_c_str, start_r_idx = coordinate_from_string(start_cell)
        start_col_idx = column_index_from_string(start_c_str)
        curr_r = start_r_idx

        # 1. Bloque de Parámetros de Entrada (si se especifican)
        param_cells: Dict[str, str] = {}
        if parameters:
            param_title = ws.cell(row=curr_r, column=start_col_idx, value="Supuestos y Parámetros")
            param_title.font = Font(name="Segoe UI", size=11, bold=True, color="0E2E63")
            curr_r += 1

            for param in parameters:
                param_name = param.get("name") or param.get("label", "Parámetro")
                param_val = param.get("value")
                param_format = param.get("format")
                param_note = param.get("note")

                if param.get("cell"):
                    val_coord = param["cell"]
                    v_cell = ws[val_coord]
                    v_cell.value = param_val
                else:
                    lbl_cell = ws.cell(row=curr_r, column=start_col_idx, value=param_name)
                    lbl_cell.font = Font(name="Segoe UI", size=10, bold=True)
                    v_cell = ws.cell(row=curr_r, column=start_col_idx + 1, value=param_val)
                    val_coord = v_cell.coordinate
                    curr_r += 1

                self.apply_style_preset(val_coord, "input_cell", sheet=dest_sheet)
                if param_format:
                    if param_format.lower() == "percent":
                        self.apply_style_preset(val_coord, "percent", sheet=dest_sheet)
                    elif param_format.lower() == "currency":
                        self.apply_style_preset(val_coord, "currency", sheet=dest_sheet)
                    else:
                        ws[val_coord].number_format = param_format

                if param_note:
                    ws[val_coord].comment = Comment(param_note, "Antigravity")

                param_cells[param_name] = val_coord

            curr_r += 1

        # 2. Título general de la tabla
        tbl_title_row = curr_r
        tbl_title_cell = ws.cell(row=tbl_title_row, column=start_col_idx, value=title)
        tbl_title_cell.font = Font(name="Segoe UI", size=13, bold=True, color="0E2E63")
        curr_r += 1

        # 3. Fila de encabezados
        header_row = curr_r
        end_col_idx = start_col_idx + len(headers) - 1
        for idx, h in enumerate(headers):
            ws.cell(row=header_row, column=start_col_idx + idx, value=h)

        hdr_range = f"{get_column_letter(start_col_idx)}{header_row}:{get_column_letter(end_col_idx)}{header_row}"
        self.apply_style_preset(hdr_range, "header", sheet=dest_sheet)
        curr_r += 1

        # 4. Filas de datos y totales
        for row_spec in rows_spec:
            is_total = row_spec.get("is_total", False) or row_spec.get("total", False)
            row_range = f"{get_column_letter(start_col_idx)}{curr_r}:{get_column_letter(end_col_idx)}{curr_r}"

            label = row_spec.get("label") or row_spec.get("name") or row_spec.get("title")
            row_vals = []
            if "values" in row_spec:
                vals = row_spec["values"]
                if len(vals) == len(headers):
                    row_vals = vals
                elif len(vals) == len(headers) - 1:
                    row_vals = [label] + vals
                else:
                    row_vals = vals
            elif "formulas" in row_spec:
                forms = row_spec["formulas"]
                if len(forms) == len(headers):
                    row_vals = forms
                elif len(forms) == len(headers) - 1:
                    row_vals = [label] + forms
                else:
                    row_vals = forms
            else:
                row_vals = [row_spec.get(h, "") for h in headers]
                if label and not row_vals[0]:
                    row_vals[0] = label

            formats = row_spec.get("formats")
            row_fmt = row_spec.get("format") or row_spec.get("number_format")
            presets = row_spec.get("presets") or row_spec.get("styles")

            for c_idx, val in enumerate(row_vals):
                if c_idx >= len(headers):
                    break
                cell = ws.cell(row=curr_r, column=start_col_idx + c_idx)
                if isinstance(val, str) and val.startswith("="):
                    cell.value = sanitize_formula_compatibility(val)
                else:
                    cell.value = val

                if not is_total:
                    cell.font = Font(name="Segoe UI", size=10)
                    cell.border = Border(
                        left=Side(style='thin', color='D9D9D9'),
                        right=Side(style='thin', color='D9D9D9'),
                        top=Side(style='thin', color='D9D9D9'),
                        bottom=Side(style='thin', color='D9D9D9')
                    )

                col_fmt = formats[c_idx] if (formats and c_idx < len(formats)) else row_fmt
                if col_fmt:
                    if col_fmt.lower() == "currency":
                        cell.number_format = FinancialStyleGuide.CURRENCY_FORMAT
                    elif col_fmt.lower() == "percent":
                        cell.number_format = FinancialStyleGuide.PERCENT_FORMAT
                    else:
                        cell.number_format = col_fmt

                if presets and c_idx < len(presets) and presets[c_idx]:
                    self.apply_style_preset(cell.coordinate, presets[c_idx], sheet=dest_sheet)

            if is_total:
                self.apply_style_preset(row_range, "total_row", sheet=dest_sheet)

            curr_r += 1

        # 5. AutoFit en columnas involucradas
        for c_idx in range(start_col_idx, end_col_idx + 1):
            col_letter = get_column_letter(c_idx)
            max_len = max(len(str(ws.cell(row=r, column=c_idx).value or '')) for r in range(start_r_idx, curr_r))
            ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

        return {
            "status": "success",
            "sheet": dest_sheet,
            "title": title,
            "table_range": f"{get_column_letter(start_col_idx)}{header_row}:{get_column_letter(end_col_idx)}{curr_r - 1}",
            "rows_count": len(rows_spec),
            "headers": headers,
            "parameters": param_cells,
            "start_cell": start_cell
        }

    def create_pivot_table(
        self,
        source: Union[str, Dict[str, Any]],
        rows: List[str],
        cols: Optional[List[str]] = None,
        value_field: Optional[str] = None,
        value_func: str = "sum",
        value_format: Optional[str] = FinancialStyleGuide.CURRENCY_FORMAT,
        dest_sheet: Optional[str] = None,
        dest_cell: str = "A3",
        table_name: Optional[str] = None,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Crea una hoja de análisis con tabla de doble entrada cruzada (cross-tabulation resumida)
        con totales generales formateados en Headless OpenPyXL.
        """
        src_ws = self._get_sheet(source_sheet) if source_sheet else self.wb.active

        # 1. Resolver rango origen y encabezados
        if isinstance(source, dict) and "range" in source:
            source_range_str = source["range"]
        else:
            source_range_str = str(source)

        min_col, min_row, max_col, max_row = 1, 1, src_ws.max_column, src_ws.max_row
        if ":" in source_range_str:
            min_col, min_row, max_col, max_row = range_boundaries(source_range_str)
        elif source_range_str in src_ws.tables:
            min_col, min_row, max_col, max_row = range_boundaries(src_ws.tables[source_range_str].ref)

        headers = {}
        for c in range(min_col, max_col + 1):
            val = src_ws.cell(row=min_row, column=c).value
            if val is not None:
                headers[str(val).strip().lower()] = c

        row_key_name = rows[0] if rows else ""
        col_key_name = cols[0] if cols else None
        metric_name = value_field or ""

        row_c_idx = headers.get(row_key_name.strip().lower())
        col_c_idx = headers.get(col_key_name.strip().lower()) if col_key_name else None
        val_c_idx = headers.get(metric_name.strip().lower())

        if not row_c_idx:
            raise ValueError(f"Campo de fila '{row_key_name}' no encontrado en encabezados.")
        if not val_c_idx:
            raise ValueError(f"Campo de valor '{metric_name}' no encontrado en encabezados.")

        # 2. Agregar datos cruzados en memoria
        # data_matrix[row_val][col_val] = list of floats
        data_matrix: Dict[str, Dict[str, List[float]]] = {}
        row_values_ordered = []
        col_values_ordered = []

        for r in range(min_row + 1, max_row + 1):
            r_val = src_ws.cell(row=r, column=row_c_idx).value
            v_val = src_ws.cell(row=r, column=val_c_idx).value
            c_val = src_ws.cell(row=r, column=col_c_idx).value if col_c_idx else "Total"

            if r_val is None:
                continue
            try:
                num = float(v_val)
            except (ValueError, TypeError):
                continue

            r_str = str(r_val).strip()
            c_str = str(c_val).strip()

            if r_str not in data_matrix:
                data_matrix[r_str] = {}
                row_values_ordered.append(r_str)
            if c_str not in data_matrix[r_str]:
                data_matrix[r_str][c_str] = []
            if c_str not in col_values_ordered:
                col_values_ordered.append(c_str)

            data_matrix[r_str][c_str].append(num)

        # 3. Preparar hoja de destino
        target_sheet_name = dest_sheet or (src_ws.title + "_Pivot")
        if target_sheet_name in self.wb.sheetnames:
            dest_ws = self.wb[target_sheet_name]
        else:
            dest_ws = self.wb.create_sheet(title=target_sheet_name)

        # Parse dest_cell
        start_c_idx, start_r_idx = coordinate_from_string(dest_cell)
        start_col_num = column_index_from_string(start_c_idx)

        # 4. Escribir tabla cruzada
        # Estilos corporativos
        header_fill = PatternFill(start_color="0E2E63", end_color="0E2E63", fill_type="solid")
        header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
        total_font = Font(name="Segoe UI", size=10, bold=True)
        data_font = Font(name="Segoe UI", size=10)
        thin_border = Border(
            left=Side(style='thin', color='D9D9D9'),
            right=Side(style='thin', color='D9D9D9'),
            top=Side(style='thin', color='D9D9D9'),
            bottom=Side(style='thin', color='D9D9D9')
        )
        total_border = Border(
            top=Side(style='thin', color='000000'),
            bottom=Side(style='double', color='000000')
        )

        def calc_agg(nums: List[float]) -> float:
            if not nums:
                return 0.0
            if value_func.lower() == "sum":
                return sum(nums)
            elif value_func.lower() in ("avg", "average", "mean"):
                return sum(nums) / len(nums)
            elif value_func.lower() == "count":
                return float(len(nums))
            elif value_func.lower() == "min":
                return min(nums)
            elif value_func.lower() == "max":
                return max(nums)
            return sum(nums)

        # Escribir encabezados de columnas
        # Fila de encabezado: [Row_Field_Name, Col1, Col2, ..., Total General]
        header_r = start_r_idx
        # Celda de fila
        c_cell = dest_ws.cell(row=header_r, column=start_col_num, value=row_key_name)
        c_cell.font = header_font
        c_cell.fill = header_fill

        num_data_cols = len(col_values_ordered)
        for idx, col_val in enumerate(col_values_ordered):
            cell = dest_ws.cell(row=header_r, column=start_col_num + 1 + idx, value=col_val)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="right")

        # Columna Total General
        tot_col_idx = start_col_num + 1 + num_data_cols
        tot_header_cell = dest_ws.cell(row=header_r, column=tot_col_idx, value="Total General")
        tot_header_cell.font = header_font
        tot_header_cell.fill = header_fill
        tot_header_cell.alignment = Alignment(horizontal="right")

        # Escribir filas de datos
        curr_r = header_r + 1
        col_accumulators: Dict[str, List[float]] = {c: [] for c in col_values_ordered}
        all_nums: List[float] = []

        for r_val in row_values_ordered:
            row_label_cell = dest_ws.cell(row=curr_r, column=start_col_num, value=r_val)
            row_label_cell.font = data_font
            row_label_cell.border = thin_border

            row_nums = []
            for idx, col_val in enumerate(col_values_ordered):
                nums = data_matrix[r_val].get(col_val, [])
                agg_res = calc_agg(nums)
                cell = dest_ws.cell(row=curr_r, column=start_col_num + 1 + idx, value=agg_res)
                cell.font = data_font
                cell.border = thin_border
                if value_format:
                    cell.number_format = value_format
                if nums:
                    col_accumulators[col_val].extend(nums)
                    row_nums.extend(nums)

            # Total de la fila
            row_tot = calc_agg(row_nums)
            tot_cell = dest_ws.cell(row=curr_r, column=tot_col_idx, value=row_tot)
            tot_cell.font = total_font
            tot_cell.border = thin_border
            if value_format:
                tot_cell.number_format = value_format
            all_nums.extend(row_nums)
            curr_r += 1

        # Fila de Total General al pie
        footer_r = curr_r
        foot_label_cell = dest_ws.cell(row=footer_r, column=start_col_num, value="Total General")
        foot_label_cell.font = total_font
        foot_label_cell.border = total_border

        for idx, col_val in enumerate(col_values_ordered):
            col_tot = calc_agg(col_accumulators[col_val])
            cell = dest_ws.cell(row=footer_r, column=start_col_num + 1 + idx, value=col_tot)
            cell.font = total_font
            cell.border = total_border
            if value_format:
                cell.number_format = value_format

        # Total absoluto
        grand_total = calc_agg(all_nums)
        grand_tot_cell = dest_ws.cell(row=footer_r, column=tot_col_idx, value=grand_total)
        grand_tot_cell.font = total_font
        grand_tot_cell.border = total_border
        if value_format:
            grand_tot_cell.number_format = value_format

        # AutoFit en columnas de la hoja destino
        total_cols_span = tot_col_idx - start_col_num + 1
        for col_idx in range(start_col_num, tot_col_idx + 1):
            col_letter = get_column_letter(col_idx)
            max_len = max(len(str(dest_ws.cell(row=r, column=col_idx).value or '')) for r in range(header_r, footer_r + 1))
            dest_ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

        end_cell = f"{get_column_letter(tot_col_idx)}{footer_r}"
        pt_range = f"{dest_cell}:{end_cell}"

        return {
            "status": "success",
            "pivot_table_name": table_name or "Pivot_Summary",
            "dest_sheet": dest_ws.title,
            "dest_range": pt_range,
            "rows": rows,
            "cols": cols or [],
            "value_field": metric_name,
            "value_func": value_func,
            "value_format": value_format,
            "grand_total": round(grand_total, 2)
        }

    def create_chart(
        self,
        source: Union[str, Dict[str, Any]],
        chart_type: str = "column_clustered",
        title: Optional[str] = None,
        dest_sheet: Optional[str] = None,
        placement: str = "auto",
        anchor: Optional[str] = None,
        width: int = 480,
        height: int = 300,
        source_sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Inserta un gráfico nativo openpyxl.chart en la posición calculada anti-solapamiento.
        """
        from openpyxl.chart import BarChart, LineChart, PieChart, AreaChart, Reference

        target_ws = self._get_sheet(dest_sheet) if dest_sheet else self.wb.active

        # Resolver rango fuente
        if isinstance(source, dict):
            if "dest_sheet" in source and not dest_sheet:
                target_ws = self._get_sheet(source["dest_sheet"])
            source_range_str = source.get("dest_range", "A3:B10")
        else:
            source_range_str = str(source)

        min_col, min_row, max_col, max_row = range_boundaries(source_range_str)

        # 1. Instanciar clase de gráfico correspondiente
        ct = chart_type.lower()
        if ct in ("column_clustered", "column"):
            chart = BarChart()
            chart.type = "col"
            chart.grouping = "clustered"
        elif ct in ("bar_clustered", "bar"):
            chart = BarChart()
            chart.type = "bar"
            chart.grouping = "clustered"
        elif ct == "line":
            chart = LineChart()
        elif ct == "pie":
            chart = PieChart()
        elif ct == "area":
            chart = AreaChart()
        else:
            chart = BarChart()
            chart.type = "col"

        # 2. Configurar título y dimensiones
        if title:
            chart.title = title
        chart.width = width / 30.0 # openpyxl usa cm aprox (480px ~ 16cm)
        chart.height = height / 30.0 # (300px ~ 10cm)

        # 3. Referencias de datos y categorías
        # Suponiendo que la primera columna son categorías y la última fila puede ser total general
        # Si la última fila es 'Total General', excluirla de la serie para no desvirtuar el gráfico
        last_row_label = str(target_ws.cell(row=max_row, column=min_col).value or "").strip().lower()
        data_end_row = max_row - 1 if last_row_label == "total general" else max_row

        # Si hay más de 2 columnas y la última columna es 'Total General', excluirla si hay desglose por columnas
        last_col_header = str(target_ws.cell(row=min_row, column=max_col).value or "").strip().lower()
        data_end_col = max_col - 1 if (last_col_header == "total general" and max_col - min_col >= 2) else max_col

        data_ref = Reference(target_ws, min_col=min_col + 1, min_row=min_row, max_col=data_end_col, max_row=data_end_row)
        cats_ref = Reference(target_ws, min_col=min_col, min_row=min_row + 1, max_row=data_end_row)

        chart.add_data(data_ref, titles_from_data=True)
        chart.set_categories(cats_ref)

        # 4. Calcular posición geométrica anti-solapamiento
        num_cols = max_col - min_col + 1
        num_rows = max_row - min_row + 1
        pos = self._calculate_placement_coordinates(
            start_col=min_col,
            start_row=min_row,
            num_cols=num_cols,
            num_rows=num_rows,
            placement=placement,
            anchor=anchor
        )

        anchor_cell = pos["cell"]
        target_ws.add_chart(chart, anchor_cell)

        return {
            "status": "success",
            "sheet": target_ws.title,
            "chart_type": chart_type,
            "title": title,
            "placement": pos["placement"],
            "anchor_cell": anchor_cell,
            "data_range": f"{get_column_letter(min_col + 1)}{min_row}:{get_column_letter(data_end_col)}{data_end_row}"
        }

    def save(self, file_path: Optional[str] = None) -> str:
        out_path = file_path or self.file_path
        if out_path:
            self.wb.save(out_path)
            return os.path.abspath(out_path)
        raise ValueError("No se especificó ruta para guardar el archivo.")

    def remove_duplicates(self, range_or_table: str, key_columns: Optional[List[str]] = None, sheet: Optional[str] = None) -> Dict[str, Any]:
        """
        Elimina filas duplicadas dentro del rango o tabla estructurada on-demand.
        En OpenPyXL, elimina filas y redimensiona tbl.ref en la tabla estructurada si existe.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if range_or_table and range_or_table in ws.tables:
            tbl = ws.tables[range_or_table]
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        elif range_or_table:
            min_col, min_row, max_col, max_row = range_boundaries(range_or_table)
            for t in ws.tables.values():
                if t.ref == range_or_table:
                    tbl = t
                    break
        elif ws.tables:
            tbl = next(iter(ws.tables.values()))
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        num_cols = max_col - min_col + 1
        data_rows_count = max_row - min_row
        if data_rows_count <= 0:
            return {
                "status": "success",
                "range_or_table": range_or_table or (tbl.name if tbl else f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"),
                "sheet": ws.title,
                "original_rows": 0,
                "duplicates_removed": 0,
                "remaining_rows": 0,
                "key_columns": key_columns or "all"
            }

        headers = [str(ws.cell(row=min_row, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        headers_lower = [h.lower() for h in headers]

        compare_cols = []
        if key_columns:
            for k in key_columns:
                k_clean = str(k).strip()
                k_lower = k_clean.lower()
                if k_lower in headers_lower:
                    idx = headers_lower.index(k_lower) + 1
                    compare_cols.append(idx)
                elif k_clean.isalpha() and len(k_clean) <= 3:
                    c_idx = column_index_from_string(k_clean)
                    rel_idx = c_idx - min_col + 1
                    if 1 <= rel_idx <= num_cols:
                        compare_cols.append(rel_idx)
            if not compare_cols:
                compare_cols = list(range(1, num_cols + 1))
        else:
            compare_cols = list(range(1, num_cols + 1))

        seen = set()
        dup_row_indices = []

        for r in range(min_row + 1, max_row + 1):
            row_vals = [ws.cell(row=r, column=min_col + c - 1).value for c in range(1, num_cols + 1)]
            key = tuple(row_vals[ci - 1] for ci in compare_cols)
            if key in seen:
                dup_row_indices.append(r)
            else:
                seen.add(key)

        duplicates_count = len(dup_row_indices)
        if duplicates_count > 0:
            for r in sorted(dup_row_indices, reverse=True):
                ws.delete_rows(r, 1)

            if tbl:
                new_max_row = max_row - duplicates_count
                tbl.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

        remaining_data_rows = data_rows_count - duplicates_count
        return {
            "status": "success",
            "range_or_table": range_or_table or (tbl.name if tbl else f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"),
            "sheet": ws.title,
            "original_rows": data_rows_count,
            "duplicates_removed": duplicates_count,
            "remaining_rows": remaining_data_rows,
            "key_columns": key_columns or "all"
        }

    def detect_and_flag_outliers(
        self,
        column: str,
        method: str = "iqr",
        range_or_table: Optional[str] = None,
        sheet: Optional[str] = None,
        flag_style: str = "warning",
        add_comment: bool = True
    ) -> Dict[str, Any]:
        """
        Detecta y señala visualmente valores atípicos (outliers) mediante IQR o Z-score on-demand.
        Aplica formato no destructivo (#FCE4D6 y texto rojo bold) y añade comentarios a las celdas.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if range_or_table and range_or_table in ws.tables:
            tbl = ws.tables[range_or_table]
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        elif range_or_table:
            min_col, min_row, max_col, max_row = range_boundaries(range_or_table)
            for t in ws.tables.values():
                if t.ref == range_or_table:
                    tbl = t
                    break
        elif ws.tables:
            tbl = next(iter(ws.tables.values()))
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        num_cols = max_col - min_col + 1
        headers = [str(ws.cell(row=min_row, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        headers_lower = [h.lower() for h in headers]
        col_clean = str(column).strip()
        col_lower = col_clean.lower()

        target_col_rel = None
        if col_lower in headers_lower:
            target_col_rel = headers_lower.index(col_lower) + 1
        elif col_clean.isalpha() and len(col_clean) <= 3:
            c_idx = column_index_from_string(col_clean)
            rel_idx = c_idx - min_col + 1
            if 1 <= rel_idx <= num_cols:
                target_col_rel = rel_idx

        if not target_col_rel:
            raise ValueError(f"Columna '{column}' no encontrada en el rango/tabla. Encabezados disponibles: {headers}")

        abs_col_idx = min_col + target_col_rel - 1

        val_list = []
        row_val_map = {}
        for r in range(min_row + 1, max_row + 1):
            v = ws.cell(row=r, column=abs_col_idx).value
            if v is not None and str(v).strip() != "":
                try:
                    fv = float(str(v).replace("$", "").replace(" ", "").replace(",", ""))
                    val_list.append(fv)
                    row_val_map[r] = fv
                except Exception:
                    pass

        if len(val_list) < 4:
            return {
                "status": "success",
                "sheet": ws.title,
                "column": column,
                "method": method,
                "total_values_evaluated": len(val_list),
                "outliers_count": 0,
                "thresholds": {},
                "flagged_cells": [],
                "outliers": []
            }

        method_clean = method.lower()
        if method_clean == "zscore":
            mean = statistics.mean(val_list)
            std = statistics.stdev(val_list) if len(val_list) > 1 else 0
            z_thresh = 2.5 if len(val_list) < 30 else 3.0
            lower_bound = mean - z_thresh * std
            upper_bound = mean + z_thresh * std
            thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "mean": mean, "std": std, "z_threshold": z_thresh}
        else:
            sorted_vals = sorted(val_list)
            q1, _, q3 = statistics.quantiles(sorted_vals, n=4, method="inclusive")
            iqr = q3 - q1
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "q1": q1, "q3": q3, "iqr": iqr}

        outliers = []
        col_letter = get_column_letter(abs_col_idx)

        warning_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        warning_font = Font(color="C00000", bold=True)

        for r, fv in row_val_map.items():
            if fv < lower_bound or fv > upper_bound:
                cell = ws.cell(row=r, column=abs_col_idx)
                cell_coord = f"{col_letter}{r}"
                cell.fill = warning_fill
                cell.font = warning_font

                if add_comment:
                    try:
                        from openpyxl.comments import Comment
                        cell.comment = Comment(
                            f"Outlier detectado ({method_clean.upper()}): valor={fv}, límites=[{lower_bound:.2f}, {upper_bound:.2f}]",
                            "Antigravity"
                        )
                    except Exception:
                        pass

                outliers.append({
                    "cell": cell_coord,
                    "row": r,
                    "value": fv,
                    "reason": f"Valor {fv} fuera de límites [{lower_bound:.2f}, {upper_bound:.2f}]"
                })

        return {
            "status": "success",
            "sheet": ws.title,
            "column": headers[target_col_rel - 1],
            "method": method_clean,
            "total_values_evaluated": len(val_list),
            "outliers_count": len(outliers),
            "thresholds": thresholds,
            "flagged_cells": [o["cell"] for o in outliers],
            "outliers": outliers
        }

    def clean_table_dataset(self, target_range_or_table: str, rules: Dict[str, Any], sheet: Optional[str] = None) -> Dict[str, Any]:
        """
        Limpia y estandariza en bloque los datos de una tabla o rango on-demand.
        Procesa en una sola pasada en memoria:
          - standardize_text: trim, casing y mapeo de sinónimos/reemplazos.
          - cast_types: conversión segura de tipos (float, int, str).
          - coerce_dates: conversión a seriales de Excel o datetimes con formato uniforme.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if target_range_or_table and target_range_or_table in ws.tables:
            tbl = ws.tables[target_range_or_table]
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        elif target_range_or_table:
            min_col, min_row, max_col, max_row = range_boundaries(target_range_or_table)
            for t in ws.tables.values():
                if t.ref == target_range_or_table:
                    tbl = t
                    break
        elif ws.tables:
            tbl = next(iter(ws.tables.values()))
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        num_cols = max_col - min_col + 1
        data_rows_count = max_row - min_row

        headers = [str(ws.cell(row=min_row, column=c).value or "").strip() for c in range(min_col, max_col + 1)]

        std_text_rules = rules.get("standardize_text", {})
        cast_type_rules = rules.get("cast_types", {})
        coerce_date_rules = rules.get("coerce_dates", {})

        col_text_cfg = {}
        col_cast_cfg = {}
        col_date_cfg = {}

        for c_idx, h in enumerate(headers, start=1):
            h_lower = h.lower()
            for rule_col, cfg in std_text_rules.items():
                if rule_col.strip().lower() == h_lower:
                    col_text_cfg[c_idx] = cfg
                    break

            for rule_col, target_t in cast_type_rules.items():
                if rule_col.strip().lower() == h_lower:
                    col_cast_cfg[c_idx] = str(target_t).lower()
                    break

            if isinstance(coerce_date_rules, list):
                if any(dc.strip().lower() == h_lower for dc in coerce_date_rules):
                    col_date_cfg[c_idx] = {"to_serial": True, "format": "yyyy-mm-dd"}
            elif isinstance(coerce_date_rules, dict):
                for dc, cfg in coerce_date_rules.items():
                    if dc.strip().lower() == h_lower:
                        col_date_cfg[c_idx] = cfg if isinstance(cfg, dict) else {"to_serial": True, "format": "yyyy-mm-dd"}
                        break

        cells_modified = 0

        if rules.get("trim_headers", True):
            for c_idx in range(1, num_cols + 1):
                cell = ws.cell(row=min_row, column=min_col + c_idx - 1)
                if isinstance(cell.value, str) and cell.value != cell.value.strip():
                    cell.value = cell.value.strip()
                    cells_modified += 1

        for r in range(min_row + 1, max_row + 1):
            for c_idx in range(1, num_cols + 1):
                cell = ws.cell(row=r, column=min_col + c_idx - 1)
                val = cell.value
                orig_val = val

                # 1. Date coercion
                if c_idx in col_date_cfg:
                    serial = parse_excel_date_serial(val)
                    if serial is not None:
                        val = serial
                        d_cfg = col_date_cfg[c_idx]
                        fmt = d_cfg.get("format", "yyyy-mm-dd") if isinstance(d_cfg, dict) else "yyyy-mm-dd"
                        cell.number_format = fmt

                # 2. Text standardization
                if c_idx in col_text_cfg and val is not None:
                    t_cfg = col_text_cfg[c_idx]
                    s = str(val)
                    if t_cfg.get("trim", True):
                        s = s.strip()
                    reps = t_cfg.get("replacements") or t_cfg.get("mapping")
                    if reps:
                        if s in reps:
                            s = reps[s]
                        elif s.lower() in {k.lower(): v for k, v in reps.items()}:
                            for k, v in reps.items():
                                if s.lower() == k.lower():
                                    s = v
                                    break
                    casing = t_cfg.get("casing") or t_cfg.get("case")
                    if casing == "upper":
                        s = s.upper()
                    elif casing == "lower":
                        s = s.lower()
                    elif casing == "title":
                        s = s.title()
                    elif casing == "capitalize":
                        s = s.capitalize()
                    val = s

                # 3. Type casting
                if c_idx in col_cast_cfg and val is not None and str(val).strip() != "":
                    tgt = col_cast_cfg[c_idx]
                    try:
                        clean_str = str(val).replace("$", "").replace(" ", "").replace(",", "")
                        if tgt in ("float", "double", "decimal"):
                            val = float(clean_str)
                        elif tgt in ("int", "integer"):
                            val = int(round(float(clean_str)))
                        elif tgt in ("str", "string", "text"):
                            val = str(val)
                    except Exception:
                        pass

                if val != orig_val:
                    cell.value = val
                    cells_modified += 1

        return {
            "status": "success",
            "sheet": ws.title,
            "target_range_or_table": target_range_or_table or (tbl.name if tbl else f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"),
            "rows_processed": data_rows_count,
            "cells_modified": cells_modified,
            "rules_applied": list(rules.keys())
        }

    def curate_pipeline(
        self,
        target_range_or_table: str,
        pipeline_spec: Dict[str, Any],
        sheet: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Pipeline unificado de curaduría y vectorización en memoria para OpenPyXL headless.
        Ejecuta deduplicación, estandarización de texto, casteo de tipos, coerción de fechas
        y detección/marcado de outliers en una sola pasada en memoria.
        """
        ws = self._get_sheet(sheet)
        tbl = None
        if target_range_or_table and target_range_or_table in ws.tables:
            tbl = ws.tables[target_range_or_table]
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        elif target_range_or_table:
            min_col, min_row, max_col, max_row = range_boundaries(target_range_or_table)
            for t in ws.tables.values():
                if t.ref == target_range_or_table:
                    tbl = t
                    break
        elif ws.tables:
            tbl = next(iter(ws.tables.values()))
            min_col, min_row, max_col, max_row = range_boundaries(tbl.ref)
        else:
            min_col, min_row, max_col, max_row = 1, 1, ws.max_column, ws.max_row

        num_cols = max_col - min_col + 1
        data_rows_count = max_row - min_row
        if data_rows_count <= 0:
            return {
                "status": "success",
                "sheet": ws.title,
                "target_range_or_table": target_range_or_table or (tbl.name if tbl else f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"),
                "original_rows": 0,
                "remaining_rows": 0,
                "duplicates_removed": 0,
                "cells_cleaned": 0,
                "rules_applied": [],
                "outliers_count": 0,
                "outliers_flagged": [],
                "outliers_summary": []
            }

        # 1. Lectura a memoria
        headers = [str(ws.cell(row=min_row, column=c).value or "").strip() for c in range(min_col, max_col + 1)]
        data_rows = []
        for r in range(min_row + 1, max_row + 1):
            row_vals = [ws.cell(row=r, column=min_col + c - 1).value for c in range(1, num_cols + 1)]
            data_rows.append(row_vals)
        orig_data_rows_count = len(data_rows)

        # 2. Deduplicación en memoria
        dedup_spec = pipeline_spec.get("dedup")
        duplicates_count = 0
        if dedup_spec:
            if isinstance(dedup_spec, dict):
                key_cols = dedup_spec.get("key_columns")
            elif isinstance(dedup_spec, (list, tuple)):
                key_cols = list(dedup_spec)
            else:
                key_cols = None

            headers_lower = [h.lower() for h in headers]
            compare_cols = []
            if key_cols:
                for k in key_cols:
                    k_clean = str(k).strip()
                    k_lower = k_clean.lower()
                    if k_lower in headers_lower:
                        compare_cols.append(headers_lower.index(k_lower) + 1)
                    elif k_clean.isalpha() and len(k_clean) <= 3:
                        c_idx = column_index_from_string(k_clean)
                        rel_idx = c_idx - min_col + 1
                        if 1 <= rel_idx <= num_cols:
                            compare_cols.append(rel_idx)
                if not compare_cols:
                    compare_cols = list(range(1, num_cols + 1))
            else:
                compare_cols = list(range(1, num_cols + 1))

            seen = set()
            unique_rows = []
            for r in data_rows:
                key = tuple(
                    str(r[ci - 1]).strip() if isinstance(r[ci - 1], str) else r[ci - 1]
                    if ci - 1 < len(r) else None
                    for ci in compare_cols
                )
                if key not in seen:
                    seen.add(key)
                    unique_rows.append(r)
            duplicates_count = len(data_rows) - len(unique_rows)
            data_rows = unique_rows

        # 3. Limpieza y estandarización en memoria
        clean_spec = pipeline_spec.get("clean")
        cells_modified = 0
        col_date_cfg = {}
        if clean_spec:
            std_text_rules = clean_spec.get("standardize_text", {})
            cast_type_rules = clean_spec.get("cast_types", {})
            coerce_date_rules = clean_spec.get("coerce_dates", {})
            trim_headers = clean_spec.get("trim_headers", True)

            col_text_cfg = {}
            col_cast_cfg = {}

            for c_idx, h in enumerate(headers, start=1):
                h_lower = h.lower()
                for rule_col, cfg in std_text_rules.items():
                    if rule_col.strip().lower() == h_lower:
                        col_text_cfg[c_idx] = cfg
                        break
                for rule_col, tgt_t in cast_type_rules.items():
                    if rule_col.strip().lower() == h_lower:
                        col_cast_cfg[c_idx] = str(tgt_t).lower()
                        break
                if isinstance(coerce_date_rules, (list, tuple)):
                    if any(dc.strip().lower() == h_lower for dc in coerce_date_rules):
                        col_date_cfg[c_idx] = {"to_serial": True, "format": "yyyy-mm-dd"}
                elif isinstance(coerce_date_rules, dict):
                    for dc, cfg in coerce_date_rules.items():
                        if dc.strip().lower() == h_lower:
                            col_date_cfg[c_idx] = cfg if isinstance(cfg, dict) else {"to_serial": True, "format": "yyyy-mm-dd"}
                            break

            if trim_headers:
                new_headers = []
                for h in headers:
                    th = h.strip() if isinstance(h, str) else h
                    if th != h:
                        cells_modified += 1
                    new_headers.append(th)
                headers = new_headers

            cleaned_data_rows = []
            for r in data_rows:
                new_row = []
                for c_idx in range(1, num_cols + 1):
                    val = r[c_idx - 1] if c_idx - 1 < len(r) else None
                    orig_val = val

                    # 1. Coerción de fechas
                    if c_idx in col_date_cfg:
                        serial = parse_excel_date_serial(val)
                        if serial is not None:
                            val = serial

                    # 2. Estandarización de texto
                    if c_idx in col_text_cfg and val is not None:
                        t_cfg = col_text_cfg[c_idx]
                        s = str(val)
                        if t_cfg.get("trim", True):
                            s = s.strip()
                        reps = t_cfg.get("replacements") or t_cfg.get("mapping")
                        if reps:
                            if s in reps:
                                s = reps[s]
                            elif s.lower() in {k.lower(): v for k, v in reps.items()}:
                                for k, v in reps.items():
                                    if s.lower() == k.lower():
                                        s = v
                                        break
                        casing = t_cfg.get("casing") or t_cfg.get("case")
                        if casing == "upper":
                            s = s.upper()
                        elif casing == "lower":
                            s = s.lower()
                        elif casing == "title":
                            s = s.title()
                        elif casing == "capitalize":
                            s = s.capitalize()
                        val = s

                    # 3. Casteo seguro de tipos
                    if c_idx in col_cast_cfg and val is not None and str(val).strip() != "":
                        tgt = col_cast_cfg[c_idx]
                        try:
                            clean_str = str(val).replace("$", "").replace(" ", "").replace(",", "")
                            if tgt in ("float", "double", "decimal"):
                                val = float(clean_str)
                            elif tgt in ("int", "integer"):
                                val = int(round(float(clean_str)))
                            elif tgt in ("str", "string", "text"):
                                val = str(val)
                        except Exception:
                            pass

                    if val != orig_val:
                        cells_modified += 1
                    new_row.append(val)
                cleaned_data_rows.append(new_row)
            data_rows = cleaned_data_rows

        # 4. Cálculo estadístico de outliers en memoria
        outliers_spec = pipeline_spec.get("outliers")
        outliers_results = []
        all_flagged_cells = []
        outlier_cells_to_format = []

        if outliers_spec:
            if isinstance(outliers_spec, dict):
                if "column" in outliers_spec or "columns" in outliers_spec:
                    spec_list = [outliers_spec]
                else:
                    spec_list = [{"column": k, **(v if isinstance(v, dict) else {})} for k, v in outliers_spec.items()]
            elif isinstance(outliers_spec, (list, tuple)):
                spec_list = []
                for it in outliers_spec:
                    if isinstance(it, str):
                        spec_list.append({"column": it})
                    elif isinstance(it, dict):
                        spec_list.append(it)
            else:
                spec_list = []

            headers_lower = [h.lower() for h in headers]
            for s_item in spec_list:
                target_col_name = str(s_item.get("column") or "").strip()
                method = str(s_item.get("method") or "iqr").lower()
                add_comment = bool(s_item.get("add_comment", True))

                target_col_rel = None
                if target_col_name.lower() in headers_lower:
                    target_col_rel = headers_lower.index(target_col_name.lower()) + 1
                elif target_col_name.isalpha() and len(target_col_name) <= 3:
                    c_idx = column_index_from_string(target_col_name)
                    rel_idx = c_idx - min_col + 1
                    if 1 <= rel_idx <= num_cols:
                        target_col_rel = rel_idx

                if not target_col_rel:
                    continue

                abs_col_idx = min_col + target_col_rel - 1
                col_letter = get_column_letter(abs_col_idx)

                val_list = []
                row_val_map = {}
                for r_offset, r_data in enumerate(data_rows):
                    sheet_r = min_row + 1 + r_offset
                    v = r_data[target_col_rel - 1] if target_col_rel - 1 < len(r_data) else None
                    if v is not None and str(v).strip() != "":
                        try:
                            fv = float(str(v).replace("$", "").replace(" ", "").replace(",", ""))
                            val_list.append(fv)
                            row_val_map[sheet_r] = fv
                        except Exception:
                            pass

                if len(val_list) < 4:
                    outliers_results.append({
                        "column": headers[target_col_rel - 1],
                        "method": method,
                        "total_values_evaluated": len(val_list),
                        "outliers_count": 0,
                        "thresholds": {},
                        "flagged_cells": [],
                        "outliers": []
                    })
                    continue

                if method == "zscore":
                    mean = statistics.mean(val_list)
                    std = statistics.stdev(val_list) if len(val_list) > 1 else 0
                    z_thresh = 2.5 if len(val_list) < 30 else 3.0
                    lower_bound = mean - z_thresh * std
                    upper_bound = mean + z_thresh * std
                    thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "mean": mean, "std": std, "z_threshold": z_thresh}
                else:
                    sorted_vals = sorted(val_list)
                    q1, _, q3 = statistics.quantiles(sorted_vals, n=4, method="inclusive")
                    iqr = q3 - q1
                    lower_bound = q1 - 1.5 * iqr
                    upper_bound = q3 + 1.5 * iqr
                    thresholds = {"lower_bound": lower_bound, "upper_bound": upper_bound, "q1": q1, "q3": q3, "iqr": iqr}

                col_outliers = []
                for sheet_r, fv in row_val_map.items():
                    if fv < lower_bound or fv > upper_bound:
                        coord = f"{col_letter}{sheet_r}"
                        reason = f"Valor {fv} fuera de límites [{lower_bound:.2f}, {upper_bound:.2f}]"
                        col_outliers.append({
                            "cell": coord,
                            "row": sheet_r,
                            "value": fv,
                            "reason": reason
                        })
                        all_flagged_cells.append(coord)
                        outlier_cells_to_format.append((coord, sheet_r, abs_col_idx, fv, reason, add_comment, method))

                outliers_results.append({
                    "column": headers[target_col_rel - 1],
                    "method": method,
                    "total_values_evaluated": len(val_list),
                    "outliers_count": len(col_outliers),
                    "thresholds": thresholds,
                    "flagged_cells": [o["cell"] for o in col_outliers],
                    "outliers": col_outliers
                })

        # 5. Volcado a celdas OpenPyXL
        new_data_rows_count = len(data_rows)
        if clean_spec and clean_spec.get("trim_headers", True):
            for c_idx in range(1, num_cols + 1):
                ws.cell(row=min_row, column=min_col + c_idx - 1).value = headers[c_idx - 1]

        for r_offset, r_data in enumerate(data_rows):
            curr_r = min_row + 1 + r_offset
            for c_idx in range(1, num_cols + 1):
                cell = ws.cell(row=curr_r, column=min_col + c_idx - 1)
                val = r_data[c_idx - 1] if c_idx - 1 < len(r_data) else None
                cell.value = val
                if c_idx in col_date_cfg:
                    fmt = col_date_cfg[c_idx].get("format", "yyyy-mm-dd") if isinstance(col_date_cfg[c_idx], dict) else "yyyy-mm-dd"
                    cell.number_format = fmt

        # Eliminar filas sobrantes si hubo duplicados
        if duplicates_count > 0:
            leftover_start = min_row + 1 + new_data_rows_count
            ws.delete_rows(leftover_start, duplicates_count)
            if tbl:
                new_max_row = max_row - duplicates_count
                tbl.ref = f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{new_max_row}"

        # Formato visual de outliers
        if outlier_cells_to_format:
            warning_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
            warning_font = Font(color="C00000", bold=True)
            from openpyxl.comments import Comment
            for coord, sheet_r, abs_col_idx, fv, reason, add_com, m_name in outlier_cells_to_format:
                cell = ws[coord]
                cell.fill = warning_fill
                cell.font = warning_font
                if add_com:
                    try:
                        cell.comment = Comment(f"Outlier detectado ({m_name.upper()}): {reason}", "Antigravity")
                    except Exception:
                        pass

        return {
            "status": "success",
            "sheet": ws.title,
            "target_range_or_table": target_range_or_table or (tbl.name if tbl else f"{get_column_letter(min_col)}{min_row}:{get_column_letter(max_col)}{max_row}"),
            "original_rows": orig_data_rows_count,
            "remaining_rows": len(data_rows),
            "duplicates_removed": duplicates_count,
            "cells_cleaned": cells_modified,
            "rules_applied": list(clean_spec.keys()) if clean_spec else [],
            "outliers_count": len(all_flagged_cells),
            "outliers_flagged": all_flagged_cells,
            "outliers_summary": outliers_results
        }

    def close(self):
        if self.wb:
            self.wb.close()


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

    def __getattr__(self, name):
        return getattr(self.backend, name)
