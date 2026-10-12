# Creado por Francisco Grandón Vergara
"""
Backend Live Excel COM de Antigravity Excel Engine.
Control en vivo de Microsoft Excel mediante Win32 COM, ROT y hook visual de escritorio (oleacc).
"""

import os
import re
import csv
import time
import datetime
import functools
import threading
import statistics
from contextlib import contextmanager
from typing import List, Dict, Any, Optional, Union

try:
    import win32com.client
    import pythoncom
    import pywintypes
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None

# Módulos secundarios de Windows
win32gui = None
win32service = None
ctypes = None
uuid = None
wintypes = None


def _get_win32gui():
    global win32gui
    if win32gui is None:
        try:
            import win32gui as _wgui
            win32gui = _wgui
        except ImportError:
            pass
    return win32gui


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
    prev_calc = -4105  # xlCalculationAutomatic

    try:
        prev_screen_updating = getattr(app, "ScreenUpdating", True)
        prev_display_alerts = getattr(app, "DisplayAlerts", True)
        prev_enable_events = getattr(app, "EnableEvents", True)
        prev_calc = getattr(app, "Calculation", -4105)

        app.ScreenUpdating = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.Calculation = -4135  # xlCalculationManual
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


def get_com_clean_address(excel_range_or_cell) -> str:
    """Obtiene la dirección en formato limpio (A1 o A1:B2) de un objeto Range o Cell de win32com."""
    try:
        addr = excel_range_or_cell.Address
        if callable(addr):
            return str(addr(False, False)).replace("$", "")
        return str(addr).replace("$", "")
    except Exception:
        return str(getattr(excel_range_or_cell, "Address", "")).replace("$", "")


from .base import BaseExcelBackend
from antigravity_excel_engine.utils.ranges import (
    coordinate_from_string,
    column_index_from_string,
    get_column_letter,
    range_boundaries,
    chunk_cell_addresses,
)
_chunk_cell_addresses = chunk_cell_addresses

from antigravity_excel_engine.utils.dates import parse_excel_date_serial
from antigravity_excel_engine.formatting import (
    FinancialStyleGuide,
    hex_to_rgb_int,
    apply_style_preset_com,
)
from antigravity_excel_engine.formulas import (
    sanitize_formula_compatibility,
    audit_com_formula_errors,
    propagate_formula_row,
)

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
        if os.environ.get("AGY_TEST_ISOLATED") == "1":
            return None

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
        return {"status": "success", "range": get_com_clean_address(target_rng), "rows": row_count, "cols": col_count}

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
        return audit_com_formula_errors(ws, range_str)

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
        
        start_row = rng.Row
        start_col = rng.Column
        rows_cnt = rng.Rows.Count
        cols_cnt = rng.Columns.Count

        # Lectura matricial masiva en 2 llamadas COM (valores y fórmulas)
        raw_vals = rng.Value
        raw_forms = rng.Formula

        # Normalizar a lista de listas 2D
        if rows_cnt == 1 and cols_cnt == 1:
            vals_matrix = [[raw_vals]]
            forms_matrix = [[raw_forms]]
        elif rows_cnt == 1:
            vals_matrix = [list(raw_vals)]
            forms_matrix = [list(raw_forms)]
        elif cols_cnt == 1:
            vals_matrix = [[r] for r in raw_vals]
            forms_matrix = [[r] for r in raw_forms]
        else:
            vals_matrix = [list(r) for r in raw_vals]
            forms_matrix = [list(r) for r in raw_forms]

        # Check first row (headers)
        for c_idx in range(cols_cnt):
            val = str(vals_matrix[0][c_idx] or "")
            col_letter = get_column_letter(start_col + c_idx)
            header_cell_addr = f"{col_letter}{start_row}"
            if val and (val.startswith(" ") or val.endswith(" ") or "  " in val):
                header_whitespace.append({
                    "cell": header_cell_addr,
                    "header": val,
                    "trimmed": val.strip()
                })

        # Check column types and text numbers
        max_row_check = min(rows_cnt, 1000)
        for c_idx in range(cols_cnt):
            col_letter = get_column_letter(start_col + c_idx)
            h_name = str(vals_matrix[0][c_idx] or f"Col_{c_idx + 1}")
            col_types = set()
            for r_idx in range(1, max_row_check):
                val = vals_matrix[r_idx][c_idx]
                form_val = str(forms_matrix[r_idx][c_idx] or "")
                if form_val.startswith("="):
                    total_formulas += 1
                if val is not None:
                    col_types.add(type(val).__name__)
                    if isinstance(val, str) and val.strip().replace(".", "", 1).replace("-", "", 1).isdigit():
                        if len(numbers_as_text) < 50:
                            cell_addr = f"{col_letter}{start_row + r_idx}"
                            numbers_as_text.append({
                                "cell": cell_addr,
                                "column": h_name,
                                "value": val
                            })
            if len(col_types) > 1:
                inconsistent_types.append({
                    "column_index": c_idx + 1,
                    "header": h_name,
                    "types_detected": list(col_types)
                })

        return {
            "sheet": ws.Name,
            "range_audited": get_com_clean_address(rng),
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
        apply_style_preset_com(rng, preset)

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

    def wait_until_ready(self, timeout_sec: float = 10.0, retry_interval: float = 0.25) -> Dict[str, Any]:
        """
        Espera hasta que la instancia interactiva de Excel esté lista y disponible
        (sin edición de celdas por el usuario y con cálculos completados).
        """
        start_time = time.time()
        deadline = start_time + timeout_sec
        busy_hresults = (-2147417846, -2147418111, -2146777998, 0x8001010A, 0x80010001, 0x800AC472)

        while time.time() < deadline:
            try:
                if not self.app:
                    break

                # 1. Comprobar si la aplicación responde y no está bloqueada por edición de celdas
                ready = getattr(self.app, "Ready", True)
                if not ready:
                    time.sleep(retry_interval)
                    continue

                # 2. Comprobar CalculationState (0 = xlDone)
                calc_state = getattr(self.app, "CalculationState", 0)
                if calc_state != 0:
                    time.sleep(retry_interval)
                    continue

                elapsed = round(time.time() - start_time, 4)
                return {
                    "ready": True,
                    "elapsed_sec": elapsed,
                    "calculation_state": "done"
                }
            except Exception as e:
                is_busy = False
                if pywintypes and isinstance(e, pywintypes.com_error):
                    hresult = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                    if hresult in busy_hresults:
                        is_busy = True
                elif pythoncom and isinstance(e, pythoncom.com_error):
                    is_busy = True

                if is_busy:
                    time.sleep(retry_interval)
                    continue
                time.sleep(retry_interval)

        elapsed = round(time.time() - start_time, 4)
        return {
            "ready": False,
            "timeout": True,
            "elapsed_sec": elapsed
        }

    def _resolve_target(self, target: str):
        """Resuelve el objeto COM destino a partir de alias o expresiones."""
        target_clean = (target or "app").strip()
        t_lower = target_clean.lower()
        if t_lower in ("app", "excel", "application"):
            return self.app
        if t_lower in ("workbook", "wb", "activeworkbook"):
            return self.wb
        if t_lower in ("activesheet", "sheet"):
            return getattr(self.app, "ActiveSheet", None) or getattr(self.wb, "ActiveSheet", None)
        if t_lower.startswith("sheet:") or t_lower.startswith("worksheet:"):
            s_name = target_clean.split(":", 1)[1].strip()
            return self.wb.Sheets(s_name)
        if t_lower.startswith("range:") or t_lower.startswith("cell:"):
            r_addr = target_clean.split(":", 1)[1].strip()
            s_obj = getattr(self.app, "ActiveSheet", None) or getattr(self.wb, "ActiveSheet", None)
            return s_obj.Range(r_addr)

        # Probar primero self.app.Evaluate(target_clean)
        try:
            eval_res = self.app.Evaluate(target_clean)
            if eval_res is not None:
                return eval_res
        except Exception:
            pass

        # Acceso encadenado de atributos seguro
        parts = [p.strip() for p in target_clean.split(".") if p.strip()]
        if not parts:
            return self.app

        current = self.app
        if parts[0].lower() in ("wb", "workbook"):
            current = self.wb
            parts = parts[1:]
        elif parts[0].lower() in ("app", "application"):
            current = self.app
            parts = parts[1:]

        for part in parts:
            m_call = re.match(r"^(\w+)\((.*)\)$", part)
            if m_call:
                p_name = m_call.group(1)
                raw_arg = m_call.group(2).strip().strip("'\"")
                sub_prop = getattr(current, p_name)
                p_val = int(raw_arg) if raw_arg.isdigit() else raw_arg
                current = sub_prop(p_val) if callable(sub_prop) else sub_prop[p_val]
            else:
                current = getattr(current, part)
        return current

    @staticmethod
    def _serialize_com_result(val: Any) -> Any:
        """Serializa de forma segura un valor o resultado COM a tipos nativos de Python."""
        if val is None or isinstance(val, (int, float, str, bool)):
            return val
        if isinstance(val, (datetime.date, datetime.datetime)):
            return val.isoformat()
        if isinstance(val, (list, tuple)):
            return [LiveExcelCOMBackend._serialize_com_result(x) for x in val]
        if isinstance(val, dict):
            return {str(k): LiveExcelCOMBackend._serialize_com_result(v) for k, v in val.items()}
        # Objetos COM intermedios (Range, Worksheet, etc.)
        try:
            if hasattr(val, "Value"):
                v = val.Value
                return LiveExcelCOMBackend._serialize_com_result(v)
        except Exception:
            pass
        try:
            if hasattr(val, "Address"):
                return str(val.Address)
        except Exception:
            pass
        try:
            if hasattr(val, "Name"):
                return str(val.Name)
        except Exception:
            pass
        return str(val)

    @retry_on_excel_busy()
    def dispatch_com(
        self,
        target: str,
        member: str,
        args: Optional[List[Any]] = None,
        kwargs: Optional[Dict[str, Any]] = None
    ) -> Dict[str, Any]:
        """
        Despacha dinámicamente métodos, propiedades o expresiones en el modelo de objetos COM de Excel.
        """
        if not self.app:
            return {
                "success": False,
                "error": "Excel Application COM no está inicializada",
                "target": target,
                "member": member
            }

        try:
            obj = self._resolve_target(target)
        except Exception as e_res:
            return {
                "success": False,
                "error": f"Error resolviendo target '{target}': {str(e_res)}",
                "target": target,
                "member": member
            }

        if obj is None:
            return {
                "success": False,
                "error": f"No se pudo resolver el target '{target}'",
                "target": target,
                "member": member
            }

        try:
            attr = getattr(obj, member)
        except AttributeError:
            attr = None

        raw_res = None
        try:
            if attr is None:
                if args is not None and len(args) == 1 and not kwargs:
                    setattr(obj, member, args[0])
                    raw_res = args[0]
                elif kwargs and "value" in kwargs:
                    setattr(obj, member, kwargs["value"])
                    raw_res = kwargs["value"]
                else:
                    return {
                        "success": False,
                        "error": f"El miembro '{member}' no existe en el target",
                        "target": target,
                        "member": member
                    }
            elif callable(attr):
                c_args = args or []
                c_kwargs = kwargs or {}
                raw_res = attr(*c_args, **c_kwargs)
            else:
                if args is not None and len(args) == 1 and not kwargs:
                    setattr(obj, member, args[0])
                    raw_res = args[0]
                elif kwargs and "value" in kwargs:
                    setattr(obj, member, kwargs["value"])
                    raw_res = kwargs["value"]
                else:
                    raw_res = attr
        except TypeError:
            try:
                raw_res = getattr(obj, member)
            except Exception as e_get:
                return {
                    "success": False,
                    "error": f"Error accediendo al miembro '{member}': {str(e_get)}",
                    "target": target,
                    "member": member
                }
        except Exception as e_exec:
            return {
                "success": False,
                "error": f"Error ejecutando '{member}': {str(e_exec)}",
                "target": target,
                "member": member
            }

        try:
            serialized = self._serialize_com_result(raw_res)
        finally:
            try:
                del raw_res
            except Exception:
                pass

        return {
            "success": True,
            "target": target,
            "member": member,
            "result": serialized
        }


    def open_workbook(self, file_path: str) -> Any:
        """Abre o enlaza un libro de trabajo en la instancia de Excel."""
        self.file_path = file_path
        self.wb = self._init_workbook(file_path)
        return self.wb

    def get_sheet(self, sheet_name: Optional[str] = None) -> Any:
        """Obtiene la hoja de trabajo activa o la especificada por nombre."""
        return self._get_sheet(sheet_name)

    def supports_formula_recalc(self) -> bool:
        """Indica si el backend soporta motor nativo de recálculo dinámico de fórmulas."""
        return True

    def is_busy(self) -> bool:
        """Verifica si la instancia de Excel está bloqueada o en edición."""
        st = self.status()
        return not st.get("ready", True)

