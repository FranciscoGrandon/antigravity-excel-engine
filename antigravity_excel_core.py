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
from abc import ABC, abstractmethod
from typing import List, Dict, Any, Optional, Union

try:
    import win32com.client
    import pythoncom
    import pywintypes
    import win32file
    import win32con
except ImportError:
    win32com = None
    pythoncom = None
    pywintypes = None
    win32file = None
    win32con = None

try:
    import openpyxl
    from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter, range_boundaries
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
    from openpyxl.comments import Comment
except ImportError:
    openpyxl = None


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

    @abstractmethod
    def aggregate(self, group_by_col: str, metric_col: str, agg_func: str = "sum", sheet: Optional[str] = None, top_n: int = 10, ascending: bool = False) -> Dict[str, Any]: pass
    
    @abstractmethod
    def save(self, file_path: Optional[str] = None) -> str: pass
    
    @abstractmethod
    def close(self): pass
    
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
        try:
            app = win32com.client.GetActiveObject("Excel.Application")
            app.Visible = True
            return app
        except Exception:
            app = win32com.client.Dispatch("Excel.Application")
            app.Visible = True
            return app

    @retry_on_excel_busy()
    def _init_workbook(self, file_path: Optional[str]):
        if file_path:
            abs_p = os.path.abspath(file_path)
            for w in self.app.Workbooks:
                if os.path.abspath(w.FullName).lower() == abs_p.lower():
                    return w
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
            return self.wb.Sheets(sheet_name)
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
    def save(self, file_path: Optional[str] = None) -> str:
        if file_path:
            abs_p = os.path.abspath(file_path)
            self.wb.SaveAs(abs_p)
            return abs_p
        self.wb.Save()
        return self.wb.FullName

    def close(self):
        pass


class HeadlessOpenPyXLBackend(BaseExcelBackend):
    def __init__(self, file_path: Optional[str] = None):
        if not openpyxl:
            raise ImportError("openpyxl no está instalado en este entorno.")
        self.file_path = file_path
        if file_path and os.path.exists(file_path):
            self.wb = self._load_workbook_resilient(file_path)
        else:
            self.wb = openpyxl.Workbook()

    def _load_workbook_resilient(self, file_path: str):
        """Carga openpyxl soportando lectura concurrente si el archivo está bloqueado por Excel."""
        try:
            return openpyxl.load_workbook(file_path, data_only=False)
        except PermissionError:
            if win32file and win32con:
                handle = win32file.CreateFile(
                    os.path.abspath(file_path),
                    win32file.GENERIC_READ,
                    win32file.FILE_SHARE_READ | win32file.FILE_SHARE_WRITE | win32file.FILE_SHARE_DELETE,
                    None,
                    win32file.OPEN_EXISTING,
                    win32file.FILE_ATTRIBUTE_NORMAL,
                    None
                )
                chunks = []
                while True:
                    hr, data = win32file.ReadFile(handle, 64 * 1024)
                    if not data:
                        break
                    chunks.append(data)
                win32file.CloseHandle(handle)
                buffer = io.BytesIO(b"".join(chunks))
                return openpyxl.load_workbook(buffer, data_only=False)
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
                cell.value = cell_info["formula"]
            elif "value" in cell_info:
                cell.value = cell_info["value"]

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

    def save(self, file_path: Optional[str] = None) -> str:
        out_path = file_path or self.file_path
        if out_path:
            self.wb.save(out_path)
            return os.path.abspath(out_path)
        raise ValueError("No se especificó ruta para guardar el archivo.")

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
            is_excel_active = False
            if win32com:
                try:
                    win32com.client.GetActiveObject("Excel.Application")
                    is_excel_active = True
                except Exception:
                    is_excel_active = False

            if is_excel_active:
                return LiveExcelCOMBackend(self.file_path)
            elif self.file_path and openpyxl:
                return HeadlessOpenPyXLBackend(self.file_path)
            elif win32com:
                return LiveExcelCOMBackend(self.file_path)
            elif openpyxl:
                return HeadlessOpenPyXLBackend(self.file_path)
            else:
                raise ImportError("No se encontró backend disponible (pywin32 u openpyxl).")
        else:
            raise ValueError(f"Modo no reconocido: '{self.mode}'. Use 'auto', 'live' o 'file'.")

    def __getattr__(self, name):
        return getattr(self.backend, name)
