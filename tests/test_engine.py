# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias para AntigravityExcelEngine (orquestador/fachada).
Verifica inicialización en modo auto/file, delegación a operations,
status, wait_until_ready, close, save y manejo de excepciones.
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.engine import AntigravityExcelEngine, ExcelEngine
from antigravity_excel_engine.backends.headless import HeadlessOpenPyXLBackend
from antigravity_excel_engine.backends.live_com import LiveExcelCOMBackend


class TestAntigravityExcelEngine(unittest.TestCase):
    def setUp(self):
        import uuid
        self.test_file = os.path.abspath(f"test_engine_{uuid.uuid4().hex[:8]}.xlsx")
        self.engine = AntigravityExcelEngine(mode="file", file_path=self.test_file)

    def tearDown(self):
        try:
            self.engine.close()
        except Exception:
            pass
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    def test_alias_excel_engine(self):
        """Verifica que ExcelEngine es un alias exacto de AntigravityExcelEngine."""
        self.assertIs(AntigravityExcelEngine, ExcelEngine)

    def test_init_file_mode(self):
        """Verifica inicialización explícita en modo file."""
        self.assertEqual(self.engine.mode, "file")
        self.assertIsInstance(self.engine.backend, HeadlessOpenPyXLBackend)
        self.assertEqual(self.engine.file_path, self.test_file)

    def test_init_invalid_mode_raises(self):
        """Verifica que un modo inválido levanta ValueError."""
        with self.assertRaises(ValueError):
            AntigravityExcelEngine(mode="unsupported_mode")

    def test_auto_detect_fallback_to_headless(self):
        """Verifica que en auto-detect sin Excel COM activo selecciona Headless."""
        with patch("win32com.client.GetActiveObject", side_effect=Exception("No COM")):
            with patch.object(LiveExcelCOMBackend, "_try_get_live_desktop_excel", return_value=None):
                engine_auto = AntigravityExcelEngine(mode="auto", file_path=self.test_file)
                self.assertIsInstance(engine_auto.backend, HeadlessOpenPyXLBackend)
                engine_auto.close()

    def test_auto_detect_with_active_com(self):
        """Verifica que en auto-detect con Excel COM activo selecciona Live COM."""
        mock_com_app = MagicMock()
        with patch("win32com.client.GetActiveObject", return_value=mock_com_app):
            with patch.object(LiveExcelCOMBackend, "_init_workbook", return_value=None):
                with patch.object(LiveExcelCOMBackend, "_init_app", return_value=mock_com_app):
                    engine_com = AntigravityExcelEngine(mode="auto", file_path=self.test_file)
                    self.assertIsInstance(engine_com.backend, LiveExcelCOMBackend)
                    engine_com.backend.wb = None
                    engine_com.backend.app = None

    def test_status_and_wait_until_ready(self):
        """Verifica status() y wait_until_ready() delegando al backend."""
        st = self.engine.status()
        self.assertIn("backend", st)
        self.assertEqual(st["backend"], "headless_openpyxl")
        self.assertEqual(self.engine.mode, "file")

        ready = self.engine.wait_until_ready(timeout_sec=1.0)
        self.assertTrue(ready.get("ready"))
        self.assertEqual(ready.get("backend"), "headless")

    def test_save_and_close(self):
        """Verifica ciclo save() y close()."""
        self.engine.set_cell_range("A1", [["Test", 123]])
        saved_path = self.engine.save()
        self.assertTrue(os.path.exists(saved_path))
        self.engine.close()

    def test_delegation_to_read_and_write(self):
        """Verifica delegación correcta a operations read y write."""
        data = [
            ["ID", "Nombre", "Monto"],
            [1, "Cliente A", 5000],
            [2, "Cliente B", 7500],
        ]
        res_write = self.engine.set_cell_range("A1", data, autofit=True)
        self.assertEqual(res_write["status"], "success")

        csv_out = self.engine.get_range_as_csv("A1:C3", delimiter=";")
        self.assertIn("Cliente A", csv_out)
        self.assertIn("5000", csv_out)

        ranges_out = self.engine.get_cell_ranges(["A1:B2"])
        self.assertIn("A1:B2", ranges_out)
        self.assertEqual(ranges_out["A1:B2"]["values"][0], ["ID", "Nombre"])

    def test_delegation_to_declarative_cells(self):
        """Verifica delegación a set_cells declarativo."""
        cells = {
            "A1": {"value": "Total"},
            "B1": {"value": 100},
        }
        res = self.engine.set_cells(cells)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["cells_updated"], 2)

    def test_delegation_to_audit_and_errors(self):
        """Verifica delegación a audit_data_quality y check_formula_errors."""
        self.engine.set_cell_range("A1", [["A", "B"], [1, 2]])
        report = self.engine.audit_data_quality("A1:B2")
        self.assertEqual(report["total_cells"], 4)
        self.assertIn("range_audited", report)

        errors = self.engine.check_formula_errors("A1:B2")
        self.assertIsInstance(errors, list)
        self.assertEqual(len(errors), 0)

    def test_delegation_to_aggregate_and_column(self):
        """Verifica delegación a aggregate y add_calculated_column."""
        data = [
            ["Region", "Ventas"],
            ["Norte", 100],
            ["Sur", 200],
            ["Norte", 150],
        ]
        self.engine.set_cell_range("A1", data)
        agg = self.engine.aggregate(group_by_col="Region", metric_col="Ventas", agg_func="sum")
        self.assertEqual(agg["groups_count"], 2)
        top = agg["ranking"][0]
        self.assertEqual(top["group"], "Norte")
        self.assertEqual(top["aggregated_value"], 250)

        col_res = self.engine.add_calculated_column(new_col_name="Impuesto", formula_template="=B2*0.19")
        self.assertEqual(col_res["status"], "success")

    def test_dispatch_com_headless_response(self):
        """Verifica que dispatch_com en backend headless retorna success False."""
        res = self.engine.dispatch_com(target="app", member="Version")
        self.assertFalse(res.get("success"))
        self.assertEqual(res.get("backend"), "headless")


if __name__ == "__main__":
    unittest.main()
