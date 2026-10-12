# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias para los Backends de Antigravity Excel Engine:
BaseExcelBackend, HeadlessOpenPyXLBackend, LiveExcelCOMBackend y utilidades COM/retry.
"""

import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.backends import (
    BaseExcelBackend,
    LiveExcelCOMBackend,
    HeadlessOpenPyXLBackend,
    retry_on_excel_busy,
    excel_fast_mode,
    get_com_clean_address,
)
import antigravity_excel_core as legacy_core


class TestBaseExcelBackendContract(unittest.TestCase):
    """Verifica que BaseExcelBackend define el contrato abstracto y métodos comunes."""

    def test_abstract_class_cannot_be_instantiated(self):
        with self.assertRaises(TypeError):
            BaseExcelBackend()

    def test_calculate_placement_coordinates(self):
        # Auto con <= 8 columnas (derecha)
        coords_right = BaseExcelBackend._calculate_placement_coordinates(
            start_col=1, start_row=1, num_cols=4, num_rows=10, placement="auto"
        )
        self.assertEqual(coords_right["placement"], "right")
        self.assertEqual(coords_right["col"], 7)  # 1 + 4 + 2 = 7 ('G')
        self.assertEqual(coords_right["row"], 1)
        self.assertEqual(coords_right["cell"], "G1")

        # Auto con > 8 columnas (abajo)
        coords_bottom = BaseExcelBackend._calculate_placement_coordinates(
            start_col=1, start_row=1, num_cols=10, num_rows=5, placement="auto"
        )
        self.assertEqual(coords_bottom["placement"], "bottom")
        self.assertEqual(coords_bottom["col"], 1)
        self.assertEqual(coords_bottom["row"], 8)  # 1 + 5 + 2 = 8
        self.assertEqual(coords_bottom["cell"], "A8")

        # Anchor customizado
        coords_custom = BaseExcelBackend._calculate_placement_coordinates(
            start_col=1, start_row=1, num_cols=4, num_rows=5, anchor="J15"
        )
        self.assertEqual(coords_custom["placement"], "custom")
        self.assertEqual(coords_custom["col"], 10)
        self.assertEqual(coords_custom["row"], 15)
        self.assertEqual(coords_custom["cell"], "J15")


class TestHeadlessBackend(unittest.TestCase):
    """Verifica el comportamiento de HeadlessOpenPyXLBackend y sus métodos comunes."""

    def setUp(self):
        self.test_file = os.path.abspath("test_headless_backend.xlsx")
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass
        self.backend = HeadlessOpenPyXLBackend(file_path=self.test_file)

    def tearDown(self):
        self.backend.close()
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    def test_metadata_and_recalc_support(self):
        self.assertFalse(self.backend.supports_formula_recalc())
        self.assertFalse(self.backend.is_busy())
        st = self.backend.status()
        self.assertEqual(st["backend"], "headless_openpyxl")
        self.assertIn("active_sheet", st)

    def test_cell_value_read_write(self):
        self.backend.set_cell_value("B2", "Antigravity")
        self.backend.set_cell_value("B3", 42)
        self.backend.save()

        val_b2 = self.backend.get_cell_value("B2")
        val_b3 = self.backend.get_cell_value("B3")
        self.assertEqual(val_b2, "Antigravity")
        self.assertEqual(val_b3, 42)

    def test_write_and_read_range(self):
        data = [
            ["ID", "Nombre", "Activo"],
            [101, "Nodo Alpha", True],
            [102, "Nodo Beta", False],
        ]
        res = self.backend.write_range("A1", data)
        self.assertEqual(res["status"], "success")

        read_res = self.backend.read_range("A1:C3")
        vals = read_res["values"]
        self.assertEqual(vals[0], ["ID", "Nombre", "Activo"])
        self.assertEqual(vals[1][1], "Nodo Alpha")
        self.assertEqual(vals[2][0], 102)

    def test_get_sheet_and_open_workbook(self):
        sheet = self.backend.get_sheet()
        self.assertIsNotNone(sheet)
        self.backend.set_cell_value("A1", "Título Test")
        self.backend.save()

        # Reabrir con open_workbook
        wb = self.backend.open_workbook(self.test_file)
        self.assertIsNotNone(wb)
        self.assertEqual(self.backend.get_cell_value("A1"), "Título Test")


class TestLiveCOMBackendUnit(unittest.TestCase):
    """Verifica utilidades y mocks para LiveExcelCOMBackend sin requerir Excel activo."""

    def test_supports_formula_recalc_method(self):
        mock_backend = object.__new__(LiveExcelCOMBackend)
        self.assertTrue(mock_backend.supports_formula_recalc())

    def test_is_busy_checks_status(self):
        mock_backend = object.__new__(LiveExcelCOMBackend)
        mock_backend.status = MagicMock(return_value={"ready": True})
        self.assertFalse(mock_backend.is_busy())

        mock_backend.status = MagicMock(return_value={"ready": False})
        self.assertTrue(mock_backend.is_busy())

    def test_retry_on_excel_busy_decorator(self):
        attempts = 0

        @retry_on_excel_busy(max_retries=3, delay_sec=0.01)
        def flaky_func():
            nonlocal attempts
            attempts += 1
            if attempts < 3:
                import pywintypes
                # Simular RPC_E_SERVERCALL_RETRYLATER
                err = pywintypes.com_error(-2147417846, "RPC_E_SERVERCALL_RETRYLATER", None, -1)
                raise err
            return "SUCCESS"

        result = flaky_func()
        self.assertEqual(result, "SUCCESS")
        self.assertEqual(attempts, 3)

    def test_get_com_clean_address(self):
        class MockRange:
            Address = "$A$1:$B$10"

        class MockCellCallable:
            def Address(self, r_abs, c_abs):
                return "$C$5"

        self.assertEqual(get_com_clean_address(MockRange()), "A1:B10")
        self.assertEqual(get_com_clean_address(MockCellCallable()), "C5")

    def test_excel_fast_mode_context_manager(self):
        mock_app = MagicMock()
        mock_app.ScreenUpdating = True
        mock_app.DisplayAlerts = True
        mock_app.EnableEvents = True
        mock_app.Calculation = -4105

        with excel_fast_mode(mock_app):
            self.assertFalse(mock_app.ScreenUpdating)
            self.assertFalse(mock_app.DisplayAlerts)
            self.assertFalse(mock_app.EnableEvents)
            self.assertEqual(mock_app.Calculation, -4135)

        self.assertTrue(mock_app.ScreenUpdating)
        self.assertTrue(mock_app.DisplayAlerts)
        self.assertTrue(mock_app.EnableEvents)
        self.assertEqual(mock_app.Calculation, -4105)
        mock_app.Calculate.assert_called_once()


class TestLegacyReexportCompatibility(unittest.TestCase):
    """Verifica que antigravity_excel_core continúe re-exportando todos los símbolos."""

    def test_reexports_exist_and_match(self):
        self.assertIs(legacy_core.BaseExcelBackend, BaseExcelBackend)
        self.assertIs(legacy_core.LiveExcelCOMBackend, LiveExcelCOMBackend)
        self.assertIs(legacy_core.HeadlessOpenPyXLBackend, HeadlessOpenPyXLBackend)
        self.assertIs(legacy_core.retry_on_excel_busy, retry_on_excel_busy)
        self.assertIs(legacy_core.excel_fast_mode, excel_fast_mode)
        self.assertIs(legacy_core.get_com_clean_address, get_com_clean_address)


if __name__ == "__main__":
    unittest.main()
