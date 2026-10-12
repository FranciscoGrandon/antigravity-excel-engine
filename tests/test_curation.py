# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias enfocadas en el pipeline de curación de datos
(clean, dedup, outliers y curate_pipeline) usando HeadlessOpenPyXLBackend
y AntigravityExcelEngine.
"""

import os
import sys
import unittest
from openpyxl.worksheet.table import Table

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.backends.headless import HeadlessOpenPyXLBackend
from antigravity_excel_engine.engine import AntigravityExcelEngine


class TestCurationPipeline(unittest.TestCase):
    def setUp(self):
        self.test_file = os.path.abspath("test_curation_tmp.xlsx")
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass
        self.engine = AntigravityExcelEngine(mode="file", file_path=self.test_file)
        self.backend = self.engine.backend

    def tearDown(self):
        self.engine.close()
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    # ─────────────────────────── Deduplicación ──────────────────────────────

    def test_remove_duplicates_headless_backend(self):
        """Verifica deduplicación directa en HeadlessOpenPyXLBackend con tabla."""
        ws = self.backend._get_sheet(None)
        data = [
            ["ID", "Nombre", "Ciudad"],
            [1, "Ana", "Santiago"],
            [2, "Carlos", "Concepción"],
            [1, "Ana", "Santiago"],  # Duplicado exacto
            [3, "Beatriz", "Valparaíso"],
        ]
        self.backend.set_cell_range("A1", data)
        tab = Table(displayName="TblDedup", ref="A1:C5")
        ws.add_table(tab)

        res = self.backend.remove_duplicates(range_or_table="TblDedup")
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["original_rows"], 4)
        self.assertEqual(res["duplicates_removed"], 1)
        self.assertEqual(res["remaining_rows"], 3)
        self.assertEqual(tab.ref, "A1:C4")

    def test_remove_duplicates_via_engine(self):
        """Verifica delegación de remove_duplicates a través de AntigravityExcelEngine."""
        data = [
            ["Codigo", "Producto"],
            ["P100", "Teclado"],
            ["P200", "Mouse"],
            ["P100", "Teclado"],  # Duplicado
        ]
        self.engine.set_cell_range("A1", data)
        res = self.engine.remove_duplicates(range_str="A1:B4")
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["duplicates_removed"], 1)
        self.assertEqual(res["remaining_rows"], 2)

    # ─────────────────────────── Limpieza (Clean) ───────────────────────────

    def test_clean_table_dataset_backend(self):
        """Verifica estandarización de texto, trim y casteo de tipos en backend."""
        data = [
            ["  Categoria  ", "Precio", "Fecha"],
            ["  electronica ", "1500.5", "2024-01-15"],
            [" ROPA ", "250", "15/02/2024"],
        ]
        self.backend.set_cell_range("A1", data)
        rules = {
            "trim_headers": True,
            "standardize_text": {
                "Categoria": {"trim": True, "casing": "title"}
            },
            "cast_types": {
                "Precio": "float"
            },
            "coerce_dates": ["Fecha"]
        }
        res = self.backend.clean_table_dataset(target_range_or_table="A1:C3", rules=rules)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["rows_processed"], 2)

        ws = self.backend._get_sheet(None)
        # Encabezado trimmed
        self.assertEqual(ws["A1"].value, "Categoria")
        # Texto estandarizado en Title case
        self.assertEqual(ws["A2"].value, "Electronica")
        self.assertEqual(ws["A3"].value, "Ropa")
        # Precio casteado a float
        self.assertIsInstance(ws["B2"].value, float)
        self.assertEqual(ws["B2"].value, 1500.5)

    def test_clean_table_dataset_via_engine(self):
        """Verifica delegación de clean_table_dataset a través de AntigravityExcelEngine."""
        data = [
            ["Cliente", "Monto"],
            ["  empresa sa  ", "5000"],
            ["particular", "1200"],
        ]
        self.engine.set_cell_range("A1", data)
        rules = {
            "standardize_text": {
                "Cliente": {"trim": True, "casing": "upper"}
            },
            "cast_types": {
                "Monto": "int"
            }
        }
        res = self.engine.clean_table_dataset(range_str="A1:B3", rules=rules)
        self.assertEqual(res["status"], "success")

        ws = self.backend._get_sheet(None)
        self.assertEqual(ws["A2"].value, "EMPRESA SA")
        self.assertIsInstance(ws["B2"].value, int)
        self.assertEqual(ws["B2"].value, 5000)

    # ─────────────────────────── Outliers ───────────────────────────────────

    def test_detect_and_flag_outliers_iqr_and_zscore(self):
        """Verifica detección de outliers IQR y Z-Score con formato visual no destructivo."""
        data = [
            ["ID", "Valor"],
            [1, 10], [2, 11], [3, 10], [4, 12], [5, 11],
            [6, 10], [7, 11], [8, 12], [9, 10], [10, 9999]  # Outlier
        ]
        self.backend.set_cell_range("A1", data)

        res = self.backend.detect_and_flag_outliers(column="Valor", method="iqr", range_or_table="A1:B11", add_comment=True)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["outliers_count"], 1)
        self.assertIn("B11", res["flagged_cells"])

        ws = self.backend._get_sheet(None)
        cell = ws["B11"]
        # Valor original intacto (no destructivo)
        self.assertEqual(cell.value, 9999)
        # Formato visual
        self.assertEqual(cell.fill.start_color.rgb, "00FCE4D6")
        self.assertTrue(cell.font.bold)

    # ─────────────────────────── Curate Pipeline ────────────────────────────

    def test_curate_pipeline_composite(self):
        """Verifica pipeline compuesto (dedup + clean + outliers) en una sola pasada."""
        ws = self.backend._get_sheet(None)
        data = [
            [" ID ", " Canal ", "Ingresos", "Fecha"],
            [" 1 ", "  web  ", "100", "2024-01-01"],
            [" 2 ", "retail", "105", "2024-02-01"],
            [" 1 ", "  web  ", "100", "2024-01-01"],  # Duplicado
            [" 3 ", "retail", "95", "2024-03-01"],
            [" 4 ", "web", "102", "2024-04-01"],
            [" 5 ", "retail", "98", "2024-05-01"],
            [" 6 ", "web", "101", "2024-06-01"],
            [" 7 ", "retail", "99", "2024-07-01"],
            [" 8 ", "web", "99999", "2024-08-01"],  # Outlier
        ]
        self.backend.set_cell_range("A1", data)
        tab = Table(displayName="TblCurate", ref="A1:D10")
        ws.add_table(tab)

        pipeline_spec = {
            "dedup": {"key_columns": ["ID", "Canal"]},
            "clean": {
                "trim_headers": True,
                "standardize_text": {
                    "Canal": {"trim": True, "casing": "upper"}
                },
                "cast_types": {
                    "ID": "int",
                    "Ingresos": "float"
                },
                "coerce_dates": ["Fecha"]
            },
            "outliers": [
                {"column": "Ingresos", "method": "iqr", "add_comment": True}
            ]
        }

        res = self.engine.curate_pipeline(target_range_or_table="TblCurate", pipeline_spec=pipeline_spec)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["original_rows"], 9)
        self.assertEqual(res["duplicates_removed"], 1)
        self.assertEqual(res["remaining_rows"], 8)
        self.assertEqual(res["outliers_count"], 1)
        self.assertEqual(tab.ref, "A1:D9")

        # Verificar encabezados limpios
        self.assertEqual(ws["A1"].value, "ID")
        self.assertEqual(ws["B1"].value, "Canal")

        # Verificar estandarización
        self.assertEqual(ws["B2"].value, "WEB")
        self.assertEqual(ws["B3"].value, "RETAIL")


if __name__ == "__main__":
    unittest.main()
