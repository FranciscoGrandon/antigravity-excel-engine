# Creado por Francisco Grandón Vergara
"""
Suite de Pruebas Automatizadas para Antigravity Excel Engine.
"""

import os
import unittest
import sys

# Agregar ruta padre
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from antigravity_excel_core import AntigravityExcelEngine


class TestAntigravityExcelEngine(unittest.TestCase):
    def setUp(self):
        self.test_file = os.path.abspath("test_unit.xlsx")
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        self.engine = AntigravityExcelEngine(mode="file", file_path=self.test_file)

    def tearDown(self):
        self.engine.close()
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    def test_set_cell_range_and_get_csv(self):
        data = [
            ["ID", "Nombre", "Valor", "IVA", "Total"],
            [1, "Servicio A", 1000, "=C2*0.19", "=C2+D2"],
            [2, "Servicio B", 2500, "=C3*0.19", "=C3+D3"]
        ]
        res = self.engine.set_cell_range("A1", data, autofit=True)
        self.assertEqual(res["status"], "success")
        self.engine.save()

        csv_str = self.engine.get_range_as_csv("A1:E3", delimiter=";")
        self.assertIn("Servicio A", csv_str)
        self.assertIn("=C2*0.19", csv_str)

    def test_set_cells_declarative(self):
        cells = {
            "A1": {"value": "Indicador", "cellStyles": {"fontWeight": "bold", "fontColor": "#FFFFFF", "backgroundColor": "#0E2E63"}},
            "B1": {"value": "Meta", "cellStyles": {"fontWeight": "bold"}},
            "A2": {"value": "Cumplimiento"},
            "B2": {"value": 0.95, "cellStyles": {"numberFormat": "0.0%"}}
        }
        res = self.engine.set_cells(cells, autofit=True)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["cells_updated"], 4)
        self.engine.save()

        ranges = self.engine.get_cell_ranges(["A1:B2"])
        values = ranges["A1:B2"]["values"]
        self.assertEqual(values[0][0], "Indicador")
        self.assertEqual(values[1][0], "Cumplimiento")
        self.assertEqual(values[1][1], 0.95)

    def test_copy_paste_range(self):
        data = [["A", 10], ["B", 20]]
        self.engine.set_cell_range("A1", data)
        res = self.engine.copy_paste_range("A1:B2", "D1", paste_type="values")
        self.assertEqual(res["status"], "success")
        self.engine.save()

        ranges = self.engine.get_cell_ranges(["D1:E2"])
        self.assertEqual(ranges["D1:E2"]["values"][0][0], "A")
        self.assertEqual(ranges["D1:E2"]["values"][1][1], 20)


if __name__ == "__main__":
    unittest.main()
