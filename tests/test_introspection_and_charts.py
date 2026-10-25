# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias para introspección de libros, análisis y extensión de gráficos,
operaciones batch y blindajes de red (Fases 1 y 2).
"""

import io
import json
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.cli import (
    _build_payload,
    _run_standalone,
    _try_daemon,
    build_parser,
    main,
)
from antigravity_excel_engine.engine import AntigravityExcelEngine
from antigravity_excel_engine.ipc.protocol import COMMANDS, VALID_COMMANDS
from antigravity_excel_engine.operations.charts import (
    extend_chart_series,
    extend_range_reference,
    inspect_charts,
    inspect_workbook,
    parse_series_formula,
    update_chart_series,
)


class TestSeriesFormulaAndRangeHelpers(unittest.TestCase):
    """Verifica el parsing de fórmulas de series y la extensión de rangos."""

    def test_parse_series_formula_standard(self):
        formula = '=SERIES("Ventas", Hoja1!$A$2:$A$10, Hoja1!$B$2:$B$10, 1)'
        res = parse_series_formula(formula)
        self.assertEqual(res["name_ref"], '"Ventas"')
        self.assertEqual(res["x_range"], "Hoja1!$A$2:$A$10")
        self.assertEqual(res["y_range"], "Hoja1!$B$2:$B$10")
        self.assertEqual(res["order"], "1")

    def test_parse_series_formula_quoted_sheet_with_commas(self):
        formula = "=SERIES('Hoja, Especial'!$B$1, 'Hoja, Especial'!$A$2:$A$10, 'Hoja, Especial'!$B$2:$B$10, 2)"
        res = parse_series_formula(formula)
        self.assertEqual(res["name_ref"], "'Hoja, Especial'!$B$1")
        self.assertEqual(res["x_range"], "'Hoja, Especial'!$A$2:$A$10")
        self.assertEqual(res["y_range"], "'Hoja, Especial'!$B$2:$B$10")
        self.assertEqual(res["order"], "2")

    def test_parse_series_formula_empty_name_or_x(self):
        f1 = "=SERIES(, Hoja1!$A$2:$A$10, Hoja1!$B$2:$B$10, 1)"
        res1 = parse_series_formula(f1)
        self.assertEqual(res1["name_ref"], "")
        self.assertEqual(res1["x_range"], "Hoja1!$A$2:$A$10")
        self.assertEqual(res1["y_range"], "Hoja1!$B$2:$B$10")

        f2 = '=SERIES("Total", , Hoja1!$B$2:$B$10, 1)'
        res2 = parse_series_formula(f2)
        self.assertEqual(res2["name_ref"], '"Total"')
        self.assertEqual(res2["x_range"], "")
        self.assertEqual(res2["y_range"], "Hoja1!$B$2:$B$10")

    def test_extend_range_reference_vertical_and_horizontal(self):
        # Extender hacia abajo (filas)
        r1 = "'Hoja 1'!$B$2:$B$10"
        ext1 = extend_range_reference(r1, to_row=25)
        self.assertEqual(ext1, "'Hoja 1'!$B$2:$B$25")

        # Extender hacia la derecha (columnas)
        r2 = "Sheet1!$B$2:$M$2"
        ext2 = extend_range_reference(r2, to_col="X")
        self.assertEqual(ext2, "Sheet1!$B$2:$X$2")

        # Extender celda individual
        r3 = "Sheet1!$A$2"
        ext3 = extend_range_reference(r3, to_row=15)
        self.assertEqual(ext3, "Sheet1!$A$2:$A$15")

        # Extender áreas múltiples
        r4 = "(Sheet1!$A$1:$A$5,Sheet1!$A$6:$A$10)"
        ext4 = extend_range_reference(r4, to_row=20)
        self.assertEqual(ext4, "(Sheet1!$A$1:$A$5,Sheet1!$A$6:$A$20)")


class TestIntrospectionAndChartsHeadless(unittest.TestCase):
    """Verifica introspección de libros y operaciones de series en backend Headless."""

    def setUp(self):
        self.test_file = os.path.abspath("test_introspection_headless.xlsx")
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass
        self.engine = AntigravityExcelEngine(mode="file", file_path=self.test_file)

        # Preparar datos base
        data = [
            ["Mes", "Ventas", "Costo"],
            ["Enero", 100, 40],
            ["Febrero", 150, 55],
            ["Marzo", 200, 70],
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.create_chart(
            source_range="A1:C4",
            chart_type="column_clustered",
            title="Desempeño Mensual",
            anchor="E2",
        )
        self.engine.save()

    def tearDown(self):
        self.engine.close()
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    def test_inspect_workbook_metadata_and_structure(self):
        info = self.engine.inspect_workbook()
        self.assertEqual(info["status"], "success")
        self.assertIn("active_workbook", info)
        self.assertIn("sheets", info)
        self.assertIn("sheets_detail", info)

        sheet_name = info["sheets"][0]
        sheet_info = info["sheets_detail"][sheet_name]
        self.assertEqual(sheet_info["last_row"], 4)
        self.assertEqual(sheet_info["headers"], ["Mes", "Ventas", "Costo"])
        self.assertGreaterEqual(len(sheet_info["charts"]), 1)

    def test_inspect_charts(self):
        res = self.engine.inspect_charts()
        self.assertEqual(res["status"], "success")
        self.assertGreaterEqual(res["charts_count"], 1)
        c0 = res["charts"][0]
        self.assertEqual(c0["title"], "Desempeño Mensual")
        self.assertGreaterEqual(c0["series_count"], 1)
        s0 = c0["series"][0]
        self.assertIn("name", s0)
        self.assertIn("formula", s0)
        self.assertIn("y_range", s0)

    def test_extend_chart_series(self):
        # Extender series hasta fila 10
        res = self.engine.extend_chart_series(to_row=10)
        self.assertEqual(res["status"], "success")
        self.assertGreaterEqual(res["updated_count"], 1)
        for upd in res["updated_series"]:
            self.assertIn("$10", upd["new_formula"])

    def test_update_chart_series(self):
        res = self.engine.update_chart_series(
            chart_name="Desempeño Mensual",
            series_index=1,
            name_ref="Ventas Actualizadas",
            y_range="'Sheet'!$B$2:$B$15",
        )
        self.assertEqual(res["status"], "success")
        self.assertIn("Ventas Actualizadas", res["new_formula"])
        self.assertIn("$B$15", res["new_formula"])


class TestBatchExecutionAndCLI(unittest.TestCase):
    """Verifica operaciones batch atómicas y nuevos subcomandos CLI."""

    def setUp(self):
        self.test_file = os.path.abspath("test_batch_cli.xlsx")
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass
        self.parser = build_parser()

    def tearDown(self):
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    def test_protocol_registered_commands(self):
        for cmd in ["inspect-workbook", "inspect-charts", "extend-chart-series", "update-chart-series", "batch"]:
            self.assertIn(cmd, COMMANDS)
            self.assertIn(cmd, VALID_COMMANDS)

    def test_engine_batch_execution(self):
        engine = AntigravityExcelEngine(mode="file", file_path=self.test_file)
        operations = [
            {
                "command": "set-range",
                "start_cell": "A1",
                "values": [["ID", "Precio", "Cantidad"], [1, 100, 2], [2, 250, 4]],
            },
            {
                "command": "add-column",
                "range": "A1:C3",
                "header": "Total",
                "formula": "=[@Precio]*[@Cantidad]",
            },
        ]
        res = engine.batch(operations)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["operations_executed"], 2)
        self.assertTrue(os.path.exists(res["saved_to"]))
        engine.close()

        # Reabrir y verificar datos escritos
        eng_read = AntigravityExcelEngine(mode="file", file_path=self.test_file)
        csv_data = eng_read.get_range_as_csv("A1:D3")
        self.assertIn("Total", csv_data)
        eng_read.close()

    def test_cli_subparsers_parsing(self):
        # inspect-workbook
        args_iw = self.parser.parse_args(["inspect-workbook", "--sheet", "Hoja1", "--no-charts"])
        self.assertEqual(args_iw.command, "inspect-workbook")
        self.assertEqual(args_iw.sheet, "Hoja1")
        self.assertTrue(args_iw.no_charts)

        # inspect-charts
        args_ic = self.parser.parse_args(["inspect-charts", "--sheet", "Reporte"])
        self.assertEqual(args_ic.command, "inspect-charts")
        self.assertEqual(args_ic.sheet, "Reporte")

        # extend-chart-series
        args_ec = self.parser.parse_args(["extend-chart-series", "--to-col", "X", "--to-row", "50", "--charts", "Chart1"])
        self.assertEqual(args_ec.command, "extend-chart-series")
        self.assertEqual(args_ec.to_col, "X")
        self.assertEqual(args_ec.to_row, 50)
        self.assertEqual(args_ec.charts, ["Chart1"])

        # update-chart-series
        args_uc = self.parser.parse_args([
            "update-chart-series", "--chart", "Ventas", "--series-index", "1",
            "--x-range", "A1:A10", "--y-range", "B1:B10"
        ])
        self.assertEqual(args_uc.command, "update-chart-series")
        self.assertEqual(args_uc.chart, "Ventas")
        self.assertEqual(args_uc.series_index, 1)

        # batch
        args_batch = self.parser.parse_args(["batch", "--operations-json", '[{"command": "status"}]'])
        self.assertEqual(args_batch.command, "batch")
        self.assertEqual(args_batch.operations_json, '[{"command": "status"}]')

    def test_cli_build_payload_batch_and_introspection(self):
        args = self.parser.parse_args([
            "batch", "--operations-json", '[{"command": "status"}]', "--file", "sample.xlsx"
        ])
        payload = _build_payload(args)
        self.assertEqual(payload["file"], "sample.xlsx")
        self.assertEqual(payload["operations"], [{"command": "status"}])

    def test_cli_unc_path_sets_60s_timeout_for_daemon(self):
        unc_file = r"\\servidor\compartido\reporte_anual.xlsx"
        test_args = ["antigravity-excel", "status", "--file", unc_file]

        with patch("sys.argv", test_args):
            with patch("antigravity_excel_engine.cli._try_daemon") as mock_try:
                mock_try.return_value = {"status": "success", "result": {"ok": True}}
                with patch("sys.exit", side_effect=SystemExit):
                    try:
                        main()
                    except SystemExit:
                        pass
                    mock_try.assert_called_once()
                    _, kwargs = mock_try.call_args
                    self.assertEqual(kwargs.get("timeout_ms"), 60000)


if __name__ == "__main__":
    unittest.main()
