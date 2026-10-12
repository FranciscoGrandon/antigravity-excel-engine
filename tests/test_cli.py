# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias para el CLI desacoplado de Antigravity Excel Engine
(antigravity_excel_engine.cli).
Verifica build_parser, _build_payload, comandos status, get-csv, get-ranges,
fallback standalone y formateo JSON.
"""

import io
import json
import os
import sys
import unittest
from unittest.mock import MagicMock, patch

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.cli import (
    build_parser,
    _build_payload,
    _run_standalone,
    _output,
    load_json_or_file,
    main,
)


class TestAntigravityExcelCLI(unittest.TestCase):
    def setUp(self):
        self.parser = build_parser()
        self.test_file = os.path.abspath("test_cli_tmp.xlsx")

    def tearDown(self):
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass

    # ─────────────────────────── Parser Tests ───────────────────────────────

    def test_build_parser_options(self):
        """Verifica que build_parser expone flags globales y subcomandos."""
        args = self.parser.parse_args(["status", "--mode", "file", "--file", "sample.xlsx"])
        self.assertEqual(args.mode, "file")
        self.assertEqual(args.file, "sample.xlsx")
        self.assertEqual(args.command, "status")

    def test_parser_daemon_flags(self):
        """Verifica parsing de flags de daemon."""
        args_start = self.parser.parse_args(["--daemon-start"])
        self.assertTrue(args_start.daemon_start)

        args_status = self.parser.parse_args(["--daemon-status"])
        self.assertTrue(args_status.daemon_status)

        args_stop = self.parser.parse_args(["--daemon-stop"])
        self.assertTrue(args_stop.daemon_stop)

    def test_parser_get_csv_command(self):
        """Verifica parsing del comando get-csv."""
        args = self.parser.parse_args(["get-csv", "A1:D10", "--sheet", "Datos", "--delimiter", ","])
        self.assertEqual(args.command, "get-csv")
        self.assertEqual(args.range, "A1:D10")
        self.assertEqual(args.sheet, "Datos")
        self.assertEqual(args.delimiter, ",")

    def test_parser_get_ranges_command(self):
        """Verifica parsing del comando get-ranges."""
        args = self.parser.parse_args(["get-ranges", "A1:B2", "C1:D2", "--formulas"])
        self.assertEqual(args.command, "get-ranges")
        self.assertEqual(args.ranges, ["A1:B2", "C1:D2"])
        self.assertTrue(args.formulas)

    # ─────────────────────────── Payload Tests ──────────────────────────────

    def test_build_payload_status(self):
        """Verifica construcción de payload para comando status."""
        args = self.parser.parse_args(["status", "--mode", "file", "--file", "demo.xlsx"])
        payload = _build_payload(args)
        self.assertEqual(payload["mode"], "file")
        self.assertEqual(payload["file"], "demo.xlsx")

    def test_build_payload_get_csv(self):
        """Verifica construcción de payload para get-csv."""
        args = self.parser.parse_args(["get-csv", "A1:B5", "--file", "demo.xlsx", "--sheet", "Hoja1"])
        payload = _build_payload(args)
        self.assertEqual(payload["range"], "A1:B5")
        self.assertEqual(payload["sheet"], "Hoja1")
        self.assertEqual(payload["delimiter"], ";")

    def test_build_payload_get_ranges(self):
        """Verifica construcción de payload para get-ranges."""
        args = self.parser.parse_args(["get-ranges", "A1:B2", "C1:C10", "--file", "demo.xlsx", "--formulas"])
        payload = _build_payload(args)
        self.assertEqual(payload["ranges"], ["A1:B2", "C1:C10"])
        self.assertTrue(payload["formulas"])

    def test_build_payload_set_range_inline(self):
        """Verifica payload de set-range con JSON inline."""
        args = self.parser.parse_args([
            "set-range", "--start-cell", "A1", "--values-json", '[[1, "Hola"], [2, "Mundo"]]', "--autofit"
        ])
        payload = _build_payload(args)
        self.assertEqual(payload["start_cell"], "A1")
        self.assertEqual(payload["values"], [[1, "Hola"], [2, "Mundo"]])
        self.assertTrue(payload["autofit"])

    def test_build_payload_missing_values_raises(self):
        """Verifica que set-range sin valores levanta ValueError."""
        args = self.parser.parse_args(["set-range", "--start-cell", "A1"])
        with self.assertRaises(ValueError):
            _build_payload(args)

    # ─────────────────────────── Helper Tests ───────────────────────────────

    def test_load_json_or_file(self):
        """Verifica carga de JSON directo o vacío."""
        self.assertIsNone(load_json_or_file(""))
        data = load_json_or_file('{"a": 1, "b": [2, 3]}')
        self.assertEqual(data["a"], 1)
        self.assertEqual(data["b"], [2, 3])

    def test_output_json_and_plain(self):
        """Verifica función _output con y sin flag JSON."""
        # Modo JSON con dict
        captured = io.StringIO()
        with patch("sys.stdout", captured):
            _output({"test": 123}, as_json=True)
        res = json.loads(captured.getvalue())
        self.assertEqual(res["test"], 123)

        # Modo texto plano con string
        captured_plain = io.StringIO()
        with patch("sys.stdout", captured_plain):
            _output("CSV_CONTENT_HERE", as_json=False)
        self.assertIn("CSV_CONTENT_HERE", captured_plain.getvalue())

    # ─────────────────────────── Standalone & Fallback ──────────────────────

    def test_run_standalone_status(self):
        """Verifica ejecución de status en modo standalone."""
        payload = {"mode": "file", "file": self.test_file}
        captured = io.StringIO()
        with patch("sys.stdout", captured):
            _run_standalone("status", payload, as_json=True)
        out = json.loads(captured.getvalue())
        self.assertIn("backend", out)
        self.assertEqual(out["backend"], "headless_openpyxl")

    def test_run_standalone_write_and_read_csv(self):
        """Verifica ciclo set-range y get-csv en standalone."""
        # 1. set-range
        payload_write = {
            "mode": "file",
            "file": self.test_file,
            "start_cell": "A1",
            "values": [["ID", "Nombre"], [1, "Test"]],
            "autofit": False,
        }
        captured_w = io.StringIO()
        with patch("sys.stdout", captured_w):
            _run_standalone("set-range", payload_write, as_json=True)
        resp_w = json.loads(captured_w.getvalue())
        self.assertEqual(resp_w["status"], "success")

        # 2. get-csv
        payload_csv = {
            "mode": "file",
            "file": self.test_file,
            "range": "A1:B2",
            "delimiter": ";",
        }
        captured_csv = io.StringIO()
        with patch("sys.stdout", captured_csv):
            _run_standalone("get-csv", payload_csv, as_json=False)
        csv_str = captured_csv.getvalue()
        self.assertIn("ID;Nombre", csv_str)
        self.assertIn("1;Test", csv_str)

    def test_main_dispatch_fallback(self):
        """Verifica que main() invoca standalone cuando el daemon está inactivo."""
        test_args = ["antigravity-excel", "status", "--mode", "file", "--file", self.test_file, "--json"]
        with patch("sys.argv", test_args):
            with patch("antigravity_excel_engine.cli._try_daemon", return_value=None):
                with patch("antigravity_excel_engine.cli._run_standalone") as mock_standalone:
                    main()
                    mock_standalone.assert_called_once()
                    self.assertEqual(mock_standalone.call_args[0][0], "status")


if __name__ == "__main__":
    unittest.main()
