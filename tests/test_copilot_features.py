# Creado por Francisco Grandón Vergara
"""
Pruebas unitarias y de integración para las funcionalidades Copilot en Antigravity Excel Engine:
1. AI Grounding & Table Schemas (inspect_table_schema, get_ai_grounding_context).
2. Gestión de Tablas de Excel ListObject (create_table, list_tables, resize_table).
3. Transpilador de Fórmulas Estructuradas (transpile_structured_formula).
4. Modo Dry-Run / Previsualización Especulativa en set_cells, set_cell_range, add_calculated_column.
5. Integración en AntigravityExcelEngine, Batch y CLI.
"""

import json
from unittest.mock import MagicMock

import pytest

from antigravity_excel_engine import (
    AntigravityExcelEngine,
    HeadlessOpenPyXLBackend,
    LiveExcelCOMBackend,
    STRUCTURED_REF_REGEX,
    create_table,
    get_ai_grounding_context,
    inspect_table_schema,
    list_tables,
    resize_table,
    transpile_structured_formula,
)
from antigravity_excel_engine.cli import build_parser, _build_payload


class TestStructuredFormulaTranspiler:
    """Pruebas del transpilador de fórmulas con referencias estructuradas."""

    def test_transpile_simple_columns(self):
        headers_map = {"Ventas": "B", "Costos": "C", "Precio": "D"}
        formula = "=[@Ventas] - [@Costos]"
        result = transpile_structured_formula(formula, headers_map, row_idx=2)
        assert result == "=B2 - C2"

        result_row5 = transpile_structured_formula(formula, headers_map, row_idx=5)
        assert result_row5 == "=B5 - C5"

    def test_transpile_columns_with_spaces_and_special_syntax(self):
        headers_map = {
            "Precio Unitario": "C",
            "Cantidad Vendida": "D",
            "Descuento": "E",
        }
        formula = "=[@[Precio Unitario]] * [@[Cantidad Vendida]] - [@Descuento]"
        result = transpile_structured_formula(formula, headers_map, row_idx=10)
        assert result == "=C10 * D10 - E10"

    def test_transpile_with_table_prefix(self):
        headers_map = {"Ventas": "B", "Costos": "C"}
        formula = "=TblVentas[@Ventas] - TblVentas[@[Costos]]"
        result = transpile_structured_formula(formula, headers_map, row_idx=4)
        assert result == "=B4 - C4"

    def test_transpile_case_insensitivity_and_whitespace(self):
        headers_map = {"ventas": "B", "costos": "C"}
        formula = "=[@ Ventas ] * 1.19"
        result = transpile_structured_formula(formula, headers_map, row_idx=3)
        assert result == "=B3 * 1.19"

    def test_transpile_unmapped_column_preserves_token(self):
        headers_map = {"Ventas": "B"}
        formula = "=[@Ventas] + [@Inexistente]"
        result = transpile_structured_formula(formula, headers_map, row_idx=2)
        assert result == "=B2 + [@Inexistente]"

    def test_transpile_empty_or_non_string_formula(self):
        assert transpile_structured_formula("", {"A": "A"}, 2) == ""
        assert transpile_structured_formula(None, {"A": "A"}, 2) is None


class TestTableManagementHeadless:
    """Pruebas de creación, listado y redimensionamiento de tablas en OpenPyXL Headless."""

    @pytest.fixture
    def populated_engine(self, tmp_path):
        f = tmp_path / "test_tables.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        headers = ["ID", "Producto", "Precio", "Stock"]
        rows = [
            [1, "Manzanas", 1200.5, 50],
            [2, "Peras", 1500.0, 30],
            [3, "Naranjas", 980.0, 80],
            [4, "Plátanos", 1100.0, 45],
        ]
        eng.set_cell_range("A1", [headers] + rows)
        eng.save()
        return eng

    def test_create_and_list_tables(self, populated_engine):
        res = populated_engine.create_table(
            range_str="A1:D5",
            table_name="TblInventario",
            style="TableStyleMedium9",
            has_headers=True,
        )
        assert res["status"] == "success"
        assert res["table_name"] == "TblInventario"
        assert res["range"] == "A1:D5"
        assert res["columns_count"] == 4

        # Listar tablas
        tables = populated_engine.list_tables()
        assert len(tables) >= 1
        found = next(t for t in tables if t["name"] == "TblInventario")
        assert found["range"] == "A1:D5"
        assert "Producto" in found["columns"]

    def test_create_table_autogenerates_name(self, populated_engine):
        res = populated_engine.create_table(range_str="A1:D5")
        assert res["status"] == "success"
        assert res["table_name"].startswith("Table")

    def test_resize_table(self, populated_engine):
        populated_engine.create_table(range_str="A1:D5", table_name="TblResize")
        # Agregar una columna en E
        populated_engine.set_cell_range("E1", [["Total"], [60025], [45000], [78400], [49500]])

        resize_res = populated_engine.resize_table("TblResize", "A1:E5")
        assert resize_res["status"] == "success"
        assert resize_res["new_range"] == "A1:E5"

        tables = populated_engine.list_tables()
        found = next(t for t in tables if t["name"] == "TblResize")
        assert found["range"] == "A1:E5"
        assert len(found["columns"]) == 5

    def test_resize_nonexistent_table_raises(self, populated_engine):
        with pytest.raises(ValueError, match="no encontrada"):
            populated_engine.resize_table("TablaFantasma", "A1:C5")


class TestAIGroundingAndSchemas:
    """Pruebas de inspección semántica de esquemas y contexto de AI Grounding."""

    @pytest.fixture
    def sample_data_engine(self, tmp_path):
        f = tmp_path / "grounding_test.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        headers = ["ID_Cliente", "Nombre", "Monto_Facturado", "Fecha_Lectura", "Activo", "Tarifa_Calculada"]
        data = [
            [101, "Empresa A", 154000.5, "2026-03-01", True, "=C2*1.19"],
            [102, "Empresa B", 89200.0, "2026-03-02", False, "=C3*1.19"],
            [103, "Empresa C", 215000.0, "2026-03-03", True, "=C4*1.19"],
        ]
        values = [headers] + [[r[0], r[1], r[2], r[3], r[4], None] for r in data]
        formulas = [headers] + [[None, None, None, None, None, r[5]] for r in data]
        eng.set_cell_range("A1", values=values, formulas=formulas)
        eng.create_table("A1:F4", table_name="TblClientes")
        eng.save()
        return eng

    def test_inspect_table_schema_by_name(self, sample_data_engine):
        schema = sample_data_engine.inspect_table_schema("TblClientes")
        assert schema["table_name"] == "TblClientes"
        assert schema["range"] == "A1:F4"
        assert schema["dimensions"]["rows"] == 4
        assert schema["dimensions"]["cols"] == 6
        assert schema["dimensions"]["data_rows"] == 3

        cols_by_name = {c["name"]: c for c in schema["columns"]}

        # ID_Cliente debe ser integer
        assert cols_by_name["ID_Cliente"]["inferred_type"] == "integer"
        assert cols_by_name["ID_Cliente"]["min"] == 101
        assert cols_by_name["ID_Cliente"]["max"] == 103

        # Monto_Facturado debe ser currency o float
        assert cols_by_name["Monto_Facturado"]["inferred_type"] in ("currency", "float")
        assert cols_by_name["Monto_Facturado"]["non_null_count"] == 3
        assert cols_by_name["Monto_Facturado"]["null_count"] == 0

        # Fecha_Lectura debe ser date
        assert cols_by_name["Fecha_Lectura"]["inferred_type"] == "date"

        # Activo debe ser boolean
        assert cols_by_name["Activo"]["inferred_type"] == "boolean"

        # Tarifa_Calculada debe tener fórmula
        assert cols_by_name["Tarifa_Calculada"]["has_formula"] is True

    def test_inspect_schema_by_free_range(self, sample_data_engine):
        schema = sample_data_engine.inspect_table_schema("A1:C4")
        assert schema["range"] == "A1:C4"
        assert len(schema["columns"]) == 3
        assert schema["columns"][0]["name"] == "ID_Cliente"

    def test_get_ai_grounding_context_json(self, sample_data_engine):
        ctx_json = sample_data_engine.get_ai_grounding_context("TblClientes", format="json")
        assert isinstance(ctx_json, str)
        parsed = json.loads(ctx_json)
        assert parsed["table_name"] == "TblClientes"
        assert len(parsed["columns"]) == 6

    def test_get_ai_grounding_context_markdown(self, sample_data_engine):
        ctx_md = sample_data_engine.get_ai_grounding_context("TblClientes", format="markdown")
        assert isinstance(ctx_md, str)
        assert "### Esquema de Tabla: `TblClientes`" in ctx_md
        assert "| Columna | Índice | Tipo |" in ctx_md
        assert "`ID_Cliente`" in ctx_md
        assert "`Monto_Facturado`" in ctx_md


class TestStructuredColumnInjectionInHeadless:
    """Pruebas de add_calculated_column usando fórmulas estructuradas de tabla."""

    def test_add_calculated_column_with_structured_syntax(self, tmp_path):
        f = tmp_path / "calc_test.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        headers = ["ID", "Ventas", "Costos"]
        data = [
            [1, 1000, 400],
            [2, 2000, 800],
            [3, 3000, 1500],
        ]
        eng.set_cell_range("A1", [headers] + data)
        eng.create_table("A1:C4", table_name="TblOperaciones")

        # Inyectar columna estructurada: =[@Ventas] - [@Costos]
        res = eng.add_calculated_column(
            header="Margen",
            formula_template="=[@Ventas] - [@Costos]",
            number_format="#,##0",
        )
        assert res["status"] == "success"
        assert res["column_index"] == 4

        # Verificar que en D2 esté =B2 - C2, en D3 esté =B3 - C3
        d_vals = eng.get_cell_ranges(["D1:D4"], include_formulas=True)["D1:D4"]
        assert d_vals["formulas"][0][0] in (None, "")  # Header
        assert d_vals["formulas"][1][0] == "=B2 - C2"
        assert d_vals["formulas"][2][0] == "=B3 - C3"
        assert d_vals["formulas"][3][0] == "=B4 - C4"


class TestDryRunSpeculativeMode:
    """Pruebas del modo dry_run en set_cell_range, set_cells y add_calculated_column."""

    @pytest.fixture
    def clean_engine(self, tmp_path):
        f = tmp_path / "dry_run_test.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        eng.set_cell_range("A1", [["A", "B"], [10, 20], [30, 40]])
        eng.save()
        return eng

    def test_set_cell_range_dry_run(self, clean_engine):
        # Simular escritura sobre A2:B3
        res = clean_engine.set_cell_range(
            start_cell="A2",
            values=[[99, 88], [77, 66]],
            dry_run=True,
        )
        assert res["status"] == "success"
        assert res["dry_run"] is True
        assert res["affected_cells_count"] == 4
        assert res["target_range"] == "A2:B3"
        assert len(res["sample_diff"]) == 4

        # Verificar primer diff
        diff0 = res["sample_diff"][0]
        assert diff0["cell"] == "A2"
        assert diff0["old_value"] == 10
        assert diff0["new_value"] == 99

        # Verificar que el libro original NO fue modificado
        actual_data = clean_engine.get_cell_ranges(["A2:B3"])["A2:B3"]["values"]
        assert actual_data == [[10, 20], [30, 40]]

    def test_set_cells_dry_run(self, clean_engine):
        cells_map = {
            "A2": {"value": 555},
            "B3": {"formula": "=SUM(A1:B1)"},
        }
        res = clean_engine.set_cells(cells_map, dry_run=True)
        assert res["status"] == "success"
        assert res["dry_run"] is True
        assert res["affected_cells_count"] == 2
        assert len(res["sample_diff"]) == 2

        # Verificar que A2 y B3 sigan con sus valores originales
        assert clean_engine.get_cell_value("A2") == 10
        assert clean_engine.get_cell_value("B3") == 40

    def test_add_calculated_column_dry_run(self, clean_engine):
        res = clean_engine.add_calculated_column(
            header="Total",
            formula_template="=A2+B2",
            dry_run=True,
        )
        assert res["status"] == "success"
        assert res["dry_run"] is True
        assert res["column_index"] == 3
        assert res["affected_cells_count"] == 3
        assert res["target_range"] == "C1:C3"
        assert len(res["sample_diff"]) >= 2
        assert res["sample_diff"][0]["new_value"] == "Total"

        # Verificar que no se creó la columna C en el libro
        c1_val = clean_engine.get_cell_value("C1")
        assert c1_val is None


class TestBatchAndCLICopilotIntegration:
    """Pruebas de la integración en execute_batch y parser de CLI."""

    def test_batch_copilot_operations(self, tmp_path):
        f = tmp_path / "batch_copilot.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        eng.set_cell_range("A1", [["ID", "Valor"], [1, 100], [2, 200]])
        eng.save()

        ops = [
            {"cmd": "create-table", "args": {"range": "A1:B3", "name": "TblBatch"}},
            {"cmd": "list-tables", "args": {}},
            {"cmd": "inspect-table", "args": {"table_or_range": "TblBatch"}},
        ]
        res = eng.batch(ops)
        assert res["status"] == "success"
        assert res["operations_executed"] == 3
        results = res["results"]
        assert results[0]["result"]["table_name"] == "TblBatch"
        assert len(results[1]["result"]) >= 1
        assert results[2]["result"]["table_name"] == "TblBatch"

    def test_cli_subparsers_for_copilot_features(self):
        parser = build_parser()

        # create-table parser
        args = parser.parse_args(["create-table", "--range", "A1:E20", "--name", "TblCLI", "--style", "TableStyleLight1"])
        assert args.command == "create-table"
        assert args.range == "A1:E20"
        assert args.table_name == "TblCLI"
        assert args.style == "TableStyleLight1"
        payload = _build_payload(args)
        assert payload["range"] == "A1:E20"
        assert payload["table_name"] == "TblCLI"

        # inspect-table parser
        args = parser.parse_args(["inspect-table", "--table", "TblCLI", "--format", "markdown"])
        assert args.command == "inspect-table"
        assert args.table_or_range == "TblCLI"
        assert args.format == "markdown"
        payload = _build_payload(args)
        assert payload["table_or_range"] == "TblCLI"
        assert payload["format"] == "markdown"

        # set-range with dry-run
        args = parser.parse_args(["set-range", "--start-cell", "A1", "--values-json", "[[1,2]]", "--dry-run"])
        payload = _build_payload(args)
        assert payload["dry_run"] is True


class TestLiveCOMBackendCopilotMocks:
    """Pruebas unitarias de LiveExcelCOMBackend para métodos de Copilot usando mocks."""

    def test_live_com_create_table_mock(self):
        backend = LiveExcelCOMBackend.__new__(LiveExcelCOMBackend)
        ws_mock = MagicMock()
        backend._get_sheet = MagicMock(return_value=ws_mock)
        tbl_mock = MagicMock()
        tbl_mock.Name = "TblMock"
        tbl_mock.ListColumns.Count = 3
        tbl_mock.ListRows.Count = 10
        tbl_mock.Range.Address = "$A$1:$C$11"
        ws_mock.ListObjects.Add.return_value = tbl_mock

        res = backend.create_table("A1:C11", table_name="TblMock")
        assert res["status"] == "success"
        assert res["table_name"] == "TblMock"
        ws_mock.ListObjects.Add.assert_called_once()

    def test_live_com_list_tables_mock(self):
        backend = LiveExcelCOMBackend.__new__(LiveExcelCOMBackend)
        ws_mock = MagicMock()
        ws_mock.Name = "Hoja1"
        backend._get_sheet = MagicMock(return_value=ws_mock)
        tbl1 = MagicMock()
        tbl1.Name = "T1"
        tbl1.Range.Address = "$A$1:$B$5"
        tbl1.ListColumns.Count = 2
        col1 = MagicMock(); col1.Name = "Col1"
        col2 = MagicMock(); col2.Name = "Col2"
        tbl1.ListColumns.side_effect = lambda idx: col1 if idx == 1 else col2
        ws_mock.ListObjects.Count = 1
        ws_mock.ListObjects.side_effect = lambda idx: tbl1

        tables = backend.list_tables(sheet="Hoja1")
        assert len(tables) == 1
        assert tables[0]["name"] == "T1"
        assert tables[0]["columns"] == ["Col1", "Col2"]

    def test_live_com_resize_table_mock(self):
        backend = LiveExcelCOMBackend.__new__(LiveExcelCOMBackend)
        ws_mock = MagicMock()
        ws_mock.Name = "Hoja1"
        backend._get_sheet = MagicMock(return_value=ws_mock)
        tbl = MagicMock()
        tbl.Name = "T1"
        ws_mock.ListObjects.side_effect = lambda name: tbl if name == "T1" else None

        res = backend.resize_table("T1", "A1:E20", sheet="Hoja1")
        assert res["status"] == "success"
        tbl.Resize.assert_called_once()


class TestPerformanceAndPrecisionOptimizations:
    """Pruebas para optimizaciones de performance, fórmulas estructuradas y precisión numérica."""

    def test_transpile_full_column_aggregation(self):
        headers = {"Ventas": "B", "Costos": "C", "Precio": "D"}
        formula = "=SUM([Ventas])"
        res = transpile_structured_formula(formula, headers, row_idx=2, row_bounds=(2, 10))
        assert res == "=SUM(B$2:B$10)"

        formula_table = "=AVERAGE(Tabla[Costos])"
        res_table = transpile_structured_formula(formula_table, headers, row_idx=2, table_range="A2:E50")
        assert res_table == "=AVERAGE(C$2:C$50)"

        formula_unbounded = "=SUM([Ventas])"
        assert transpile_structured_formula(formula_unbounded, headers, row_idx=2) == "=SUM(B:B)"

    def test_transpile_string_literals_with_brackets_protected(self):
        headers = {"Ventas": "B"}
        formula = '="Total [@Ventas]: " & [@Ventas]'
        res = transpile_structured_formula(formula, headers, row_idx=2)
        assert res == '="Total [@Ventas]: " & B2'

        formula2 = '="Valor [@X]: " & [@Ventas]'
        assert transpile_structured_formula(formula2, headers, row_idx=5) == '="Valor [@X]: " & B5'

        formula_quotes = '="Columna ""[Ventas]"" en tabla: " & [@Ventas]'
        assert transpile_structured_formula(formula_quotes, headers, row_idx=3) == '="Columna ""[Ventas]"" en tabla: " & B3'

    def test_transpile_formal_this_row_syntax(self):
        headers = {"Ventas": "B", "Costos": "C", "Precio Unitario": "D"}
        formula = "=Tabla[[#This Row],[Ventas]] + [[#This Row], [Costos]] + [@[Precio Unitario]]"
        res = transpile_structured_formula(formula, headers, row_idx=5)
        assert res == "=B5 + C5 + D5"

    def test_grounding_leading_zeros_and_ruts_inferred_as_text(self):
        from antigravity_excel_engine.operations.grounding import _infer_column_type
        # Códigos numéricos con cero inicial deben ser texto
        assert _infer_column_type(["00123", "00124", "00125"], "Codigo") == "text"
        assert _infer_column_type(["0987", "0988"], "Folio") == "text"
        # RUTs y números de teléfono con guiones
        assert _infer_column_type(["12345678-9", "9876543-K"], "RUT") == "text"
        assert _infer_column_type(["+56-9-12345678", "9-87654321"], "Telefono") == "text"
        # Números enteros y negativos válidos
        assert _infer_column_type(["10", "20", "30"], "Cantidad") == "integer"
        assert _infer_column_type(["-10", "-20", "-30"], "Diferencia") == "integer"

    def test_grounding_matrix_transposition_and_sampling(self, tmp_path):
        f = tmp_path / "large_grounding.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        headers = ["ID", "Codigo", "Monto"]
        # Generar 1200 filas de prueba
        rows = [[i, f"00{i}", float(i * 10)] for i in range(1, 1201)]
        eng.set_cell_range("A1", [headers] + rows)
        eng.save()

        schema = eng.inspect_table_schema("A1:C1201", max_sample_rows=300)
        assert schema["dimensions"]["data_rows"] == 1200
        assert schema["dimensions"]["cols"] == 3
        cols = {c["name"]: c for c in schema["columns"]}
        assert cols["ID"]["inferred_type"] == "integer"
        assert cols["Codigo"]["inferred_type"] == "text"
        assert cols["Monto"]["inferred_type"] in ("float", "currency")
        eng.close()

    def test_aggregate_math_fsum_precision(self, tmp_path):
        from antigravity_excel_engine.operations.aggregate import compute_aggregation_metric
        # Verificación directa de math.fsum vs sum convencional en flotantes
        values = [0.1] * 10
        total = compute_aggregation_metric(values, "sum")
        assert total == 1.0

        # Con valores None, vacíos y cadenas numéricas
        dirty_values = [10.5, None, "", "20.5", "invalid", 5.0]
        assert compute_aggregation_metric(dirty_values, "sum") == 36.0
        assert compute_aggregation_metric(dirty_values, "avg") == 12.0

        # Verificación con backend headless
        f = tmp_path / "agg_precision.xlsx"
        eng = AntigravityExcelEngine(mode="file", file_path=str(f))
        data = [
            ["Categoria", "Valor"],
            ["A", 0.1],
            ["A", 0.2],
            ["A", 0.3],
            ["B", 10.0],
            ["B", 20.0],
        ]
        eng.set_cell_range("A1", data)
        eng.save()

        res = eng.aggregate(group_by_col="Categoria", metric_col="Valor", agg_func="sum")
        assert res["groups_count"] == 2
        group_a = next(g for g in res["ranking"] if g["group"] == "A")
        assert group_a["aggregated_value"] == 0.6
        eng.close()
