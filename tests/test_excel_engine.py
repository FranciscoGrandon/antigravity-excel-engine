# Creado por Francisco Grandón Vergara
"""
Suite de Pruebas Automatizadas para Antigravity Excel Engine.
"""

import os
import unittest
import sys

import datetime

# Agregar ruta padre
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from antigravity_excel_core import AntigravityExcelEngine, FinancialStyleGuide, sanitize_formula_compatibility, parse_excel_date_serial


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

    def test_audit_data_quality(self):
        data = [
            ["ID", " Nombre ", "Year"],
            [1, "Cliente A", "2024"],
            [2, "Cliente B", "2025"]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        report = self.engine.audit_data_quality()
        self.assertEqual(len(report["header_whitespace_issues"]), 1)
        self.assertEqual(report["header_whitespace_issues"][0]["trimmed"], "Nombre")
        self.assertEqual(len(report["numbers_stored_as_text"]), 2)
        self.assertFalse(report["has_dynamic_formulas"])

    def test_aggregate(self):
        data = [
            ["Region", "Ventas"],
            ["Norte", 100],
            ["Sur", 200],
            ["Norte", 150],
            ["Sur", 50]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        res = self.engine.aggregate("Region", "Ventas", agg_func="sum")
        self.assertEqual(res["groups_count"], 2)
        top = res["ranking"][0]
        self.assertEqual(top["group"], "Norte")
        self.assertEqual(top["aggregated_value"], 250)
        self.assertEqual(top["max_cell"], "B4")

    def test_add_calculated_column(self):
        data = [
            ["Precio", "Cantidad"],
            [10, 2],
            [20, 3]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        res = self.engine.add_calculated_column("Total", "=A2*B2", number_format="$#,##0.00")
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["header"], "Total")
        self.assertEqual(res["column_index"], 3)
        self.engine.save()

        # Check values and formulas
        ranges = self.engine.get_cell_ranges(["C1:C3"], include_formulas=True)
        self.assertEqual(ranges["C1:C3"]["values"][0][0], "Total")
        self.assertEqual(ranges["C1:C3"]["formulas"][1][0], "=A2*B2")
        self.assertEqual(ranges["C1:C3"]["formulas"][2][0], "=A3*B3")

    def test_calculate_placement_coordinates(self):
        # 1. auto with cols <= 8 -> right (start_col + num_cols + 2)
        pos_right = self.engine.backend._calculate_placement_coordinates(start_col=1, start_row=3, num_cols=5, num_rows=10, placement="auto")
        self.assertEqual(pos_right["placement"], "right")
        self.assertEqual(pos_right["col"], 8) # 1 + 5 + 2 = 8 ('H')
        self.assertEqual(pos_right["row"], 3)
        self.assertEqual(pos_right["cell"], "H3")

        # 2. auto with cols > 8 -> bottom (start_row + num_rows + 2)
        pos_bottom = self.engine.backend._calculate_placement_coordinates(start_col=1, start_row=3, num_cols=12, num_rows=10, placement="auto")
        self.assertEqual(pos_bottom["placement"], "bottom")
        self.assertEqual(pos_bottom["col"], 1)
        self.assertEqual(pos_bottom["row"], 15) # 3 + 10 + 2 = 15
        self.assertEqual(pos_bottom["cell"], "A15")

        # 3. anchor specified
        pos_custom = self.engine.backend._calculate_placement_coordinates(start_col=1, start_row=1, num_cols=2, num_rows=2, anchor="I3")
        self.assertEqual(pos_custom["placement"], "custom")
        self.assertEqual(pos_custom["col"], 9)
        self.assertEqual(pos_custom["row"], 3)
        self.assertEqual(pos_custom["cell"], "I3")

    def test_create_pivot_table_headless(self):
        data = [
            ["Pais", "Segmento", "Ventas"],
            ["Chile", "Empresa", 100],
            ["Chile", "Gobierno", 200],
            ["Peru", "Empresa", 150],
            ["Peru", "Gobierno", 250]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        res = self.engine.create_pivot_table(
            source="A1:C5",
            rows=["Pais"],
            cols=["Segmento"],
            value_field="Ventas",
            value_func="sum",
            value_format="$#,##0.00",
            dest_sheet="Resumen_Pivot",
            dest_cell="A3"
        )
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["dest_sheet"], "Resumen_Pivot")
        self.assertEqual(res["grand_total"], 700.0)
        self.engine.save()

        # Verify cross-tabulation in new sheet
        ranges = self.engine.get_cell_ranges(["A3:D6"], sheet="Resumen_Pivot")
        matrix = ranges["A3:D6"]["values"]
        self.assertEqual(matrix[0][0], "Pais")
        self.assertEqual(matrix[0][3], "Total General")
        self.assertEqual(matrix[3][0], "Total General")
        self.assertEqual(matrix[3][3], 700.0)

    def test_create_chart_headless(self):
        data = [
            ["Categoria", "Valor"],
            ["A", 50],
            ["B", 80],
            ["C", 30]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        res = self.engine.create_chart(
            source="A1:B4",
            chart_type="column_clustered",
            title="Comparativa",
            placement="auto"
        )
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["placement"], "right")
        self.assertEqual(res["anchor_cell"], "E1") # 1 + 2 + 2 = 5 ('E')
        self.engine.save()

    def test_sanitize_formula_compatibility(self):
        # 1. Functions needing prefix
        self.assertEqual(
            sanitize_formula_compatibility('=TEXTJOIN(",", TRUE, A1:A5)'),
            '=_xlfn.TEXTJOIN(",", TRUE, A1:A5)'
        )
        self.assertEqual(
            sanitize_formula_compatibility("=MAXIFS(D2:D10, A2:A10, \">0\")"),
            "=_xlfn.MAXIFS(D2:D10, A2:A10, \">0\")"
        )
        self.assertEqual(
            sanitize_formula_compatibility("=CONCAT(A1, B1) + MINIFS(C1:C10, D1:D10, 1)"),
            "=_xlfn.CONCAT(A1, B1) + _xlfn.MINIFS(C1:C10, D1:D10, 1)"
        )
        self.assertEqual(
            sanitize_formula_compatibility("=IFS(A1>10, 1, A1>5, 2) + SWITCH(B1, 1, \"A\", \"B\")"),
            "=_xlfn.IFS(A1>10, 1, A1>5, 2) + _xlfn.SWITCH(B1, 1, \"A\", \"B\")"
        )
        # 2. Already prefixed: should not duplicate
        self.assertEqual(
            sanitize_formula_compatibility('=_xlfn.TEXTJOIN(",", TRUE, A1:A5)'),
            '=_xlfn.TEXTJOIN(",", TRUE, A1:A5)'
        )
        # 3. Standard formulas without modern functions should remain unchanged
        self.assertEqual(
            sanitize_formula_compatibility("=SUM(A1:A10) + AVERAGE(B1:B10) * 1.19"),
            "=SUM(A1:A10) + AVERAGE(B1:B10) * 1.19"
        )

    def test_financial_style_guide_defaults(self):
        # Check constants
        self.assertEqual(FinancialStyleGuide.CURRENCY_FORMAT, "$ #,##0.00;($ #,##0.00);\"-\"")
        self.assertEqual(FinancialStyleGuide.PERCENT_FORMAT, "0.0%")
        self.assertEqual(FinancialStyleGuide.COLOR_INPUT, "#1F4E79")
        self.assertEqual(FinancialStyleGuide.COLOR_FORMULA, "#000000")
        self.assertEqual(FinancialStyleGuide.COLOR_ASSUMPTION_FILL, "#FFF2CC")

        # Test add_calculated_column with default currency format and modern formula
        data = [
            ["Item", "Monto"],
            ["Venta 1", 100],
            ["Venta 2", -50],
            ["Venta 3", 0]
        ]
        self.engine.set_cell_range("A1", data)
        self.engine.save()

        # Add calculated column without explicit format: should default to FinancialStyleGuide.CURRENCY_FORMAT
        res_col = self.engine.add_calculated_column("Maximo", '=MAXIFS(B2:B4, A2:A4, "<>0")')
        self.assertEqual(res_col["status"], "success")
        self.assertIn("_xlfn.MAXIFS", res_col["formula_applied"])
        self.engine.save()

        # Verify cell number format in openpyxl
        ws = self.engine.backend._get_sheet(None)
        self.assertEqual(ws["C2"].number_format, FinancialStyleGuide.CURRENCY_FORMAT)
        self.assertEqual(ws["C2"].value, '=_xlfn.MAXIFS(B2:B4, A2:A4, "<>0")')

        # Test create_pivot_table with default currency format
        res_pivot = self.engine.create_pivot_table(
            source="A1:B4",
            rows=["Item"],
            value_field="Monto",
            value_func="sum",
            dest_sheet="Pivot_Financiero"
        )
        self.assertEqual(res_pivot["status"], "success")
        self.assertEqual(res_pivot["value_format"], FinancialStyleGuide.CURRENCY_FORMAT)
        self.engine.save()

        # Verify number format in pivot sheet
        ws_pivot = self.engine.backend._get_sheet("Pivot_Financiero")
        self.assertEqual(ws_pivot["B4"].number_format, FinancialStyleGuide.CURRENCY_FORMAT)

    def test_apply_style_preset_headless(self):
        data = [["Header", "Input", "Total", "Currency", "Percent", "DeltaPos", "DeltaNeg"]]
        self.engine.set_cell_range("A1", data)
        
        self.engine.apply_style_preset("A1", "header")
        self.engine.apply_style_preset("B1", "input_cell")
        self.engine.apply_style_preset("C1", "total_row")
        self.engine.apply_style_preset("D1", "currency")
        self.engine.apply_style_preset("E1", "percent")
        self.engine.apply_style_preset("F1", "delta_positive")
        self.engine.apply_style_preset("G1", "delta_negative")
        self.engine.save()

        ws = self.engine.backend._get_sheet(None)
        
        # 1. header: fondo azul #0E2E63, texto blanco, bold, centrado vertical
        self.assertTrue(str(ws["A1"].fill.start_color.rgb).endswith("0E2E63"))
        self.assertTrue(ws["A1"].font.bold)
        self.assertTrue(str(ws["A1"].font.color.rgb).endswith("FFFFFF"))
        self.assertEqual(ws["A1"].alignment.vertical, "center")

        # 2. input_cell: fondo amarillo #FFF2CC, texto azul #1F4E79, borde medio, centrado horizontal
        self.assertTrue(str(ws["B1"].fill.start_color.rgb).endswith("FFF2CC"))
        self.assertTrue(str(ws["B1"].font.color.rgb).endswith("1F4E79"))
        self.assertEqual(ws["B1"].border.left.style, "medium")
        self.assertEqual(ws["B1"].alignment.horizontal, "center")

        # 3. total_row: bold, borde superior simple, borde inferior doble
        self.assertTrue(ws["C1"].font.bold)
        self.assertEqual(ws["C1"].border.top.style, "thin")
        self.assertEqual(ws["C1"].border.bottom.style, "double")

        # 4. currency: FinancialStyleGuide.CURRENCY_FORMAT
        self.assertEqual(ws["D1"].number_format, FinancialStyleGuide.CURRENCY_FORMAT)

        # 5. percent: FinancialStyleGuide.PERCENT_FORMAT
        self.assertEqual(ws["E1"].number_format, FinancialStyleGuide.PERCENT_FORMAT)

        # 6. delta_positive: verde #008000, bold
        self.assertTrue(ws["F1"].font.bold)
        self.assertTrue(str(ws["F1"].font.color.rgb).endswith("008000"))

        # 7. delta_negative: rojo #C00000, bold
        self.assertTrue(ws["G1"].font.bold)
        self.assertTrue(str(ws["G1"].font.color.rgb).endswith("C00000"))

    def test_create_summary_table_headless(self):
        parameters = [
            {"name": "Tasa Crecimiento", "value": 0.05, "format": "percent"},
            {"name": "Inflación", "value": 0.03, "format": "percent"}
        ]
        headers = ["Escenario", "Ingresos", "Costos", "Margen", "Variación"]
        rows_spec = [
            {"label": "Base", "values": [100000, 70000, "=C6-D6", 0.30], "formats": ["currency", "currency", "currency", "percent"], "is_total": False},
            {"label": "Optimista", "values": [120000, 80000, "=C7-D7", 0.33], "formats": ["currency", "currency", "currency", "percent"], "is_total": False},
            {"label": "Total / Consolidado", "values": ["=SUM(C6:C7)", "=SUM(D6:D7)", "=C8-D8", 0.31], "formats": ["currency", "currency", "currency", "percent"], "is_total": True}
        ]

        res = self.engine.create_summary_table(
            title="Resumen Comparativo de Escenarios",
            headers=headers,
            rows_spec=rows_spec,
            dest_sheet="Escenarios_Simulacion",
            start_cell="B2",
            parameters=parameters
        )
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["sheet"], "Escenarios_Simulacion")
        self.assertEqual(res["rows_count"], 3)
        self.assertIn("Tasa Crecimiento", res["parameters"])
        self.engine.save()

        ws = self.engine.backend._get_sheet("Escenarios_Simulacion")

        # Verificar cuadrícula activa
        self.assertTrue(ws.views.sheetView[0].showGridLines)

        # Verificar bloque de parámetros
        self.assertEqual(ws["B2"].value, "Supuestos y Parámetros")
        self.assertEqual(ws["B3"].value, "Tasa Crecimiento")
        self.assertEqual(ws["C3"].value, 0.05)
        self.assertEqual(ws["C3"].number_format, FinancialStyleGuide.PERCENT_FORMAT)
        self.assertTrue(str(ws["C3"].fill.start_color.rgb).endswith("FFF2CC"))

        # Verificar título de la tabla y encabezados
        self.assertEqual(ws["B6"].value, "Resumen Comparativo de Escenarios")
        self.assertEqual(ws["B7"].value, "Escenario")
        self.assertEqual(ws["C7"].value, "Ingresos")
        self.assertTrue(str(ws["B7"].fill.start_color.rgb).endswith("0E2E63"))
        self.assertTrue(str(ws["B7"].font.color.rgb).endswith("FFFFFF"))

        # Verificar filas de datos y total
        self.assertEqual(ws["B8"].value, "Base")
        self.assertEqual(ws["C8"].value, 100000)
        self.assertEqual(ws["C8"].number_format, FinancialStyleGuide.CURRENCY_FORMAT)

        self.assertEqual(ws["B10"].value, "Total / Consolidado")
        self.assertTrue(ws["B10"].font.bold)
        self.assertEqual(ws["B10"].border.bottom.style, "double")
        self.assertEqual(ws["B10"].border.top.style, "thin")

    def test_add_calculated_column_absolute_ref(self):
        data = [
            ["Ventas", "Costo"],
            [100, 80],
            [200, 150],
            [300, 220]
        ]
        self.engine.set_cell_range("A1", data)

        # Crear hoja auxiliar Params con celda C4
        ws_params = self.engine.backend._get_sheet("Params")
        ws_params["C4"] = 1.15

        # 1. Probar referencia absoluta con nombre de hoja externa: =A2*'Params'!$C$4
        res1 = self.engine.add_calculated_column("Ajustado", "=A2*'Params'!$C$4")
        self.assertEqual(res1["status"], "success")

        # 2. Probar referencia absoluta de fila en misma hoja: =$C$2*A2
        res2 = self.engine.add_calculated_column("Factorizado", "=$C$2*A2")
        self.assertEqual(res2["status"], "success")

        # 3. Probar plantilla con placeholder {r}
        res3 = self.engine.add_calculated_column("Placeholder", "=A{r}*'Params'!$C$4")
        self.assertEqual(res3["status"], "success")

        self.engine.save()

        ws = self.engine.backend._get_sheet(None)

        # Verificar que $C$4 se mantuvo estrictamente intacto en todas las filas
        self.assertEqual(ws["C2"].value, "=A2*'Params'!$C$4")
        self.assertEqual(ws["C3"].value, "=A3*'Params'!$C$4")
        self.assertEqual(ws["C4"].value, "=A4*'Params'!$C$4")

        # Verificar que $C$2 se mantuvo estrictamente intacto mientras A2 avanzó
        self.assertEqual(ws["D2"].value, "=$C$2*A2")
        self.assertEqual(ws["D3"].value, "=$C$2*A3")
        self.assertEqual(ws["D4"].value, "=$C$2*A4")

        # Verificar que {r} se sustituyó dinámicamente preservando $C$4
        self.assertEqual(ws["E2"].value, "=A2*'Params'!$C$4")
        self.assertEqual(ws["E3"].value, "=A3*'Params'!$C$4")
        self.assertEqual(ws["E4"].value, "=A4*'Params'!$C$4")

    def test_parse_excel_date_serial(self):
        # 1. Fechas formato ISO
        self.assertEqual(parse_excel_date_serial("2014-01-01"), 41640)
        self.assertEqual(parse_excel_date_serial("2014-01-01 3:00:00 AM"), 41640)
        self.assertEqual(parse_excel_date_serial("2014-01-01T15:30:00"), 41640)

        # 2. Fechas con slashes y marcas de hora
        self.assertEqual(parse_excel_date_serial("01/01/2014 3:00:00 AM"), 41640)
        self.assertEqual(parse_excel_date_serial("15/09/2013"), 41532)
        self.assertEqual(parse_excel_date_serial("09/15/2013"), 41532)

        # 3. Objetos nativos datetime y date
        self.assertEqual(parse_excel_date_serial(datetime.datetime(2014, 1, 1, 3, 0)), 41640)
        self.assertEqual(parse_excel_date_serial(datetime.date(2014, 1, 1)), 41640)

        # 4. Entero serial existente
        self.assertEqual(parse_excel_date_serial(41640), 41640)

        # 5. Valores inválidos o nulos
        self.assertIsNone(parse_excel_date_serial(None))
        self.assertIsNone(parse_excel_date_serial(""))
        self.assertIsNone(parse_excel_date_serial("invalido_no_fecha"))

    def test_remove_duplicates_headless(self):
        from openpyxl.worksheet.table import Table
        ws = self.engine.backend._get_sheet(None)
        data = [
            ["ID", "Segment", "Country", "Sales"],
            [1, "Government", "Canada", 100],
            [2, "Midmarket", "France", 200],
            [1, "Government", "Canada", 100],  # Duplicado exacto
            [3, "Enterprise", "Germany", 300],
            [2, "Midmarket", "France", 200],  # Duplicado exacto
        ]
        self.engine.set_cell_range("A1", data)

        # Crear tabla estructurada openpyxl
        tab = Table(displayName="TestTable", ref="A1:D6")
        ws.add_table(tab)

        # Ejecutar deduplicación on-demand
        res = self.engine.remove_duplicates(range_or_table="TestTable")
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["original_rows"], 5)
        self.assertEqual(res["duplicates_removed"], 2)
        self.assertEqual(res["remaining_rows"], 3)

        # Verificar que la tabla estructurada se redimensionó correctamente a 4 filas (1 encabezado + 3 datos)
        self.assertEqual(tab.ref, "A1:D4")
        self.assertEqual(ws.max_row, 4)

        # Verificar valores remanentes
        remaining_ids = [ws.cell(r, 1).value for r in range(2, 5)]
        self.assertEqual(remaining_ids, [1, 2, 3])

    def test_detect_and_flag_outliers_headless(self):
        data = [
            ["ID", "Profit"],
            [1, 100],
            [2, 105],
            [3, 95],
            [4, 102],
            [5, 98],
            [6, 101],
            [7, 99],
            [8, 103],
            [9, 99999]  # Outlier extremo
        ]
        self.engine.set_cell_range("A1", data)

        # Detectar outliers con método IQR
        res = self.engine.detect_and_flag_outliers(column="Profit", method="iqr", range_or_table="A1:B10", add_comment=True)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["outliers_count"], 1)
        self.assertIn("B10", res["flagged_cells"])

        ws = self.engine.backend._get_sheet(None)
        cell_outlier = ws["B10"]

        # Verificar que el valor NO fue eliminado (no destructivo)
        self.assertEqual(cell_outlier.value, 99999)

        # Verificar formato visual warning (#FCE4D6) y texto rojo oscuro
        self.assertIsNotNone(cell_outlier.fill)
        self.assertEqual(cell_outlier.fill.start_color.rgb, "00FCE4D6")
        self.assertTrue(cell_outlier.font.bold)
        self.assertEqual(cell_outlier.font.color.rgb, "00C00000")

        # Verificar comentario insertado
        self.assertIsNotNone(cell_outlier.comment)
        self.assertIn("Outlier detectado (IQR)", cell_outlier.comment.text)

        # Probar también con método zscore
        res_z = self.engine.detect_and_flag_outliers(column="Profit", method="zscore", range_or_table="A1:B10", add_comment=False)
        self.assertEqual(res_z["status"], "success")
        self.assertEqual(res_z["outliers_count"], 1)

    def test_clean_table_dataset_headless(self):
        data = [
            ["Segment", " Country ", "Units Sold", "Date", "Year"],
            ["  gov  ", "canada", "1500", "2014-01-01", "2014"],
            ["enterprise", "  usa", "2000.5", "01/06/2014 3:00:00 AM", "2014"]
        ]
        self.engine.set_cell_range("A1", data)

        rules = {
            "trim_headers": True,
            "standardize_text": {
                "Segment": {"trim": True, "casing": "title", "replacements": {"Gov": "Government"}},
                "Country": {"trim": True, "casing": "title", "replacements": {"Usa": "United States"}}
            },
            "cast_types": {
                "Units Sold": "float",
                "Year": "int"
            },
            "coerce_dates": ["Date"]
        }

        res = self.engine.clean_table_dataset(target_range_or_table="A1:E3", rules=rules)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["rows_processed"], 2)
        self.assertGreater(res["cells_modified"], 0)

        ws = self.engine.backend._get_sheet(None)

        # Verificar que el encabezado fue limpiado
        self.assertEqual(ws["B1"].value, "Country")

        # Verificar estandarización de texto
        self.assertEqual(ws["A2"].value, "Government")
        self.assertEqual(ws["B3"].value, "United States")

        # Verificar casteo de tipos numéricos
        self.assertIsInstance(ws["C2"].value, float)
        self.assertEqual(ws["C2"].value, 1500.0)
        self.assertIsInstance(ws["C3"].value, float)
        self.assertEqual(ws["C3"].value, 2000.5)

        self.assertIsInstance(ws["E2"].value, int)
        self.assertEqual(ws["E2"].value, 2014)

        # Verificar coerción de fechas a números seriales con formato uniforme
        self.assertEqual(ws["D2"].value, 41640)
        self.assertEqual(ws["D2"].number_format, "yyyy-mm-dd")
        self.assertEqual(ws["D3"].value, 41791)  # 2014-06-01
        self.assertEqual(ws["D3"].number_format, "yyyy-mm-dd")

    def test_curate_pipeline_headless(self):
        from openpyxl.worksheet.table import Table
        ws = self.engine.backend._get_sheet(None)
        data = [
            [" ID ", " Segment ", " Country ", "Sales", "Date", "Year"],
            [" 1 ", "  gov  ", "canada", "200", "2014-01-01", "2014"],
            [" 2 ", "midmarket", "france", "205.5", "2014-02-01", "2014"],
            [" 1 ", "  gov  ", "canada", "200", "2014-01-01", "2014"],  # Duplicado exacto
            [" 3 ", "enterprise", "  usa", "195", "2014-03-01", "2014"],
            [" 4 ", "government", "germany", "99999", "2014-04-01", "2014"],  # Outlier en Sales
            [" 5 ", "midmarket", "france", "202", "2014-05-01", "2014"],
            [" 6 ", "enterprise", "mexico", "198", "2014-06-01", "2014"],
            [" 7 ", "government", "canada", "201", "2014-07-01", "2014"],
        ]
        self.engine.set_cell_range("A1", data)


        tab = Table(displayName="CurateTable", ref="A1:F9")
        ws.add_table(tab)

        pipeline_spec = {
            "dedup": {"key_columns": ["ID", "Segment"]},
            "clean": {
                "trim_headers": True,
                "standardize_text": {
                    "Segment": {"trim": True, "casing": "title", "replacements": {"Gov": "Government"}},
                    "Country": {"trim": True, "casing": "title", "replacements": {"Usa": "United States"}}
                },
                "cast_types": {
                    "ID": "int",
                    "Sales": "float",
                    "Year": "int"
                },
                "coerce_dates": ["Date"]
            },
            "outliers": [
                {"column": "Sales", "method": "iqr", "add_comment": True}
            ]
        }

        res = self.engine.curate_pipeline(target_range_or_table="CurateTable", pipeline_spec=pipeline_spec)
        self.assertEqual(res["status"], "success")
        self.assertEqual(res["original_rows"], 8)
        self.assertEqual(res["duplicates_removed"], 1)
        self.assertEqual(res["remaining_rows"], 7)
        self.assertGreater(res["cells_cleaned"], 0)
        self.assertEqual(res["outliers_count"], 1)
        self.assertEqual(len(res["outliers_flagged"]), 1)
        self.assertIn("D5", res["outliers_flagged"])

        # Verificar redimensionamiento de tabla estructurada: 1 encabezado + 7 datos = fila 8
        self.assertEqual(tab.ref, "A1:F8")
        self.assertEqual(ws.max_row, 8)

        # Verificar limpieza de encabezado
        self.assertEqual(ws["A1"].value, "ID")
        self.assertEqual(ws["B1"].value, "Segment")

        # Verificar que el duplicado fue eliminado (remanentes IDs: 1, 2, 3, 4, 5, 6, 7)
        ids = [ws.cell(r, 1).value for r in range(2, 9)]
        self.assertEqual(ids, [1, 2, 3, 4, 5, 6, 7])

        # Verificar texto estandarizado
        self.assertEqual(ws["B2"].value, "Government")
        self.assertEqual(ws["C4"].value, "United States")

        # Verificar tipos casteados
        self.assertIsInstance(ws["D2"].value, float)
        self.assertEqual(ws["D2"].value, 200.0)
        self.assertIsInstance(ws["F2"].value, int)

        self.assertEqual(ws["F2"].value, 2014)

        # Verificar fecha serial
        self.assertEqual(ws["E2"].value, 41640)
        self.assertEqual(ws["E2"].number_format, "yyyy-mm-dd")

        # Verificar formato visual de outlier (no destructivo)
        cell_outlier = ws["D5"]
        self.assertEqual(cell_outlier.value, 99999.0)
        self.assertIsNotNone(cell_outlier.fill)
        self.assertEqual(cell_outlier.fill.start_color.rgb, "00FCE4D6")
        self.assertTrue(cell_outlier.font.bold)
        self.assertEqual(cell_outlier.font.color.rgb, "00C00000")
        self.assertIsNotNone(cell_outlier.comment)
        self.assertIn("Outlier detectado", cell_outlier.comment.text)


if __name__ == "__main__":
    unittest.main()

