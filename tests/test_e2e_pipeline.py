# Creado por Francisco Grandón Vergara
"""
Suite de pruebas End-to-End (E2E) para flujos completos combinados y sucesivos
sobre tablas de datos en AntigravityExcelEngine (modo headless / file).

Verifica:
1. test_e2e_full_lifecycle_data_table: Ciclo de vida completo de 9 pasos sobre dataset operativo 'Ventas_Agua'.
2. test_e2e_chained_formula_injection_and_resize: 3 fórmulas encadenadas e interdependientes con expansión de tabla.
3. test_e2e_curate_aggregate_and_pivot_flow: Flujo curación -> agregación -> tabla dinámica cruzada y gráfico.
4. test_e2e_declarative_scenario_and_styling: Modelo de supuestos (set_cells), fórmulas dinámicas, estilos y exportación CSV.
"""

import os
import sys
import uuid
import unittest
import openpyxl
from openpyxl.worksheet.table import Table, TableColumn

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from antigravity_excel_engine.engine import AntigravityExcelEngine
from antigravity_excel_engine.formatting import FinancialStyleGuide


class TestE2EPipeline(unittest.TestCase):
    def setUp(self):
        self.test_file = os.path.abspath(f"test_e2e_{uuid.uuid4().hex[:8]}.xlsx")
        if os.path.exists(self.test_file):
            try:
                os.remove(self.test_file)
            except Exception:
                pass
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

    def test_e2e_full_lifecycle_data_table(self):
        """
        Prueba E2E #1: Ciclo de vida completo de 9 pasos sobre dataset operativo de facturación de agua.
        Paso 1: Auditoría inicial de calidad (detecta problemas de espacios en blanco y números como texto).
        Paso 2: Curación en pipeline único (dedup + clean + outliers IQR).
        Paso 3: Re-auditoría de calidad verificando 0 incidencias.
        Paso 4: Inyección de columnas calculadas encadenadas (IVA, Total_Con_IVA, Consumo_Diario).
        Paso 5: Auditoría de errores de fórmulas (0 errores #VALUE!, #REF!, etc.).
        Paso 6: Agregación analítica por grupo sobre columna calculada.
        Paso 7: Tabla dinámica y gráfico analítico en hoja independiente.
        Paso 8: Estilización ejecutiva con paleta corporativa y presets de formato.
        Paso 9: Operaciones declarativas de celdas, copia de rango y exportación CSV.
        """
        ws = self.engine.backend._get_sheet(None)
        ws.title = "Operaciones"

        # Dataset operativo con impurezas:
        # - Espacios en blanco en encabezados (" Localidad ", " Consumo_m3 ")
        # - Duplicado exacto en fila 4 y fila 11 (ID 103)
        # - Valores numéricos como texto ("12500", "2024")
        # - Outlier evidente en Consumo_m3 (fila 9, ID 108 con 9999 m3)
        raw_headers = [
            " ID_Cliente ", " Localidad ", "Tipo_Tarifa", " Consumo_m3 ",
            "Tarifa_Base", "Fecha_Lectura", " Periodo_Anio ", "Estado_Pago"
        ]
        raw_data = [
            raw_headers,
            ["101", "  concepcion ", "residencial", "25.5", "12500", "2024-01-10", "2024", "PAGADO"],
            ["102", "TALCAHUANO", "comercial", "80.0", "45000", "2024-01-11", "2024", "PAGADO"],
            ["103", "  concepcion ", "industrial", "310.2", "180000", "2024-01-12", "2024", "PENDIENTE"],
            ["104", "CHILLAN", "residencial", "18.0", "9500", "2024-01-13", "2024", "PAGADO"],
            ["105", "concepcion", "comercial", "95.5", "52000", "2024-01-14", "2024", "PAGADO"],
            ["106", "TALCAHUANO", "residencial", "22.0", "11000", "2024-01-15", "2024", "PENDIENTE"],
            ["107", "CHILLAN", "industrial", "250.0", "150000", "2024-01-16", "2024", "PAGADO"],
            ["108", "concepcion", "industrial", "9999.0", "500000", "2024-01-17", "2024", "EN_REVISION"], # Outlier
            ["109", "talcahuano", "residencial", "20.0", "10500", "2024-01-18", "2024", "PAGADO"],
            ["103", "  concepcion ", "industrial", "310.2", "180000", "2024-01-12", "2024", "PENDIENTE"], # Duplicado exacto
        ]

        self.engine.set_cell_range("A1", raw_data)
        tab = Table(displayName="Ventas_Agua", ref="A1:H11")
        tab.tableColumns = [TableColumn(id=i, name=c.strip()) for i, c in enumerate(raw_headers, 1)]
        ws.add_table(tab)
        self.engine.save()

        # --- Paso 1: Auditoría inicial de calidad de datos ---
        audit_pre = self.engine.audit_data_quality(sheet="Operaciones")
        self.assertGreater(len(audit_pre["header_whitespace_issues"]), 0)
        self.assertGreater(len(audit_pre["numbers_stored_as_text"]), 0)

        # --- Paso 2: Curación compuesta con curate_pipeline ---
        pipeline_spec = {
            "dedup": {
                "key_columns": ["ID_Cliente", "Localidad", "Fecha_Lectura"]
            },
            "clean": {
                "trim_headers": True,
                "standardize_text": {
                    "Localidad": {"trim": True, "casing": "upper"},
                    "Tipo_Tarifa": {"trim": True, "casing": "title"}
                },
                "cast_types": {
                    "ID_Cliente": "int",
                    "Consumo_m3": "float",
                    "Tarifa_Base": "float",
                    "Periodo_Anio": "int"
                },
                "coerce_dates": ["Fecha_Lectura"]
            },
            "outliers": [
                {"column": "Consumo_m3", "method": "iqr", "add_comment": True}
            ]
        }

        res_curate = self.engine.curate_pipeline(target_range_or_table="Ventas_Agua", pipeline_spec=pipeline_spec, sheet="Operaciones")
        self.assertEqual(res_curate["status"], "success")
        self.assertEqual(res_curate["original_rows"], 10)
        self.assertEqual(res_curate["duplicates_removed"], 1)
        self.assertEqual(res_curate["remaining_rows"], 9)
        self.assertEqual(res_curate["outliers_count"], 1)
        # La tabla debe haberse redimensionado a 10 filas totales (1 cabecera + 9 datos)
        self.assertEqual(tab.ref, "A1:H10")

        # --- Paso 3: Re-auditoría post-curación ---
        audit_post = self.engine.audit_data_quality(sheet="Operaciones")
        self.assertEqual(len(audit_post["header_whitespace_issues"]), 0)
        # Los encabezados ya están trimmeados y limpios
        self.assertEqual(ws["A1"].value, "ID_Cliente")
        self.assertEqual(ws["B1"].value, "Localidad")
        self.assertEqual(ws["D1"].value, "Consumo_m3")

        # --- Paso 4: Inyección de columnas calculadas encadenadas ---
        # Columna E es Tarifa_Base (col 5)
        # Columna I (col 9): IVA = E{r} * 0.19
        res_iva = self.engine.add_calculated_column(
            header="IVA",
            formula_template="=E{r}*0.19",
            number_format=FinancialStyleGuide.CURRENCY_CLP,
            sheet="Operaciones"
        )
        self.assertEqual(res_iva["status"], "success")
        self.assertEqual(res_iva["column_index"], 9)

        # Columna J (col 10): Total_Con_IVA = E{r} + I{r}
        res_total = self.engine.add_calculated_column(
            header="Total_Con_IVA",
            formula_template="=E{r}+I{r}",
            number_format=FinancialStyleGuide.CURRENCY_CLP,
            sheet="Operaciones"
        )
        self.assertEqual(res_total["status"], "success")
        self.assertEqual(res_total["column_index"], 10)

        # Columna K (col 11): Consumo_Diario = D{r} / 30
        res_diario = self.engine.add_calculated_column(
            header="Consumo_Diario",
            formula_template="=D{r}/30",
            number_format="0.00",
            sheet="Operaciones"
        )
        self.assertEqual(res_diario["status"], "success")
        self.assertEqual(res_diario["column_index"], 11)

        # Verificar expansión de la tabla a A1:K10
        self.assertEqual(tab.ref, "A1:K10")
        self.assertEqual(len(tab.tableColumns), 11)

        # --- Paso 5: Auditoría de errores de fórmulas ---
        formula_audit = self.engine.check_formula_errors(sheet="Operaciones")
        self.assertIsInstance(formula_audit, list)
        self.assertEqual(len(formula_audit), 0)

        # --- Paso 6: Agregación analítica sobre columna calculada ---
        agg_res = self.engine.aggregate(
            group_by_col="Localidad",
            metric_col="Total_Con_IVA",
            agg_func="sum",
            sheet="Operaciones"
        )
        self.assertGreater(agg_res["groups_count"], 0)
        self.assertIn("ranking", agg_res)
        top_group = agg_res["ranking"][0]["group"]
        self.assertIn(top_group, ["CONCEPCION", "CHILLAN", "TALCAHUANO"])

        # --- Paso 7: Tabla dinámica y Gráfico analítico ---
        res_pivot = self.engine.create_pivot_table(
            source="A1:K10",
            rows=["Localidad"],
            cols=["Tipo_Tarifa"],
            value_field="Total_Con_IVA",
            value_func="sum",
            value_format="$#,##0",
            dest_sheet="Resumen_Tarifas",
            dest_cell="A3",
            source_sheet="Operaciones"
        )
        self.assertEqual(res_pivot["status"], "success")
        self.assertEqual(res_pivot["dest_sheet"], "Resumen_Tarifas")
        self.assertGreater(res_pivot["grand_total"], 0)

        # Crear gráfico en la misma hoja resumen
        res_chart = self.engine.create_chart(
            source="A3:D7",
            chart_type="column_clustered",
            title="Consolidado de Facturación por Localidad",
            placement="auto",
            dest_sheet="Resumen_Tarifas"
        )
        self.assertEqual(res_chart["status"], "success")
        self.assertEqual(res_chart["sheet"], "Resumen_Tarifas")

        # --- Paso 8: Aplicación de estilos corporativos ---
        res_style = self.engine.apply_style_preset("A3", "header", sheet="Resumen_Tarifas")
        self.assertEqual(res_style["status"], "success")
        ws_pivot = self.engine.backend._get_sheet("Resumen_Tarifas")
        self.assertEqual(ws_pivot["A3"].fill.start_color.rgb, "000E2E63")
        self.assertEqual(ws_pivot["A3"].font.color.rgb, "00FFFFFF")

        # --- Paso 9: Operaciones declarativas, copia y exportación CSV ---
        # Escribir fila de totales y notas de auditoría con set_cells declarativo (dict)
        dec_cells = {
            "A12": {"value": "Total General Auditado:", "style": "header"},
            "B12": {"value": res_pivot["grand_total"], "number_format": FinancialStyleGuide.CURRENCY_CLP}
        }
        res_set = self.engine.set_cells(dec_cells, sheet="Resumen_Tarifas")
        self.assertEqual(res_set["status"], "success")
        self.assertEqual(res_set["cells_updated"], 2)

        # Copiar rango de resumen a bloque histórico en la misma hoja
        res_copy = self.engine.copy_paste_range(
            source_range="A12:B12",
            dest_cell="D12",
            source_sheet="Resumen_Tarifas",
            dest_sheet="Resumen_Tarifas"
        )
        self.assertEqual(res_copy["status"], "success")

        # Exportar rango como CSV para interoperabilidad
        csv_export = self.engine.get_range_as_csv("A12:B12", sheet="Resumen_Tarifas", delimiter=";")
        self.assertIn("Total General Auditado:", csv_export)

        # Guardar y validar persistencia en disco
        saved_file = self.engine.save()
        self.assertTrue(os.path.exists(saved_file))

        # Reabrir con openpyxl nativo para validación cruzada
        wb_check = openpyxl.load_workbook(saved_file, data_only=False)
        self.assertIn("Operaciones", wb_check.sheetnames)
        self.assertIn("Resumen_Tarifas", wb_check.sheetnames)
        ws_op_check = wb_check["Operaciones"]
        self.assertEqual(ws_op_check["I2"].value, "=E2*0.19")
        self.assertEqual(ws_op_check["J2"].value, "=E2+I2")
        self.assertEqual(ws_op_check["K2"].value, "=D2/30")
        wb_check.close()

    def test_e2e_chained_formula_injection_and_resize(self):
        """
        Prueba E2E #2: 3 fórmulas encadenadas e interdependientes que expanden la estructura
        y dimensiones de una tabla de datos de 3 a 6 columnas, verificando referencias y consistencia.
        """
        ws = self.engine.backend._get_sheet(None)
        ws.title = "Costos"

        headers = ["Producto", "Unidades", "Precio_Unitario", "Costo_Unitario"]
        data = [
            headers,
            ["Medidor DN15", 150, 12000, 7500],
            ["Medidor DN20", 80, 18500, 11000],
            ["Válvula Esférica", 300, 4500, 2200],
            ["Abrazadera Reparación", 120, 8900, 4100],
            ["Tubo PEAD 32mm", 50, 25000, 14500]
        ]
        self.engine.set_cell_range("A1", data)

        tab = Table(displayName="TblCostos", ref="A1:D6")
        tab.tableColumns = [TableColumn(id=i, name=c) for i, c in enumerate(headers, 1)]
        ws.add_table(tab)
        self.engine.save()

        # Fórmula 1: Ingreso_Bruto = B{r} * C{r} (Columna E, index 5)
        res_f1 = self.engine.add_calculated_column(
            header="Ingreso_Bruto",
            formula_template="=B{r}*C{r}",
            number_format=FinancialStyleGuide.CURRENCY_CLP
        )
        self.assertEqual(res_f1["status"], "success")
        self.assertEqual(res_f1["column_index"], 5)
        self.assertEqual(tab.ref, "A1:E6")

        # Fórmula 2: Costo_Total = B{r} * D{r} (Columna F, index 6)
        res_f2 = self.engine.add_calculated_column(
            header="Costo_Total",
            formula_template="=B{r}*D{r}",
            number_format=FinancialStyleGuide.CURRENCY_CLP
        )
        self.assertEqual(res_f2["status"], "success")
        self.assertEqual(res_f2["column_index"], 6)
        self.assertEqual(tab.ref, "A1:F6")

        # Fórmula 3: Margen_Neto = E{r} - F{r} (Columna G, index 7, depende de F1 y F2)
        res_f3 = self.engine.add_calculated_column(
            header="Margen_Neto",
            formula_template="=E{r}-F{r}",
            number_format=FinancialStyleGuide.CURRENCY_CLP
        )
        self.assertEqual(res_f3["status"], "success")
        self.assertEqual(res_f3["column_index"], 7)
        self.assertEqual(tab.ref, "A1:G6")

        # Verificar lista de columnas registradas en openpyxl Table
        self.assertEqual(len(tab.tableColumns), 7)
        col_names = [c.name for c in tab.tableColumns]
        self.assertEqual(col_names, ["Producto", "Unidades", "Precio_Unitario", "Costo_Unitario", "Ingreso_Bruto", "Costo_Total", "Margen_Neto"])

        # Verificar fórmulas en filas
        ranges = self.engine.get_cell_ranges(["E2:G6"], include_formulas=True)
        formulas_matrix = ranges["E2:G6"]["formulas"]
        for row_idx, r in enumerate(range(2, 7)):
            self.assertEqual(formulas_matrix[row_idx][0], f"=B{r}*C{r}")
            self.assertEqual(formulas_matrix[row_idx][1], f"=B{r}*D{r}")
            self.assertEqual(formulas_matrix[row_idx][2], f"=E{r}-F{r}")

        # Auditoría de errores en fórmulas encadenadas
        err_check = self.engine.check_formula_errors()
        self.assertIsInstance(err_check, list)
        self.assertEqual(len(err_check), 0)

        # Evaluación aritmética en backend headless vía aggregate
        agg_margen = self.engine.aggregate(group_by_col="Producto", metric_col="Margen_Neto", agg_func="sum")
        self.assertEqual(agg_margen["groups_count"], 5)
        # Medidor DN15: Ingreso=150*12000=1.800.000, Costo=150*7500=1.125.000 -> Margen=675.000
        dn15_group = next(g for g in agg_margen["ranking"] if g["group"] == "Medidor DN15")
        self.assertEqual(dn15_group["aggregated_value"], 675000.0)

    def test_e2e_curate_aggregate_and_pivot_flow(self):
        """
        Prueba E2E #3: Curación de dataset -> Agregación resumen -> Tabla dinámica
        con verificación matemática (Grand Total == suma de valores agrupados) -> Gráfico.
        """
        ws = self.engine.backend._get_sheet(None)
        ws.title = "Produccion"

        headers = ["Planta", "Zona", "Volumen_Tratado_m3", "Costo_Quimicos"]
        raw_data = [
            headers,
            ["Planta La Mochita", "Sur", "120000", "4500000"],
            ["Planta Curicó", "Norte", "85000", "2800000"],
            ["Planta La Mochita", "Sur", "120000", "4500000"], # Duplicado exacto
            ["Planta Chiguayante", "Sur", "65000", "2100000"],
            ["Planta Talca", "Norte", "92000", "3100000"],
            ["Planta Chillán", "Centro", "78000", "2600000"],
            ["Planta Los Ángeles", "Sur", "88000", "2950000"]
        ]
        self.engine.set_cell_range("A1", raw_data)

        # 1. Curación
        pipeline_spec = {
            "dedup": {"key_columns": ["Planta", "Zona", "Volumen_Tratado_m3"]},
            "clean": {
                "trim_headers": True,
                "cast_types": {
                    "Volumen_Tratado_m3": "float",
                    "Costo_Quimicos": "float"
                }
            }
        }
        res_curate = self.engine.curate_pipeline(target_range_or_table="A1:D8", pipeline_spec=pipeline_spec)
        self.assertEqual(res_curate["duplicates_removed"], 1)
        self.assertEqual(res_curate["remaining_rows"], 6)

        # 2. Agregación por Zona
        res_agg = self.engine.aggregate(group_by_col="Zona", metric_col="Costo_Quimicos", agg_func="sum")
        self.assertEqual(res_agg["groups_count"], 3)
        expected_total = sum(g["aggregated_value"] for g in res_agg["ranking"])
        # Valores: Sur (4500000 + 2100000 + 2950000 = 9550000), Norte (2800000 + 3100000 = 5900000), Centro (2600000)
        # Total esperado: 18050000.0
        self.assertEqual(expected_total, 18050000.0)

        # 3. Tabla dinámica cruzada en hoja Pivot_Analitica
        res_pivot = self.engine.create_pivot_table(
            source="A1:D7",
            rows=["Planta"],
            cols=["Zona"],
            value_field="Costo_Quimicos",
            value_func="sum",
            value_format=FinancialStyleGuide.CURRENCY_CLP,
            dest_sheet="Pivot_Analitica",
            dest_cell="B2"
        )
        self.assertEqual(res_pivot["status"], "success")
        self.assertEqual(res_pivot["grand_total"], expected_total)

        # 4. Gráfico de barras en Pivot_Analitica
        res_chart = self.engine.create_chart(
            source="B2:F9",
            chart_type="column_clustered",
            title="Costos Químicos por Planta y Zona",
            placement="bottom",
            sheet="Pivot_Analitica"
        )
        self.assertEqual(res_chart["status"], "success")
        self.assertEqual(res_chart["placement"], "bottom")

    def test_e2e_declarative_scenario_and_styling(self):
        """
        Prueba E2E #4: Modelo financiero declarativo con supuestos (set_cells),
        fórmulas de proyección ($B$2), presets de estilo (input_cell, total_row, header)
        y exportación limpia en CSV.
        """
        ws = self.engine.backend._get_sheet(None)
        ws.title = "Modelo_Tarifario"

        # 1. Definición declarativa de supuestos / inputs (formato dict)
        assumption_cells = {
            "A1": {"value": "MODELO DE SIMULACIÓN TARIFARIA"},
            "A3": {"value": "Factor Inflación (IPC):"},
            "B3": {"value": 0.045},
            "A4": {"value": "Tasa Crecimiento Clientes:"},
            "B4": {"value": 0.020},
            "A6": {"value": "Proyección Quinquenal"}
        }
        self.engine.set_cells(assumption_cells)
        self.engine.apply_style_preset("A1", "header")
        self.engine.apply_style_preset("B3", "input_cell")
        self.engine.apply_style_preset("B4", "input_cell")
        self.engine.apply_style_preset("A6", "header")

        # 2. Tabla de proyección
        table_headers = [["Año", "Clientes_Base", "Ingreso_Proyectado"]]
        table_data = [
            [2025, 500000, "=B8*(1+$B$4)*15000*(1+$B$3)"],
            [2026, "=B8*(1+$B$4)", "=B9*(1+$B$4)*15000*(1+$B$3)"],
            [2027, "=B9*(1+$B$4)", "=B10*(1+$B$4)*15000*(1+$B$3)"],
        ]
        self.engine.set_cell_range("A7", table_headers)
        self.engine.set_cell_range("A8", table_data)

        # Fila de totales declarativa (formato dict)
        summary_cells = {
            "A11": {"value": "Total Proyección:"},
            "C11": {"value": "=SUM(C8:C10)"}
        }
        self.engine.set_cells(summary_cells)
        self.engine.apply_style_preset("A11", "total_row")
        self.engine.apply_style_preset("C11", "total_row")
        self.engine.apply_style_preset("C11", "currency_clp")

        # 3. Verificación de estilos aplicados
        self.assertTrue(str(ws["B3"].fill.start_color.rgb).endswith("FFF2CC"))
        self.assertTrue(str(ws["B3"].font.color.rgb).endswith("1F4E79"))

        self.assertTrue(ws["A11"].font.bold)
        self.assertEqual(ws["A11"].border.top.style, "thin")
        self.assertEqual(ws["A11"].border.bottom.style, "double")

        # 4. Comprobación de errores de fórmula
        err_res = self.engine.check_formula_errors(range_str="A7:C11")
        self.assertIsInstance(err_res, list)
        self.assertEqual(len(err_res), 0)

        # 5. Exportación CSV y validación de contenido
        csv_content = self.engine.get_range_as_csv("A7:C11", delimiter=",")
        self.assertIn("Total Proyección:", csv_content)
        self.assertIn("Año,Clientes_Base,Ingreso_Proyectado", csv_content)

        # Guardar archivo final
        saved_path = self.engine.save()
        self.assertTrue(os.path.exists(saved_path))


if __name__ == "__main__":
    unittest.main()
