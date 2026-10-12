# Creado por Francisco Grandón Vergara
"""
Directrices, constantes y paletas visuales para modelado financiero y presentación en Excel.
"""

from typing import Any, Optional, Dict


class FinancialStyleGuide:
    """
    Constantes y directrices de modelado financiero estándar.
    Proporciona formatos numéricos y paletas visuales corporativas para
    modelado financiero profesional en Excel.
    """
    # Formatos de moneda y numéricos
    CURRENCY_FORMAT = "$ #,##0.00;($ #,##0.00);\"-\""
    CURRENCY_FORMAT_LOCAL = "$ #.##0,00;($ #.##0,00);\"-\""
    CURRENCY_CLP = "$ #,##0;($ #,##0);\"-\""
    CURRENCY_CLP_LOCAL = "$ #.##0;($ #.##0);\"-\""
    CURRENCY_USD = "$ #,##0.00;($ #,##0.00);\"-\""
    PERCENT_FORMAT = "0.0%"
    PERCENT_FORMAT_LOCAL = "0,0%"
    PERCENT = "0.0%"
    DATE_FORMAT = "YYYY-MM-DD"
    DATE_FORMAT_LOCAL = "DD-MM-YYYY"
    DATE = "YYYY-MM-DD"

    # Colores base corporativos (Hexadecimal)
    COLOR_PRIMARY = "#0E2E63"          # Azul marino corporativo Essbio
    COLOR_SECONDARY = "#1F4E79"        # Azul oscuro institucional
    COLOR_INPUT = "#1F4E79"            # Azul corporativo para celdas de entrada
    COLOR_FORMULA = "#000000"          # Negro para fórmulas
    COLOR_ASSUMPTION_FILL = "#FFF2CC"  # Amarillo marfil para supuestos / escenarios
    COLOR_HEADER_FILL = "#0E2E63"      # Fondo de cabeceras de tablas
    COLOR_HEADER_TEXT = "#FFFFFF"      # Texto de cabeceras
    COLOR_ZEBRA_EVEN = "#F2F4F8"       # Alternancia de filas (par)
    COLOR_ZEBRA_ODD = "#FFFFFF"        # Alternancia de filas (impar)
    COLOR_KPI_CARD_FILL = "#F8F9FA"    # Fondo para tarjetas de resumen / KPIs
    COLOR_KPI_CARD_BORDER = "#D9D9D9"  # Borde suave para tarjetas
    COLOR_DELTA_POS = "#008000"        # Verde para variaciones positivas
    COLOR_DELTA_NEG = "#C00000"        # Rojo oscuro para variaciones negativas
    COLOR_WARNING_FILL = "#FCE4D6"     # Salmón suave para alertas / outliers
    COLOR_BORDER_DEFAULT = "#D9D9D9"   # Borde gris claro para celdas estándar
    COLOR_BORDER_STRONG = "#1F4E79"    # Borde azul oscuro para entradas destacadas

    # Presets canónicos estructurados
    PRESETS: Dict[str, Dict[str, Any]] = {
        "header": {
            "fill": "#0E2E63",
            "font_color": "#FFFFFF",
            "bold": True,
            "alignment_vertical": "center",
            "alignment_horizontal": None,
        },
        "input_cell": {
            "fill": "#FFF2CC",
            "font_color": "#1F4E79",
            "bold": False,
            "border_color": "#1F4E79",
            "border_style": "medium",
            "alignment_horizontal": "center",
            "alignment_vertical": "center",
        },
        "total_row": {
            "bold": True,
            "border_top_style": "thin",
            "border_top_color": "#000000",
            "border_bottom_style": "double",
            "border_bottom_color": "#000000",
        },
        "kpi_card": {
            "fill": "#F8F9FA",
            "font_color": "#0E2E63",
            "bold": True,
            "border_color": "#D9D9D9",
            "border_style": "thin",
            "alignment_horizontal": "center",
            "alignment_vertical": "center",
        },
        "zebra_even": {
            "fill": "#F2F4F8",
        },
        "zebra_odd": {
            "fill": "#FFFFFF",
        },
        "delta_positive": {
            "font_color": "#008000",
            "bold": True,
        },
        "delta_negative": {
            "font_color": "#C00000",
            "bold": True,
        },
        "currency": {
            "number_format": CURRENCY_FORMAT,
        },
        "currency_clp": {
            "number_format": CURRENCY_CLP,
        },
        "percent": {
            "number_format": PERCENT_FORMAT,
        },
        "date": {
            "number_format": DATE,
        },
    }

    @staticmethod
    def apply_number_format_resilient(excel_cell_or_range, fmt_type_or_mask: str):
        """Aplica formato numérico probando NumberFormatLocal y NumberFormat para máxima compatibilidad regional."""
        if not fmt_type_or_mask:
            return
        t = fmt_type_or_mask.lower()
        if t in ("currency", "currency_usd"):
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.CURRENCY_FORMAT_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.CURRENCY_FORMAT
        elif t == "currency_clp":
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.CURRENCY_CLP_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.CURRENCY_CLP
        elif t in ("percent", "percentage"):
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.PERCENT_FORMAT_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.PERCENT_FORMAT
        elif t == "date":
            try:
                excel_cell_or_range.NumberFormatLocal = FinancialStyleGuide.DATE_FORMAT_LOCAL
            except Exception:
                excel_cell_or_range.NumberFormat = FinancialStyleGuide.DATE_FORMAT
        else:
            try:
                excel_cell_or_range.NumberFormatLocal = fmt_type_or_mask
            except Exception:
                excel_cell_or_range.NumberFormat = fmt_type_or_mask
