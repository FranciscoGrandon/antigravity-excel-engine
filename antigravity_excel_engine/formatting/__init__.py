# Creado por Francisco Grandón Vergara
"""
Subpaquete de formateo, estilos visuales y directrices contables de Excel.
"""

from .presets import FinancialStyleGuide
from .styles import (
    apply_borders_openpyxl,
    apply_number_format,
    apply_style_preset,
    apply_style_preset_com,
    apply_style_preset_openpyxl,
    hex_to_argb_hex,
    hex_to_rgb_int,
)

__all__ = [
    "FinancialStyleGuide",
    "hex_to_rgb_int",
    "hex_to_argb_hex",
    "apply_number_format",
    "apply_borders_openpyxl",
    "apply_style_preset_openpyxl",
    "apply_style_preset_com",
    "apply_style_preset",
]
