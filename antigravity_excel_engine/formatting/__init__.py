# Creado por Francisco Grandón Vergara
"""
Subpaquete de formateo, estilos visuales y directrices contables de Excel.
"""

from .presets import FinancialStyleGuide
from .styles import (
    hex_to_rgb_int,
    hex_to_argb_hex,
    apply_number_format,
    apply_borders_openpyxl,
    apply_style_preset_openpyxl,
    apply_style_preset_com,
    apply_style_preset,
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
