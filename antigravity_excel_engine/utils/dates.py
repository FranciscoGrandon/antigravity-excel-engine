# Creado por Francisco Grandón Vergara
"""
Utilidades de conversión y serialización de fechas en formato nativo de Excel.
Base estándar: 1899-12-30 (compatibilidad total con el bug bisiesto de Excel 1900).
"""

import datetime
import re
from typing import Any, Optional, Union

EXCEL_BASE_DATE = datetime.date(1899, 12, 30)
EXCEL_BASE_DATETIME = datetime.datetime(1899, 12, 30)


def parse_excel_date_serial(value: Any) -> Optional[int]:
    """
    Convierte un valor de fecha a su número serial exacto de Excel (días transcurridos desde 1899-12-30).
    Soporta:
      - Cadenas con formato ISO 'YYYY-MM-DD' o 'YYYY/MM/DD'.
      - Formatos 'DD/MM/YYYY' o 'MM/DD/YYYY' (con o sin marcas de hora como '3:00:00 AM' o '15:30:00').
      - Formatos ISO con hora ('YYYY-MM-DDTHH:MM:SS').
      - Objetos nativos datetime.datetime y datetime.date.
      - Enteros o flotantes que ya correspondan a seriales válidos de Excel.
    Retorna None si no es convertible o es nulo.
    """
    if value is None:
        return None

    if isinstance(value, (datetime.datetime, datetime.date)):
        d = value.date() if isinstance(value, datetime.datetime) else value
        return (d - EXCEL_BASE_DATE).days

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 2958465:
            return int(value)
        return None

    val_str = str(value).strip()
    if not val_str:
        return None

    # 1. ISO YYYY-MM-DD o YYYY/MM/DD con o sin hora
    m_iso = re.match(r"^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:[ T].*)?$", val_str)
    if m_iso:
        y, m, d = int(m_iso.group(1)), int(m_iso.group(2)), int(m_iso.group(3))
        try:
            return (datetime.date(y, m, d) - EXCEL_BASE_DATE).days
        except Exception:
            pass

    # 2. DD/MM/YYYY o MM/DD/YYYY con 4 dígitos en el año con o sin hora
    m_slash = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{4})(?:[ T].*)?$", val_str)
    if m_slash:
        n1, n2, y = int(m_slash.group(1)), int(m_slash.group(2)), int(m_slash.group(3))
        if n1 > 12 and n2 <= 12:
            d, m = n1, n2
        elif n2 > 12 and n1 <= 12:
            m, d = n1, n2
        else:
            d, m = n1, n2
        try:
            return (datetime.date(y, m, d) - EXCEL_BASE_DATE).days
        except Exception:
            pass

    # 3. Formato con 2 dígitos en el año (DD/MM/YY o MM/DD/YY)
    m_slash2 = re.match(r"^(\d{1,2})[-/](\d{1,2})[-/](\d{2})(?:[ T].*)?$", val_str)
    if m_slash2:
        n1, n2, y_short = int(m_slash2.group(1)), int(m_slash2.group(2)), int(m_slash2.group(3))
        y = y_short + 2000 if y_short < 50 else y_short + 1900
        if n1 > 12 and n2 <= 12:
            d, m = n1, n2
        elif n2 > 12 and n1 <= 12:
            m, d = n1, n2
        else:
            d, m = n1, n2
        try:
            return (datetime.date(y, m, d) - EXCEL_BASE_DATE).days
        except Exception:
            pass

    # 4. Fallback exhaustivo con strptime
    formats = [
        "%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %I:%M:%S %p", "%d/%m/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M:%S %p",
        "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f"
    ]
    for fmt in formats:
        try:
            dt = datetime.datetime.strptime(val_str, fmt)
            return (dt.date() - EXCEL_BASE_DATE).days
        except Exception:
            continue

    return None


def coerce_date_to_excel_serial(value: Any, default: Optional[int] = None) -> Optional[int]:
    """
    Coacciona un valor a serial de fecha de Excel. Si no es convertible,
    retorna el valor por defecto provisto (por omisión None).
    """
    res = parse_excel_date_serial(value)
    return res if res is not None else default


def excel_serial_to_datetime(serial: Union[int, float]) -> Optional[datetime.datetime]:
    """
    Convierte un número serial de Excel (incluyendo fracción de día como hora) a un datetime de Python.
    Retorna None si el serial no es válido o está fuera de rango.
    """
    if serial is None or not isinstance(serial, (int, float)) or isinstance(serial, bool):
        return None
    if not (1 <= serial <= 2958465):
        return None
    try:
        days = int(serial)
        fraction = serial - days
        dt = EXCEL_BASE_DATETIME + datetime.timedelta(days=days, seconds=round(fraction * 86400))
        return dt
    except Exception:
        return None


def excel_serial_to_date(serial: Union[int, float]) -> Optional[datetime.date]:
    """Convierte un número serial de Excel a un datetime.date de Python."""
    dt = excel_serial_to_datetime(serial)
    return dt.date() if dt else None
