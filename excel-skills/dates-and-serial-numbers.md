<!-- Creado por Francisco Grandón Vergara -->
# Skill: Dates & Serial Numbers (Fechas y Números Seriales en Excel)

> **Misión:** Dominar la representación interna de fechas y horas en Microsoft Excel, implementar conversiones bidireccionales ultrarrápidas entre cadenas ISO-8601 y números seriales, y navegar con precisión las anomalías históricas (Época 1899 y el bug bisiesto de Lotus 1-2-3).

---

## 1. El Modelo Mental: ¿Cómo Almacena Excel las Fechas?

En Microsoft Excel, **las fechas no existen como tipo de dato nativo independiente**. Una fecha es estrictamente un **número de punto flotante** con una máscara de visualización aplicada (`NumberFormat = "yyyy-mm-dd"`):

* **Parte Entera:** Número de días transcurridos desde la fecha base.
* **Parte Fraccionaria:** Fracción de un día de 24 horas transcurrida (hora, minutos y segundos).
  - `0.0` = 00:00:00 (Medianoche)
  - `0.25` = 06:00:00 (Mañana)
  - `0.5` = 12:00:00 (Mediodía)
  - `0.75` = 18:00:00 (Tarde)

**Ejemplo:**
- `45200` = `2023-10-01`
- `45200.75` = `2023-10-01 18:00:00`

---

## 2. Las Dos Trampas Históricas del Motor de Fechas

### Trampa 1: La Verdadera Época Base (30 de Diciembre de 1899)
Aunque la documentación popular afirma que el día 1 de Excel es el `1 de enero de 1900`, en términos matemáticos modernos la época base computacional para cualquier fecha contemporánea es:
$$\text{Época Base} = \mathbf{1899\text{-}12\text{-}30}$$

### Trampa 2: El Bug Bisiesto de Lotus 1-2-3 (Serial 60)
* En la década de 1980, la hoja de cálculo dominante **Lotus 1-2-3** contenía un error en su algoritmo de calendario: asumió incorrectamente que el año **1900 fue un año bisiesto**.
  - Según el calendario gregoriano, un año secular (divisible por 100) **no** es bisiesto salvo que sea divisible por 400. Por lo tanto, el año 1900 tuvo 365 días; el **29 de febrero de 1900 nunca existió**.
* Cuando Microsoft diseñó Excel en 1985, decidió **replicar deliberadamente este error** para garantizar compatibilidad total con los archivos y fórmulas de Lotus 1-2-3.
* En Excel:
  - Serial `59` = 28 de Febrero de 1900
  - Serial `60` = 29 de Febrero de 1900 *(Fecha inexistente)*
  - Serial `61` = 1 de Marzo de 1900
* **Consecuencia para Agentes:** Para cualquier fecha posterior al **1 de marzo de 1900** (es decir, prácticamente el 100% de los datos de negocio actuales), la fórmula matemática exacta en Python es restar `date(1899, 12, 30)`. Este desfase de un día absorbe perfectamente el bug bisiesto de 1900.

---

## 3. Receta 1: Conversión Bidireccional Ultrarrápida en Python

Esta implementación pura en Python no requiere librerías externas pesadas y ejecuta en microsegundos:

```python
import datetime
from typing import Optional, Union

EXCEL_BASE_DATE = datetime.date(1899, 12, 30)
EXCEL_BASE_DATETIME = datetime.datetime(1899, 12, 30, 0, 0, 0)

def date_to_excel_serial(d: Union[datetime.date, datetime.datetime, str]) -> int:
    """Convierte una fecha a número serial entero de Excel."""
    if isinstance(d, str):
        # Parsear cadena ISO 'YYYY-MM-DD'
        d = datetime.date.fromisoformat(d.strip()[:10])
    elif isinstance(d, datetime.datetime):
        d = d.date()
    return (d - EXCEL_BASE_DATE).days

def datetime_to_excel_serial(dt: datetime.datetime) -> float:
    """Convierte un timestamp con hora a número serial flotante de Excel."""
    delta = dt - EXCEL_BASE_DATETIME
    return delta.days + (delta.seconds + delta.microseconds / 1e6) / 86400.0

def excel_serial_to_date(serial: Union[int, float]) -> datetime.date:
    """Convierte un número serial de Excel a objeto datetime.date nativo."""
    return EXCEL_BASE_DATE + datetime.timedelta(days=int(serial))

def excel_serial_to_iso(serial: Union[int, float]) -> str:
    """Convierte un serial de Excel a cadena ISO-8601 ('YYYY-MM-DD')."""
    return excel_serial_to_date(serial).isoformat()
```

---

## 4. Receta 2: Coerción Robusta con `parse_excel_date_serial`

En `antigravity_excel_core.py`, el método `parse_excel_date_serial` unifica la detección automática de múltiples variantes textuales (fechas con guiones, barras, horas y formatos mixtos):

```python
from antigravity_excel_core import parse_excel_date_serial

# Ejemplos de uso directo:
serial1 = parse_excel_date_serial("2024-05-15")           # 45427
serial2 = parse_excel_date_serial("15/05/2024 14:30:00")  # 45427
serial3 = parse_excel_date_serial(datetime.date(2024, 5, 15)) # 45427
```

### Aplicación en Curaduría por Lote Vía CLI
```bash
# Estandarizar y coercionar fechas a seriales en una tabla completa
python antigravity_excel_cli.py clean \
  --source "financials" \
  --rules-json '{"coerce_dates": ["Date", "Created_At"]}' \
  --sheet "Sheet1"
```

---

## 5. Catálogo de Trampas Comunes (*Traps*) y Remedios

### Trampa 1: Fechas Almacenadas como Texto Plano (Prefijo `'`)
* **Síntoma:** El usuario o una exportación web insertó fechas como cadenas de texto (`'2024-01-15`). Las funciones `=YEAR()`, `=MONTH()` o los filtros de fecha de Excel fallan o no agrupan por año/trimestre.
* **Causa:** La celda tiene formato de texto (`@`) o contiene un apóstrofe inicial invisible.
* **Solución:**
  1. Convertir la cadena a serial entero mediante `parse_excel_date_serial`.
  2. Asignar el valor numérico a la celda.
  3. Establecer la máscara contable de fecha:
  ```python
  cell.Value = 45306  # Serial para 2024-01-15
  cell.NumberFormat = "yyyy-mm-dd"
  ```

### Trampa 2: Ambigüedad Regional entre Día y Mes (`DD/MM` vs `MM/DD`)
* **Síntoma:** `04/05/2024` es interpretado como 5 de abril en EE.UU., pero como 4 de mayo en Chile o Europa.
* **Causa:** Ingesta de archivos CSV sin especificación explícita de formato local.
* **Regla Heurística del Agente:**
  - Si uno de los dos primeros números es $> 12$ (ej. `25/03/2024`), el primer número es inequívocamente el día.
  - Si ambos son $\le 12$ (ej. `01/02/2024`), revisar la columna completa para determinar la convención predominante del dataset antes de transformar.
  - Al generar o inyectar fechas, usar siempre formato ISO canónico `YYYY-MM-DD`.

### Trampa 3: Desfases por Zona Horaria (UTC vs Local)
* **Síntoma:** Fechas que al importarse desde una base de datos o API aparecen desfasadas en un día (ej. `2024-05-14 23:00` en lugar de `2024-05-15 00:00`).
* **Causa:** Excel no almacena información de huso horario (`tzinfo`). Trata todos los números como tiempo local ingenuo (*naive*).
* **Solución:** Convertir los timestamps con zona horaria a hora local del usuario antes de calcular el serial de Excel:
  ```python
  dt_local = dt_utc.astimezone(timezone_local)
  serial = datetime_to_excel_serial(dt_local.replace(tzinfo=None))
  ```

---

## 6. Checklist de Validación Temporal para Agentes

- [ ] ¿Se representaron las fechas como números seriales acompañados de `NumberFormat = "yyyy-mm-dd"` en lugar de strings crudos?
- [ ] ¿Se utilizó la época base `1899-12-30` en lugar de `1900-01-01` para el cálculo?
- [ ] ¿Se eliminaron zonas horarias antes de la conversión a flotante?
- [ ] ¿Se validó que el rango temporal esté dentro del intervalo válido de Excel (`1` a `2958465`, correspondiente al año 1900 hasta el 9999)?
