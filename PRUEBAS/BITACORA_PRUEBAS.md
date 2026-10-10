# Bitácora de Pruebas: Antigravity Excel Engine

> **Fecha:** 10 de Octubre de 2026  
> **Directorio de Pruebas:** `D:\Users\francisco.grandon\OneDrive - essbio\Documentos\ANTIGRAVITY\EXCEL_ENGINE\PRUEBAS`  
> **Restricción de Integridad:** Directorio `CLAVES` estrictamente excluido y no consultado.

---

## Prueba 1: Comprensión del Libro (`1_Comprension\Financial_Sample.xlsx`)

### 1. Metodología de Acceso
- El archivo se encontraba abierto en Microsoft Excel por el usuario con bloqueo exclusivo de escritura (`[Errno 13] Permission denied`).
- Se implementó lectura compartida Win32 (`CreateFile` con flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`) hacia buffer en memoria para permitir auditoría concurrente sin alterar la sesión viva de Excel del usuario.

### 2. Estructura General
- **Hojas:** 1 hoja (`Sheet1`).
- **Tabla Estructurada:** Rango `A1:P701` nombrado `financials`.
- **Dimensiones:** 701 filas (1 fila de encabezado + 700 registros de datos) y 16 columnas (`A` a `P`).

### 3. Columnas y Tipos de Datos Detectados
1. **A (`Segment`):** Texto (`str`). Valores categóricos: *Government, Enterprise, Midmarket, Small Business, Channel Partners*.
2. **B (`Country`):** Texto (`str`). Países: *Canada, Germany, France, Mexico, United States of America*.
3. **C (`Product`):** Texto (`str`). Productos: *Carretera, Montana, Paseo, Velo, VTT, Amarilla*.
4. **D (`Discount Band`):** Texto (`str`). Niveles: *None, Low, Medium, High*.
5. **E (`Units Sold`):** Numérico (`float` / `int`). Unidades vendidas.
6. **F (`Manufacturing Price`):** Numérico entero (`int`). Formato Contabilidad (`$`).
7. **G (`Sale Price`):** Numérico entero (`int`). Formato Contabilidad (`$`).
8. **H (`Gross Sales`):** Numérico (`float` / `int`). Ventas brutas.
9. **I (`Discounts`):** Numérico (`float` / `int`). Descuentos.
10. **J (` Sales`):** Numérico (`float` / `int`). Ventas netas (nota: espacio inicial en el encabezado).
11. **K (`COGS`):** Numérico (`float` / `int`). Costo de bienes vendidos.
12. **L (`Profit`):** Numérico (`float` / `int`). Ganancia neta (contiene valores negativos).
13. **M (`Date`):** Fecha (`datetime`). Formato `mm-dd-yy`.
14. **N (`Month Number`):** Numérico entero (`int`, rango 1 a 12).
15. **O (`Month Name`):** Texto (`str`). Nombre en inglés del mes.
16. **P (`Year`):** Texto (`str`). Almacenado como texto en lugar de número ('2013', '2014').

### 4. Rango de Fechas
- **Fecha Inicial:** `01/09/2013` (celdas con fecha mínima, ej. `M6`, `M7`).
- **Fecha Final:** `01/12/2014` (celdas con fecha máxima, ej. `M18`, `M19`).
- **Particularidad temporal:** Todas las fechas del dataset corresponden al primer día de cada mes (`01/MM/AAAA`).

### 5. Hallazgos y Problemas de Calidad de Datos
1. **Espacio en blanco en encabezado (Columna J):**  
   - Celda `J1` contiene `' Sales'` (con espacio al inicio) en lugar de `'Sales'`. Causa habitual de fallos en joins de bases de datos y scripts de análisis.
2. **Tipo de dato inconsistente en Año (Columna P):**  
   - Celdas `P2:P701` contienen texto con formato `@` (ej. `'2014'`, `'2013'`) en vez de número entero, a diferencia de `Month Number` (`N2:N701`) que sí es entero.
3. **Unidades vendidas fraccionarias (Columna E):**  
   - 36 registros presentan decimales `.5` en bienes manufacturados físicos (ej. `E2` = `1618.5`, `E13` = `2665.5`, `E49` = `4219.5`, `E72` = `1375.5`, `E77` = `4492.5`).
4. **Formato numérico contable aplicado erróneamente a columnas de texto:**  
   - Las columnas de texto `Product` (`C2:C701`), `Discount Band` (`D2:D701`) y `Month Name` (`O2:O701`) tienen asignada la máscara contable de moneda `_("$"* #,##0.00_)`.
5. **Pérdidas operacionales (Margen negativo en Columna L):**  
   - 58 registros presentan márgenes negativos por descuentos excesivos donde COGS supera a Sales (ej. `L234` = `-4533.75`, `L239` = `-3740.00`, `L240` = `-2981.25`, `L248` = `-1076.25`, `L249` = `-880.00`).
6. **Ausencia de fórmulas nativas dinámicas:**  
   - Las 700 filas contienen valores calculados estáticos pre-renderizados; no hay fórmulas nativas vivas (`=SUM`, `=PRODUCT`, etc.) en las columnas `H`, `J`, `K` y `L`.

### 6. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Lectura Resiliente Concurrente (`HeadlessOpenPyXLBackend._load_workbook_resilient`):**  
   - Resuelve el error `PermissionError [Errno 13]` cuando el usuario tiene el archivo abierto en Excel. Utiliza `win32file.CreateFile` con flags de compartición total (`FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`) y carga el flujo de bytes a memoria con `io.BytesIO`.
2. **Método de Auditoría Automatizada de Calidad (`audit_data_quality`):**  
   - Implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. Detecta en una sola llamada:
     - Espacios residuales en nombres de encabezados (`header_whitespace_issues`).
     - Celdas numéricas almacenadas erróneamente como texto (`numbers_stored_as_text`).
     - Inconsistencia de tipos de datos en una misma columna (`inconsistent_column_types`).
     - Existencia y conteo de fórmulas dinámicas (`has_dynamic_formulas`, `formula_count`).
3. **Comando CLI `audit-quality`:**  
   - Nuevo comando expuesto en `antigravity_excel_cli.py` con soporte nativo de flags `--range`, `--sheet`, `--file` y `--json`.
4. **Normalización de Argumentos en CLI:**  
   - Argumentos globales (`--file`, `--mode`, `--json`) unificados en `parent_parser`, permitiendo su posición antes o después del subcomando.
5. **Verificación Automatizada:**  
   - Suite `pytest tests/ -v` ampliada y validada al 100% de aprobación (4/4 tests pasando).

---

## Prueba 2: Preguntas con Cita (`2_Preguntas_con_cita\Financial_Sample.xlsx`)

### 1. Metodología de Ejecución
- Archivo abierto interactivamente por el usuario.
- Acceso headless y extracción analítica mediante `AntigravityExcelEngine` con backend resiliente (sin bloqueo ni error COM).
- Agrupación agregada y ordenamiento de utilidad (`Profit`, Columna L) sobre el rango `L2:L701` cruzado con `Segment` (Columna A) y `Country` (Columna B).

### 2. Resultados Consolidados de Utilidad

#### A. Por Segmento (Columna A vs Columna L)
- **Segmento que genera MÁS utilidad:** **`Government`**  
  - **Utilidad Total:** **$11,388,173.17 USD** (distribuida en 300 transacciones).
  - **Celdas de sustento:**
    - Celdas con mayores utilidades individuales del segmento: **`L194`** ($262,200.00), **`L47`** ($247,500.00), **`L126`** ($246,178.00).
    - Celdas iniciales del segmento: **`L2`** ($16,185.00), **`L3`** ($13,210.00), **`L7`** ($136,170.00).
    - Celdas finales del segmento: **`L696`** ($2,051.00), **`L699`** ($1,299.60), **`L700`** ($686.85).
- **Segmento que genera MENOS utilidad:** **`Enterprise`**  
  - **Utilidad Total:** **-$614,545.62 USD** (pérdida neta acumulada en 100 transacciones).
  - **Celdas de sustento:**
    - Celdas con mayores pérdidas (mínima utilidad) del segmento: **`L694`** (-$40,617.50), **`L669`** (-$38,046.25), **`L661`** (-$35,550.00).
    - Celdas iniciales del segmento: **`L13`** ($13,327.50), **`L16`** ($1,725.00), **`L34`** ($9,020.00).
    - Celdas finales del segmento: **`L688`** (-$33,522.50), **`L694`** (-$40,617.50), **`L695`** (-$7,590.00).

#### B. Por País (Columna B vs Columna L)
- **País que genera MÁS utilidad:** **`France`**  
  - **Utilidad Total:** **$3,781,020.78 USD** (en 140 transacciones).
  - **Celdas de sustento:**
    - Celdas con mayores utilidades individuales del país: **`L47`** ($247,500.00), **`L355`** ($188,378.00), **`L69`** ($186,407.50).
    - Celdas iniciales del país: **`L4`** ($10,890.00), **`L10`** ($18,990.00), **`L24`** ($2,745.00).
    - Celdas finales del país: **`L686`** (-$9,116.25), **`L696`** ($2,051.00), **`L697`** ($12,375.00).
- **País que genera MENOS utilidad:** **`Mexico`**  
  - **Utilidad Total:** **$2,907,523.11 USD** (en 140 transacciones).
  - **Celdas de sustento:**
    - Celdas con mayores pérdidas (mínima utilidad) del país: **`L633`** (-$35,262.50), **`L540`** (-$21,560.00), **`L550`** (-$21,560.00).
    - Celdas iniciales del país: **`L6`** ($12,350.00), **`L12`** ($12,350.00), **`L14`** ($47,900.00).
    - Celdas finales del país: **`L684`** ($3,600.00), **`L698`** ($2,730.00), **`L699`** ($1,299.60).


