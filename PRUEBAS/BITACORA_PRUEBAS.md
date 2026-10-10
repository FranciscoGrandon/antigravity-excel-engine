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
### 3. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Método de Agregación Analítica (`aggregate`):**  
   - Integrado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. Permite tabular métricas agregadas (`sum`, `avg`/`mean`, `min`, `max`, `count`) agrupando por cualquier columna categórica y ordenando resultados de forma automática.
   - Cuenta con normalización insensible a espacios y mayúsculas (`strip().lower()`), previniendo errores ante encabezados sucios como `' Sales'`.
2. **Rastreo Automático de Celdas de Citación:**  
   - Cada grupo retornado por `aggregate` reporta explícitamente:
     - `max_cell` y `max_value`: La celda de mayor valor del grupo (ej. `L194` para Government).
     - `min_cell` y `min_value`: La celda de menor valor del grupo (ej. `L694` para Enterprise).
     - `sample_citation_cells`: Muestra representativa de celdas iniciales y finales para trazabilidad de auditoría.
3. **Comando CLI `aggregate`:**  
   - Expuesto en `antigravity_excel_cli.py`:
     ```bash
     python antigravity_excel_cli.py aggregate --group-by "Segment" --metric "Profit" --file "..." --json
     ```
4. **Verificación Automatizada:**  
   - Suite `pytest tests/ -v` con 5/5 pruebas aprobadas al 100%.

---

## Prueba 3: Inyección de Fórmulas Dinámicas (`3_Formulas\Financial_Sample.xlsx`)

### 1. Metodología de Ejecución
- Archivo abierto interactivamente por el usuario en Microsoft Excel.
- Vinculación en vivo a través del canal COM de la ventana interactiva (`XLMAIN` / `EXCEL7`) utilizando `WM_GETOBJECT` y `oleacc.ObjectFromLresult`.
- Inyección de fórmulas dinámicas nativas de Excel con evaluación en tiempo real y propagación sobre la tabla estructurada oficial `financials`.

### 2. Implementación de Nuevas Columnas

#### Columna Q: `Margen %`
- **Encabezado:** Celda **`Q1`** = `"Margen %"`.
- **Fórmula asignada:** **`=L{r}/J{r}`** (Utilidad neta `Profit` dividida por ventas netas ` Sales`).
  - Ejemplo fila 2: **`Q2`** = `=L2/J2` (evaluado en vivo como `50.0%`).
  - Ejemplo fila 701: **`Q701`** = `=L701/J701`.
- **Formato numérico:** Porcentaje con un decimal (`0.0%`).

#### Columna R: `Utilidad por unidad`
- **Encabezado:** Celda **`R1`** = `"Utilidad por unidad"`.
- **Fórmula asignada:** **`=L{r}/E{r}`** (Utilidad neta `Profit` dividida por unidades vendidas `Units Sold`).
  - Ejemplo fila 2: **`R2`** = `=L2/E2` (evaluado en vivo como `$10.00`).
  - Ejemplo fila 701: **`R701`** = `=L701/E701`.
- **Formato numérico:** Formato contable estándar de moneda (`$#,##0.00`).

### 4. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Conector Nativo a Ventanas Interactivas de Excel vía Windows Accessibility API (`oleacc`):**  
   - Integrado en `LiveExcelCOMBackend._try_get_live_desktop_excel`. Resuelve la limitación donde Excel no se publica en la tabla ROT y `GetActiveObject` falla con `0x800401E3`. Localiza la ventana activa del usuario en `WinSta0\Default` y engancha el puntero COM `IDispatch` mediante `oleacc.ObjectFromLresult`.
2. **Método Atómico de Inyección de Columnas Calculadas (`add_calculated_column`):**  
   - Implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. Permite agregar columnas con fórmulas relativas propagadas para todas las filas de datos, aplicar formateo numérico y autoajustar ancho de columna.
3. **Redimensionamiento Automático de Tablas Estructuradas (`ListObjects` / `Table`):**  
   - Al invocar `add_calculated_column` en hojas que contienen tablas oficiales (como `financials`), el motor redimensiona automáticamente el rango de la tabla (ej. `A1:P701` $\rightarrow$ `A1:R701`) incorporando formalmente las nuevas columnas al esquema de datos.
4. **Comando CLI `add-column`:**  
   - Expuesto en `antigravity_excel_cli.py`:
     ```bash
     python antigravity_excel_cli.py add-column --header "Margen %" --formula "=L2/J2" --number-format "0.0%" --file "..." --json
     ```
5. **Verificación Automatizada:**  
   - Suite `pytest tests/ -v` ampliada con `test_add_calculated_column` (6/6 pruebas aprobadas al 100%).

---

## Prueba 4: Creación de Tabla Dinámica y Gráfico Estadístico (`4_Tabla_dinamica\Financial_Sample.xlsx`)

### 1. Metodología de Ejecución
- Archivo abierto interactivamente por el usuario en Microsoft Excel sobre OneDrive / SharePoint.
- Vinculación en vivo a través del backend COM interactivo (`LiveExcelCOMBackend`) enganchado mediante `oleacc`.
- Creación de una nueva pestaña de análisis denominada `Resumen_Pais` con cuadrícula de celdas visible activa (`DisplayGridlines = True`).
- Construcción de un `PivotCache` nativo alimentado desde la tabla estructurada `financials` (`A1:P701`) y despliegue del `PivotTable` (`PT_Utilidad_Pais`) en la celda `A3`.
- Configuración de campos: `Country` como fila (`xlRowField`), `Profit` como valor agregado (`xlSum`) con formato de moneda (`$#,##0.00`) y autoajuste automático de columnas (`AutoFit`).
- Incorporación de un gráfico de columnas agrupadas (`xlColumnClustered`) enlazado a la tabla dinámica, posicionado de manera no invasiva a la derecha (`placement='auto'` $\rightarrow$ `E3`, margen de 2 columnas) para evitar cualquier solapamiento visual con la tabla.

### 2. Resultados Consolidados de la Tabla Dinámica (`Resumen_Pais!A3:B9`)

| País (`Country`) | Utilidad Neta Total (`Sum de Profit`) |
| :--- | :---: |
| **France** | **$3,781,020.78** |
| **Germany** | **$3,680,388.82** |
| **Canada** | **$3,529,228.89** |
| **United States of America** | **$2,995,540.67** |
| **Mexico** | **$2,907,523.11** |
| **Total General** | **$16,893,702.26** |

- **Gráfico Asociado:** Objeto `ChartObject` de tipo columna agrupada con título *"Utilidad Neta por País"*, anclado en `E3` (Left: 330.0 pt, Top: 28.8 pt, Width: 480 pt, Height: 300 pt).

### 3. Conclusión
La generación automatizada de tablas dinámicas y gráficos vinculados en vivo valida la capacidad de análisis multidimensional de Antigravity Excel Engine sin alterar los datos crudos originales. El cálculo geométrico anti-solapamiento garantiza reportes ejecutivos limpios y listos para presentación gerencial, con total paridad entre el entorno interactivo COM y el procesamiento por lotes Headless. La coherencia matemática del 100% frente a la agregación de la Prueba 2 ratifica la precisión analítica y robustez del motor.

### 4. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Helper de Posicionamiento Geométrico Anti-Solapamiento (`_calculate_placement_coordinates`):**  
   - Integrado en `BaseExcelBackend`. Evalúa las dimensiones de la tabla origen; si `placement='auto'`, ubica el elemento a la derecha dejando 2 columnas libres cuando el número de columnas es $\le 8$, o abajo dejando 2 filas libres cuando es $> 8$. Admite también posicionamiento explícito `'right'`, `'bottom'` o anclaje a celda personalizada (ej. `'I3'`).
2. **Método Atómico de Tabla Dinámica (`create_pivot_table`):**  
   - Implementado en `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`:
     - En **Live COM**: crea el `PivotCache` (tipo `xlDatabase`), instancia la tabla dinámica en la hoja destino (creándola automáticamente si no existe), asigna campos a filas y columnas, inyecta el campo de valor con función de agregación (`sum`, `count`, `average`/`avg`, `min`, `max`), aplica máscara numérica y asegura `DisplayGridlines` y `AutoFit`.
     - En **Headless OpenPyXL**: procesa los datos en memoria para generar una tabla cruzada (*cross-tabulation*) de doble entrada con filas, columnas, subtotales, fila/columna de Total General, bordes corporativos, tipografía Segoe UI y formato numérico.
3. **Método Atómico de Gráficos No Invasivos (`create_chart`):**  
   - Implementado en `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`:
     - En **Live COM**: crea un `ChartObject` en las coordenadas geométricas calculadas para evitar solapamientos, enlaza la fuente (`SetSourceData`), configura el tipo (`column_clustered`, `bar_clustered`, `line`, `pie`, `area`) y establece el título.
     - En **Headless OpenPyXL**: inserta el gráfico correspondiente (`openpyxl.chart.BarChart`, `LineChart`, etc.) excluyendo los totales generales para mantener la escala adecuada de la serie, y lo ancla en la celda calculada.
4. **Comandos CLI `create-pivot` y `create-chart`:**  
   - Expuestos en `antigravity_excel_cli.py`:
     ```bash
     python antigravity_excel_cli.py create-pivot --source "financials" --rows Country --value-field Profit --value-func sum --value-format "$#,##0.00" --dest-sheet "Resumen_Pais" --dest-cell "A3" --table-name "PT_Utilidad_Pais" --file "..." --json
     python antigravity_excel_cli.py create-chart --source "PT_Utilidad_Pais" --chart-type "column_clustered" --title "Utilidad Neta por País" --dest-sheet "Resumen_Pais" --placement "auto" --file "..." --json
     ```
5. **Verificación Automatizada:**  
   - Suite `pytest tests/ -v` ampliada con `test_calculate_placement_coordinates`, `test_create_pivot_table_headless` y `test_create_chart_headless` (9/9 pruebas aprobadas al 100%).

---

## Prueba 5: Simulación de Escenarios y Modelado Comparativo

### 1. Metodología de Ejecución
- Creación de un modelo de simulación y comparativa financiera profesional basado en parámetros y supuestos macroeconómicos / de negocio vinculados.
- Despliegue de pestaña analítica ejecutiva con visualización de cuadrícula activa (`showGridLines = True` en OpenPyXL / `DisplayGridlines = True` en Live COM).
- Inyección de bloque declarativo de supuestos y parámetros clave de entrada con tipografía institucional y celdas estilizadas bajo el preset `input_cell` (fondo amarillo marfil suave `#FFF2CC`, texto azul `#1F4E79`, borde medio y alineación centrada).
- Estructuración de tabla resumen comparativa multidimensional con encabezados corporativos bajo preset `header` (fondo azul marino corporativo `#0E2E63`, texto blanco `#FFFFFF`, negrita y centrado vertical).
- Incorporación de filas de datos y escenarios proyectados con formatos contables estándar (`FinancialStyleGuide.CURRENCY_FORMAT` y `PERCENT_FORMAT`), y fila de totales con subrayado contable doble bajo preset `total_row` (borde superior simple, borde inferior doble contable y negrita).
- Protección estricta de referencias absolutas vs. relativas en propagación de fórmulas vectorizadas mediante el nuevo motor `propagate_formula_row`.

### 2. Especificación del Modelo de Simulación
- **Hoja Destino:** `Simulacion_Escenarios`
- **Bloque de Supuestos y Parámetros (`B2:C4`):**
  - **Tasa de Crecimiento Proyectada:** `C3` = `5.0%` (preset `input_cell`, formato `percent`).
  - **Inflación / Costo de Fondos:** `C4` = `3.0%` (preset `input_cell`, formato `percent`).
- **Tabla Comparativa de Escenarios (`B6:F10`):**
  - **Título:** *"Resumen Comparativo de Escenarios Proyectados"* (celda `B6`, Segoe UI 13 pt bold, color `#0E2E63`).
  - **Encabezados (`B7:F7`):** `["Escenario", "Ingresos Proyectados", "Costos Estimados", "Margen Operacional", "Rentabilidad %"]` (preset `header`).
  - **Escenario Base (`B8:F8`):** Ingresos $100,000.00, Costos $70,000.00, Margen `=C8-D8` ($30,000.00), Rentabilidad `=E8/C8` (30.0%).
  - **Escenario Optimista (`B9:F9`):** Ingresos $120,000.00, Costos $80,000.00, Margen `=C9-D9` ($40,000.00), Rentabilidad `=E9/C9` (33.3%).
  - **Total / Consolidado (`B10:F10`):** Ingresos `=SUM(C8:C9)`, Costos `=SUM(D8:D9)`, Margen `=C10-D10`, Rentabilidad `=E10/C10` (preset `total_row`).

### 3. Conclusión
La incorporación del sistema universal de presets estilísticos y el motor genérico de tablas resumen permite a los agentes de IA construir modelos financieros y cuadros de mando comparativos de alta calidad gráfica y consistencia matemática. El soporte dual (Live COM y Headless OpenPyXL) garantiza que cualquier modelo diseñado por el motor mantenga una presentación ejecutiva impecable sin requerir retoques manuales por parte del usuario.

### 4. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Sistema Declarativo Universal de Presets de Estilo (`apply_style_preset`):**  
   - Integrado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`.
   - Soporta 7 presets canónicos:
     - `header`: Fondo azul corporativo `#0E2E63`, texto blanco, negrita, centrado vertical.
     - `input_cell`: Fondo amarillo marfil `#FFF2CC`, texto azul `#1F4E79`, borde medio, centrado horizontal.
     - `total_row`: Negrita, borde inferior doble contable (`double` / `xlDouble`), borde superior simple.
     - `currency`: Máscara contable oficial `FinancialStyleGuide.CURRENCY_FORMAT` (`$ #,##0.00;($ #,##0.00);"-"`).
     - `percent`: Formato porcentual oficial `FinancialStyleGuide.PERCENT_FORMAT` (`0.0%`).
     - `delta_positive`: Texto verde corporativo `#008000`, negrita.
     - `delta_negative`: Texto rojo advertencia `#C00000`, negrita.
2. **Motor Genérico de Tablas de Resumen / Comparativas con Parámetros Vinculados (`create_summary_table`):**  
   - Integrado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`.
   - Crea automáticamente la hoja de destino, activa la cuadrícula (`DisplayGridlines`), genera el bloque de supuestos/palancas estilizado como celdas de entrada, aplica tipografía de destaque al título general, renderiza encabezados con preset `header`, escribe y formatea filas de datos, aplica preset `total_row` en filas de totales y ejecuta `AutoFit` en las columnas involucradas.
3. **Protección de Referencias Absolutas vs. Relativas en Propagación de Fórmulas (`propagate_formula_row`):**  
   - Función auxiliar en `antigravity_excel_core.py` integrada en `add_calculated_column`.
   - Detecta si las plantillas contienen referencias absolutas (`$A$1`, `'Params'!$C$4`, `$C$2`, etc.) y las mantiene estrictamente inalteradas en todas las filas de datos, al tiempo que sustituye dinámicamente las referencias relativas (`A2`, `$B2`, `{r}`) según el índice de fila de destino.
4. **Exposición en CLI Oficial (`apply-preset` y `summary-table`):**  
   - Nuevos subcomandos en `antigravity_excel_cli.py`:
     ```bash
     python antigravity_excel_cli.py apply-preset --range "B2:F2" --preset "header" --file "data.xlsx" --json
     python antigravity_excel_cli.py summary-table --spec-json @spec.json --file "data.xlsx" --json
     ```
5. **Verificación Automatizada:**  
   - Cobertura de tests unitarios ampliada en `tests/test_excel_engine.py` con `test_apply_style_preset_headless`, `test_create_summary_table_headless` y `test_add_calculated_column_absolute_ref` (14/14 pruebas aprobadas al 100%).

---

## Prueba 6: Limpieza, Estandarización y Outliers (`6_Limpieza\Financial_Sample_SUCIO.xlsx`)

### 1. Metodología de Ejecución
- Archivo de prueba con datos degradados y anomalías inyectadas sobre la tabla estructurada `financials` (`A1:P716`).
- Lectura resiliente no invasiva en modo headless/live sin alterar el archivo bloqueado.
- Detección previa de anomalías mediante `audit_data_quality`:
  * 15 filas duplicadas inyectadas al final de la tabla (filas 702 a 716).
  * Encabezado con espacio en blanco: `' Sales'` en la columna J (`J1`).
  * Inconsistencias de tipo: valores numéricos almacenados como texto en `Units Sold` (ej. `'663'`, `'2905'`) y en `Year` (`'2014'`, `'2013'`).
  * Formatos de fecha heterogéneos y fechas representadas como texto ('YYYY-MM-DD', 'DD/MM/YYYY con marcas de hora').
  * Valores atípicos extremos inyectados en `Units Sold` (`999999.0` en filas 325, 412, 689).
- Ejecución estricta **ON-DEMAND** de los módulos de curación: ninguna operación destructiva o de limpieza se ejecuta de forma implícita o automática al abrir o consultar libros.

### 2. Resultados de las Mejoras Implementadas

#### A. Deduplicación Inteligente On-Demand (`remove_duplicates`)
- Se ejecutó `remove_duplicates` sobre la tabla estructurada `financials`:
  * Total de filas evaluadas: 715 filas de datos (filas 2 a 716).
  * Duplicados exactos detectados y eliminados: **15 filas** (filas 702 a 716).
  * Filas resultantes: **700 filas de datos** (dimensiones exactas del dataset original `Financial_Sample.xlsx`).
  * Redimensionamiento automático de la tabla estructurada `financials`: `A1:P716` ajustado limpiamente a `A1:P701` (`tbl.ref` en OpenPyXL y `tbl.Resize` en Live COM), sin dejar filas vacías residuales ni corromper rangos.

#### B. Detección y Marcado Visual No Destructivo de Outliers (`detect_and_flag_outliers`)
- Ejecución sobre `Units Sold` y `Profit` utilizando metodología estadística dual (IQR y Z-Score):
  * **IQR en Units Sold:** Cuartiles calculados Q1 = 896.5, Q3 = 2258.0, IQR = 1361.5. Límites calculados: inferior = -1145.75, superior = 4300.25.
  * **Valores anómalos detectados:** 4 registros, incluyendo los 3 outliers artificiales (`999999.0` en `E325`, `E412`, `E689`) y el valor extremo superior legítimo (`4492.5` en `E77`).
  * **Preservación de Integridad (No Destructivo):** Los valores originales permanecen intactos en sus celdas. Se aplica estilización visual suave de advertencia (fondo `#FCE4D6` y tipografía roja oscura `#C00000` negrita) y comentarios de celda explicativos con los límites calculados.
  * **Z-score Adaptativo:** Implementación de umbral adaptativo $z = 2.5$ para muestras de tamaño reducido ($N < 30$) y $z = 3.0$ para muestras poblacionales mayores, garantizando rigor matemático ante muestras de cualquier tamaño.

#### C. Coerción Universal de Fechas a Número de Serie de Excel (`parse_excel_date_serial`)
- Soporte para cadenas de fecha en formato ISO ('YYYY-MM-DD', 'YYYY/MM/DD', 'YYYY-MM-DDTHH:MM:SS'), fechas latinoamericanas/europeas con slashes ('DD/MM/YYYY', 'MM/DD/YYYY') con o sin marcas de hora ('3:00:00 AM', '15:30:00') y objetos nativos `datetime.date` / `datetime.datetime`.
- Conversión exacta a días transcurridos desde `1899-12-30` (época oficial de Excel 1900 con compensación de año bisiesto): ej. `2014-01-01` $\rightarrow$ `41640`, `2013-09-15` $\rightarrow$ `41532`.

#### D. Limpieza y Estandarización en Bloque (`clean_table_dataset`)
- Procesamiento en memoria en una sola pasada de alto rendimiento dentro del context manager `excel_fast_mode`:
  * `trim_headers`: Normalización del encabezado `' Sales'` $\rightarrow$ `'Sales'`.
  * `standardize_text`: Eliminación de espacios residuales (trim), unificación de casing (`title`, `upper`, `lower`) y mapeo declarativo de sinónimos/reemplazos.
  * `cast_types`: Conversión de números almacenados como texto a tipos primitivos nativos (`float`, `int`).
  * `coerce_dates`: Conversión de strings de fecha a números de serie de Excel con máscara de formato uniforme (`yyyy-mm-dd`).
- Volcado masivo y atómico en bloque a la hoja, optimizando los tiempos de procesamiento en órdenes de magnitud.

### 3. Conclusión
Las capacidades de curación, deduplicación y detección de anomalías bajo demanda dotan a Antigravity Excel Engine de herramientas de higiene de datos profesionales, listas para pipelines de datos automatizados y agentes de IA. Se mantiene la premisa de máxima velocidad en operaciones matriciales y respeto irrestricto a la no-destructividad de los datos del usuario.

### 4. Mejoras de Ingeniería Implementadas en EXCEL_ENGINE
1. **Conversión Universal de Fechas a Seriales de Excel (`parse_excel_date_serial`):**  
   - Función auxiliar en `antigravity_excel_core.py` que interpreta formatos ISO, DD/MM/YYYY, marcas de hora y datetimes, transformándolos a enteros seriales exactos respecto al 30 de diciembre de 1899.
2. **Deduplicación Inteligente On-Demand (`remove_duplicates`):**  
   - Implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. Soporta filtrado por columnas clave (`key_columns`) o fila completa, eliminación ordenada y redimensionamiento de tablas estructuradas (`ListObject.Resize` y `tbl.ref`).
3. **Detección y Señalización Visual No Destructiva de Outliers (`detect_and_flag_outliers`):**  
   - Métodos IQR y Z-score adaptativo implementados en ambos backends. Aplica estilo visual `#FCE4D6` con texto rojo bold y comentarios explicativos sin eliminar datos.
4. **Motor de Estandarización y Limpieza de Datasets (`clean_table_dataset`):**  
   - Operación en una sola pasada matricial en memoria que ejecuta `standardize_text`, `cast_types`, `coerce_dates` y `trim_headers` bajo `excel_fast_mode`.
---

## Hito Especial: Optimización Integral de Velocidad y Arquitectura de Máximo Rendimiento

### 1. Diagnóstico de Latencia y Metodología
- Se identificó que el 96% de la latencia en respuestas interactivas provenía del arranque en frío de Python en Windows (~400 ms) y la importación estática masiva de módulos pesados (`openpyxl`, `win32gui`, etc. ~726 ms), antes de interactuar con Excel.
- Se implementó un plan de optimización de tres capas bajo el marco **Teamwork Custom (Modo Benchmark)**:
  1. **Lazy Loading de Módulos:** Eliminación de imports ansiosos en la raíz de `antigravity_excel_core.py` y `antigravity_excel_cli.py`, con utilidades de coordenadas en Python puro y resolución diferida con cacheo singleton.
  2. **Pipeline Unificado de Curación (`curate_pipeline`):** Ejecución encadenada en memoria (deduplicación, estandarización de texto, coerción de fechas, casteo de tipos y cálculo estadístico de outliers) en un único viaje COM de lectura y escritura matricial.
  3. **Vectorización Multi-Rango No Contigua:** Formateo y marcado de múltiples celdas disjuntas (ej. `E325,E412,E689`) en llamadas COM combinadas mediante fragmentación de cadenas seguras (`_chunk_cell_addresses`), eliminando bucles celda por celda.
  4. **Servidor Daemon Residente en Memoria (`antigravity_excel_daemon.py`):** Proceso en segundo plano conectado a un Named Pipe nativo de Windows (`\\.\pipe\antigravity_excel`) con búferes bidireccionales de 16 MB. Mantiene la instancia COM en caliente y responde solicitudes en < 20 ms.
  5. **Cliente Rápido con Fallback Transparente en CLI:** `antigravity_excel_cli.py` intenta comunicarse con el pipe atómicamente; si el daemon está en ejecución, responde de inmediato; si está apagado, conmuta de forma transparente y sin fallar al modo standalone en < 1 ms.

### 2. Resultados Cuantitativos del Benchmark (`tests/benchmark_speed.py`)

| Métrica de Desempeño | Antes (Baseline) | Después (Optimizado) | Factor de Mejora |
| :--- | :---: | :---: | :---: |
| **Import `antigravity_excel_core` en frío** | ~726,0 ms | **61,06 ms** | **11,9x más rápido (91,6% reducción)** |
| **Import `antigravity_excel_cli` en frío** | ~379,0 ms | **4,45 ms** | **85x más rápido (98,8% reducción)** |
| **Invocación CLI Help / Parseo** | ~817,0 ms | **~200,0 ms** | **4,1x más rápido** |
| **Latencia IPC Named Pipe (Ping)** | *No existía* | **~5,25 ms** | **Instantáneo** |
| **Consulta de Estado (`status`) vía Daemon** | ~820,0 ms | **~129,0 ms** | **6,3x más rápido** |
| **Curación Integral de Datos (700 filas)** | ~4.700,0 ms | **~500,0 ms** | **9,4x más rápido** |

### 3. Certificación de la Suite de Pruebas Automatizadas
- Suite `pytest tests/ -v` ampliada con `tests/test_daemon_pipe.py` (8 nuevos tests de IPC y resiliencia de fallback).
- **Total Suite:** **27 pruebas ejecutadas, 27 aprobadas al 100% en 1,06 segundos** (código de salida 0).

---

## Prueba 7: Auditoría de Fórmulas y Consistencia de Modelos (`7_Auditoria\Financial_Sample_AUDITORIA.xlsx`)

### 1. Metodología de Acceso e Inspección
- El libro se encontraba abierto en Microsoft Excel (`Live COM`).
- Se inspeccionaron de forma no destructiva tanto la hoja `Resumen` como la hoja `Datos` (701 filas x 16 columnas), verificando correspondencia entre fórmulas y valores, integridad de rangos de agregación y consistencia entre tablas cruzadas.

### 2. Hallazgos y Causa Raíz Detectada
1. **Resumen!C10 (Ventas por Segmento):** Rango truncado `=SUM(C5:C8)` que omitía la fila 9 (`Small Business`, $42.42M).
2. **Resumen!D17 (COGS México):** Factor de escala erróneo `/1000` en fórmula `=SUMIFS(...)/1000`, subestimando el costo de $18.04M a $18.04K.
3. **Datos!H602 (Gross Sales):** Fórmula desfasada `=E601*G602` multiplicando por unidades de la fila anterior en lugar de `=E602*G602`.
4. **Datos!J455 (Sales Netas):** Signo aritmético invertido `=H455+I455` (sumaba el descuento en vez de restarlo).
5. **Datos!L205 (Profit):** Valor numérico estático incrustado (`181997.2`) en vez de la fórmula dinámica `=J205-K205`.

### 3. Conclusión
La auditoría profunda de fórmulas permite aislar distorsiones tanto en capas de presentación agregada (hojas de resumen) como en registros atómicos de datos fuente. Tras aplicar las 5 correcciones propuestas, ambas tablas de la hoja `Resumen` quedaron matemática y financieramente cuadradas al 100% (Ventas: $118.73M, COGS: $101.83M, Utilidad: $16.89M, Margen: 14.23%).









