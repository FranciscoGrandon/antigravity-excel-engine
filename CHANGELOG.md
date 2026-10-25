<!-- Creado por Francisco Grandón Vergara -->
# Changelog

Todas las modificaciones notables de este proyecto serán documentadas en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.1.0] - 2026-10-25

### Añadido
- **Optimización de Rendimiento, Precisión Numérica y Fórmulas Estructuradas:**
  - **Pruebas Automatizadas de Precisión y Performance (`tests/test_copilot_features.py`):**
    - 6 nuevas pruebas en `TestPerformanceAndPrecisionOptimizations` validando agregación de columnas completas, protección de cadenas con comillas, sintaxis formal `[#This Row]`, inferencia de textos numéricos/RUTs, transposición y muestreo matricial, y precisión IEEE 754 de `math.fsum`.
    - Total de la suite ampliado a 137 pruebas unitarias aprobadas al 100%.
  - **Función de Métrica Precisa (`operations/aggregate.py`):**
    - `compute_aggregation_metric(values, agg_func)`: agregación precisa con `math.fsum` sanitizando `None`, vacíos o valores no numéricos.
    - Exportación canónica en `antigravity_excel_engine/__init__.py` y `antigravity_excel_engine/operations/__init__.py`.
  - **Muestreo Inteligente de Esquemas en AI Grounding (`operations/grounding.py`):**
    - Parámetro opcional `max_sample_rows: int = 1000` en `inspect_table_schema`, `get_ai_grounding_context` y fachada `AntigravityExcelEngine`.

### Cambiado
- **Transpilador de Fórmulas Estructuradas (`formulas/structured.py`):**
  - Protección estricta de literales de texto entre comillas dobles, impidiendo sustituciones accidentales dentro de cadenas.
  - Soporte de sintaxis formal de fila actual: `[@Col]`, `[@[Col con espacios]]`, `Tabla[@Col]`, `Tabla[@[Col con espacios]]` y sintaxis estándar `Tabla[[#This Row],[Col]]` o `[[#This Row],[Col]]`.
  - Referencias a columnas completas para agregación: `=SUM([Ventas])` o `=AVERAGE(Tabla[Precio])` con resolución de rangos verticales vía `row_bounds` o `table_range` (ej. `B$2:B$50`) o formato de columna completa `B:B`.
- **Rendimiento Matricial en AI Grounding (`operations/grounding.py`):**
  - Transposición matricial vectorizada en un solo paso con `zip(*data_rows)` en `inspect_table_schema`, eliminando bucles celda a celda $C \times N$.
  - Muestreo uniforme/inicial para tablas extensas que exceden `max_sample_rows`, manteniendo el uso de memoria y CPU acotados.
- **Precisión Aritmética de Agregación (`operations/aggregate.py`, `backends/headless.py`, `backends/live_com.py`):**
  - Reemplazo de `sum()` por `math.fsum` en el cálculo de `sum` y `avg` tanto en operaciones como en backends, erradicando imprecisiones por representación de coma flotante IEEE 754.
- **Rendimiento GUI Live COM (`backends/live_com.py`):**
  - Envolvimiento de `create_table` y `resize_table` con el context manager `with excel_fast_mode(self.app):` para suprimir el repintado de pantalla y acelerar la creación de tablas estructuradas.
- **Latencia CLI (`cli.py`):**
  - Valor por defecto de timeout en `_try_daemon` y en el despacho de rutas no-UNC establecido en `DEFAULT_TIMEOUT_MS` (1000 ms).

### Corregido
- **Velocidad de la Suite de Pruebas (`tests/test_introspection_and_charts.py`):**
  - Corrección del mock de `sys.exit` con `side_effect=SystemExit` en `test_cli_unc_path_sets_60s_timeout_for_daemon`, eliminando la ejecución espuria del fallback standalone sobre rutas UNC y reduciendo la duración del test de 10.27s a 0.00s.
  - Tiempo total de ejecución de la suite de pruebas completa reducido de 12.83s a 2.20s (< 3.5 segundos requeridos).
- **Inferencia Semántica de Tipos (`operations/grounding.py`):**
  - Identificación rigurosa de cadenas numéricas con ceros a la izquierda (`"00123"`, `"0987"`) y valores con guiones como RUTs o teléfonos (`"12345678-9"`) clasificándolos como `text` en lugar de `integer`.


### Añadido
- **Capacidades Copilot en Antigravity Excel Engine:**
  - **AI Grounding & Table Schemas (`operations/grounding.py`):**
    - `inspect_table_schema(backend, table_or_range, sheet=None)`: inspección exhaustiva de metadatos de tabla/rango, dimensiones y columnas con inferencia automática de tipos (`integer`, `float`, `date`, `currency`, `boolean`, `text`), conteo de no nulos, nulos, muestra de 3 valores representativos, rango min/max y detección de fórmulas.
    - `get_ai_grounding_context(backend, table_or_range, sheet=None, format="json"|"markdown")`: serialización ultraligera estructurada en JSON o tablas Markdown optimizada para inyección directa en prompts de LLMs con costo mínimo de tokens.
  - **Gestión de Tablas de Excel `ListObject` (`operations/tables.py`):**
    - `create_table(backend, range_str, table_name=None, style="TableStyleMedium9", has_headers=True, sheet=None)`: conversión oficial de rangos libres en tablas estructuradas tanto en Live COM (`ListObjects.Add`) como en Headless OpenPyXL (`openpyxl.worksheet.table.Table` y `ws.add_table()`), con población automática de `tableColumns`.
    - `list_tables(backend, sheet=None)`: listado de todas las tablas estructuradas con nombre, rango, hoja y nombres de columnas.
    - `resize_table(backend, table_name, new_range, sheet=None)`: redimensionamiento de tablas estructuradas con sincronización de columnas.
  - **Transpilador de Fórmulas Estructuradas (`formulas/structured.py`):**
    - `transpile_structured_formula(formula, headers_map, row_idx)`: traducción de sintaxis de referencias estructuradas de tablas como `=[@Ventas] - [@Costos]` o `=[@[Precio Unitario]] * [@Cantidad]` a coordenadas relativas como `=B2 - C2`.
    - Integración nativa en `add_calculated_column` en Headless OpenPyXL y Live COM para expansión automática de columnas con sintaxis estructurada.
  - **Modo Dry-Run / Previsualización Especulativa:**
    - Parámetro opcional `dry_run: bool = False` en `set_cells`, `set_cell_range` y `add_calculated_column`.
    - Simulación segura sin alterar celdas ni libros en memoria ni en disco, retornando `dry_run=True`, `affected_cells_count`, `target_range` y `sample_diff` de celdas afectadas con valores anteriores, nuevos y fórmulas.
  - **Integración en Facade, CLI e IPC:**
    - Exposición de métodos en `BaseExcelBackend`, `LiveExcelCOMBackend`, `HeadlessOpenPyXLBackend` y la fachada `AntigravityExcelEngine`.
    - Re-exportación completa en `antigravity_excel_engine/__init__.py` y `antigravity_excel_engine/operations/__init__.py`.
    - Soporte en CLI (`cli.py`) con comandos `inspect-table`, `create-table` y bandera `--dry-run` en `set-range`, `set-cells` y `add-column`.
    - Soporte en protocolo IPC (`ipc/protocol.py`) y despachador del daemon (`ipc/daemon.py`).
  - **Suite de Pruebas Unitarias (`tests/test_copilot_features.py`):**
    - 23 nuevas pruebas unitarias cubriendo transpilación de fórmulas, creación/listado/redimensionado de tablas, schemas semánticos, grounding en JSON/Markdown, inyección calculada estructurada, dry-run y mocks de Live COM.
    - 131/131 pruebas aprobadas al 100% en pytest (código de salida 0).
- **Fases 1 y 2 de Optimización y Rendimiento de Excel Engine (`plan_mejora_performance_excel_engine.md`):**
  - **Blindaje y Ciclo de Vida en `LiveExcelCOMBackend` (`backends/live_com.py`):**
    - Preservación íntegra de rutas UNC de red (`\\...`) sin manipulación indebida de letras de unidad en `_init_workbook`.
    - Apertura segura con parámetros explícitos `Workbooks.Open(abs_p, 0, False)` y fallback defensivo con reintentos.
    - Registro y seguimiento de propiedad de la instancia de Excel vía `self._owns_app`; cierre limpio mediante `app.Quit()` únicamente si la instancia fue creada por el motor y `Workbooks.Count == 0`, evitando el cierre intempestivo de libros abiertos del usuario.
    - Enriquecimiento del comando `status()` retornando `full_name`, `read_only` y `saved` del libro activo.
    - Control de guardado de solo lectura levantando `ExcelLockError` al invocar `save(None)` sobre libros con bandera `ReadOnly=True`.
  - **Introspección y Manipulación de Series de Gráficos (`operations/charts.py`, `backends/` y `engine.py`):**
    - Implementación de `inspect_workbook` tanto en Live COM como Headless: inspecciona metadatos, hojas, visibilidad, dimensiones, tablas estructuradas y conteo de gráficos.
    - Implementación de `inspect_charts`: lista y describe todos los gráficos existentes por hoja con títulos, tipos y series detalladas (valores, categorías y nombres).
    - Implementación de `extend_chart_series`: extiende rangos de series de datos vertical u horizontalmente sin alterar propiedades visuales del gráfico.
    - Implementación de `update_chart_series`: actualiza rangos de valores o categorías para una serie específica por índice o nombre.
    - Parser robusto de fórmulas `SERIES(...)` en `parse_series_formula`, tolerante a hojas con comas, comillas simples, paréntesis y argumentos nulos.
    - Despliegue en `BaseExcelBackend`, `LiveExcelCOMBackend`, `HeadlessOpenPyXLBackend` y delegación en `AntigravityExcelEngine`.
  - **CLI, Operaciones por Lotes (Batch) e Integración IPC (`cli.py`, `ipc/`):**
    - Nuevos subcomandos en CLI: `inspect-workbook`, `inspect-charts`, `extend-chart-series`, `update-chart-series` y `batch`.
    - Comando `batch`: ejecución secuencial de múltiples operaciones dentro del motor con un único guardado al final (`save()`), reduciendo la latencia de I/O en flujos complejos.
    - Registro de comandos en `COMMANDS` y `VALID_COMMANDS` (`ipc/protocol.py`) y despacho en `AntigravityExcelDaemon._execute_command` (`ipc/daemon.py`).
    - Detección de rutas UNC en CLI con timeout ampliado a 60 segundos (`60000ms`) en `_try_daemon` para prevenir cancelaciones prematuras en unidades de red lentas.
  - **Pruebas Automatizadas (`tests/test_introspection_and_charts.py` y `tests/test_backends.py`):**
    - 13 pruebas unitarias nuevas cubriendo parsing de fórmulas `SERIES`, extensión de rangos en gráficos openpyxl, introspección de libros, llamadas batch, subcomandos CLI y timeout UNC.
    - Total de 108/108 pruebas aprobadas al 100% en pytest.

- **Suite de Pruebas E2E de Flujos Completos en Pipelines de Datos (`tests/test_e2e_pipeline.py`):**
  - Implementación de 4 pruebas de integración y pipelines de punta a punta con `AntigravityExcelEngine` en modo headless (`mode="file"`):
    1. `test_e2e_full_lifecycle_data_table`: Ciclo de vida completo de 9 pasos sobre dataset operativo `Ventas_Agua` (inyección cruda, auditoría inicial, curación compuesta con `curate_pipeline`, re-auditoría, inyección de 3 columnas calculadas encadenadas con fórmulas dinámicas, verificación de errores de fórmulas, agregación analítica, creación de tabla dinámica y gráfico sin solapamiento, aplicación de estilo corporativo, escritura declarativa de celdas, copia de rangos y exportación CSV).
    2. `test_e2e_chained_formula_injection_and_resize`: Inyección encadenada de fórmulas interdependientes con expansión dinámica de tabla de 3 a 6 columnas y redimensionamiento automático de rango de tabla abierta (`openpyxl.worksheet.table.Table`).
    3. `test_e2e_curate_aggregate_and_pivot_flow`: Flujo integrado de curación de datos sucios (eliminación de duplicados, normalización de texto y fechas), agregación analítica multivariable y generación cruzada de tabla dinámica y gráfico en hoja separada.
    4. `test_e2e_declarative_scenario_and_styling`: Construcción de escenarios de simulación con entradas de supuestos, celdas de impacto dinámico, aplicación de presets visuales corporativos (`header`, `input_cell`, `total_row`) y exportación a CSV con delimitador personalizado.
- **Caso Práctico Reproducible de Flujo Integral (`pruebas/9_Flujo_Integral_Combinado/`):**
  - Script ejecutor demostrativo `ejecutar_prueba_integral.py` para la orquestación secuencial de 9 etapas avanzadas sobre facturación sanitaria de Biobío y Ñuble.
  - Generación del libro Excel procesado `Facturacion_Agua_PROCESADO.xlsx` (14.7 KB, 3 hojas: `Datos_Facturacion`, `Resumen_Zonal` y `Escenario_Tarifario`, con tabla estructurada `TblFacturacion` expandida a `A1:K55`, gráfico nativo y formato corporativo).
  - Documentación de la Prueba 9 en [`pruebas/BITACORA_PRUEBAS.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/pruebas/BITACORA_PRUEBAS.md).

### Corregido y Optimizado
- **Mapeo de Parámetros en `copy_paste_range` (`operations/write.py`):**
  - Corrección de mapeo de argumentos entre `operations.write.copy_paste_range` y `BaseExcelBackend.copy_paste_range`, transformando adecuadamente `source_range` a `src_range` y `dest_cell` a `dst_start_cell`, habilitando el funcionamiento idéntico en `AntigravityExcelEngine.copy_paste_range` sin discrepancias de firma.
- **Rendimiento e Integración Live COM en Tiempo Real (`excel_engine`):**
  - **Optimización Matricial de `audit_data_quality`:** Sustitución de bucles celda a celda en COM (~11,200 llamadas RPC) por lectura bidimensional en un solo viaje COM (`Range.Value` y `Range.Formula`) procesada vectorialmente en memoria local, reduciendo el tiempo de auditoría de ~4.5s a <15ms.
  - **Corrección de Dirección COM:** Corrección de error `TypeError: 'str' object is not callable` al invocar `target_rng.Address(False, False)` en win32com, unificando el acceso con `get_com_clean_address()`.
  - **Serialización Segura IPC y CLI:** Incorporación de serialización automática con `default=str` en `cli.py`, `daemon.py` y `client.py`, resolviendo incompatibilidades con valores numéricos `Decimal` y fechas devueltos por Excel.
  - **Ajuste de Timeout en `_try_daemon`:** Ampliación del timeout del dispatcher CLI de 50ms a 10,000ms para evitar cancelaciones prematuras y caídas forzadas a standalone cuando el daemon está procesando rangos extensos.
  - **Normalización de Firmas en Operaciones de Visualización y Estilos:** Homogeneización de argumentos (`source`/`source_range`, `preset`/`preset_name`) en `operations/charts.py`, `engine.py` y backends para `create_pivot_table`, `create_chart` y `apply_style_preset`.

### Añadido
- **Pruebas y Modelado Financiero Interactivo en `Financial_Sample.xlsx`:**
  - Auditoría integral de calidad de datos sobre las 11,216 celdas de la tabla estructurada `financials`.
  - Creación de columnas calculadas con fórmulas vivas: `Margen %` (`=L2/J2`) y `Utilidad por unidad` (`=L2/E2`).
  - Creación de tabla dinámica analítica `PT_Utilidad_Segmento_Pais` en hoja `Resumen_Utilidad` y gráfico de columnas agrupadas vinculado sin solapamiento.
  - Modelado de escenario interactivo en hoja `Simulacion_Escenario` con celda editable (`C5: 50.0%`) y matriz de impacto dinámico en ventas netas y utilidad total.

### Añadido
- **Fase 6 del Refactor de EXCEL_ENGINE (Calidad, Tests Modulares y CI):**
  - **Suites de Pruebas Unitarias Adicionales (`tests/`):**
    - `tests/test_engine.py`: 12 pruebas unitarias para `AntigravityExcelEngine` y `ExcelEngine` (inicialización en modo auto y file, fallback inteligente de COM a headless, delegación completa a `operations/read`, `write`, `audit`, `aggregate`, `status`, `wait_until_ready`, `save`, `close` y control de excepciones).
    - `tests/test_cli.py`: 14 pruebas unitarias para el CLI desacoplado (`build_parser`, flags de daemon, `_build_payload`, lectura de JSON/archivos, formateo plano/JSON, ejecución standalone de comandos como `status`, `get-csv`, `get-ranges`, `set-range` y fallback transparente).
    - `tests/test_curation.py`: 6 pruebas unitarias modulares para el pipeline de preparación de datos (`remove_duplicates` con tablas estructuradas, `clean_table_dataset` con trim/casing/tipos/fechas, `detect_and_flag_outliers` con IQR y Z-Score no destructivo, y `curate_pipeline` compuesto en una sola pasada en memoria).
  - **Configuración Unificada de Calidad en `pyproject.toml`:**
    - Declaración de dependencias de desarrollo en `[project.optional-dependencies]`: `pytest`, `pytest-cov`, `ruff` y `mypy`.
    - Configuración centralizada de herramientas: `tool.pytest.ini_options`, `tool.coverage.run`, `tool.coverage.report`, `tool.ruff` y `tool.mypy`.
  - **Integración Continua para Windows (`.github/workflows/ci.yml`):**
    - Workflow automatizado en `windows-latest` con matriz para Python 3.10, 3.11 y 3.12.
    - Ejecución de linting con `ruff`, chequeo de tipos con `mypy` y suite de pruebas completa con reporte de cobertura (`pytest-cov`).
  - **Resultados de Cobertura y Verificación:**
    - Total de pruebas en la suite: **87/87 pruebas aprobadas al 100%** (4.31s).
    - Cobertura global alcanzada en suite mixta: **49%** (100% en `engine.py`, `protocol.py`, `models`, `operations/read`, `formatting`, `formulas` y módulos core headless).
- **Fase 5 del Refactor de EXCEL_ENGINE (IPC/Daemon y CLI modular):**
  - Subpaquete `antigravity_excel_engine.ipc`:
    - `protocol.py`: Constantes canónicas del Named Pipe (`PIPE_NAME`, `BUF_SIZE`, `DEFAULT_TIMEOUT_MS`), constructores de mensajes (`build_command_request`, `build_action_request`, `build_success_response`, `build_error_response`) y validador de solicitudes entrantes (`validate_request`).
    - `client.py`: Funciones de alto nivel (`send_pipe_message`, `ping_daemon`, `call_command`, `start_daemon_process`, `stop_daemon_process`, `status_daemon`) con logging explícito y sin `except` silenciosos.
    - `daemon.py`: `AntigravityExcelDaemon` desacoplado — usa `engine.py` + `protocol.py` + `models/errors.py`; respuestas siempre estructuradas `{status, result|error, [latency_ms]}`; manejo explícito de `ExcelLockError`, `WorkbookNotFoundError`, `FormulaValidationError`, `UnsupportedOperationError`; `logger.exception` en errores inesperados; sin `except Exception: pass`.
  - `antigravity_excel_engine/cli.py`: CLI desacoplado con `build_parser()`, `_build_payload()` y `_run_standalone()`. Usa `ipc/client.call_command` en lugar de importar el daemon directamente. Sin lógica de negocio Excel.
  - `antigravity_excel_daemon.py`: Convertido en shim de ~75 líneas que re-exporta desde `ipc/`.
  - `antigravity_excel_cli.py`: Convertido en shim de ~20 líneas que re-exporta desde `antigravity_excel_engine/cli.py`.
  - Suite de pruebas: **55/55 pruebas aprobadas** sin regresiones (1.16s).
- **Fase 4 del Refactor de EXCEL_ENGINE (Operations y Engine Orquestador):**
  - Subpaquete `antigravity_excel_engine.operations` con separación completa de responsabilidades por dominio:
    - `read.py`: `read_ranges`, `get_range_as_csv` — lectura sin lógica de negocio, delegación pura al backend.
    - `write.py`: `write_cell_range`, `write_cells_declarative`, `copy_paste_range` — escritura declarativa.
    - `audit.py`: `audit_data_quality`, `check_formula_errors`, `detect_and_flag_outliers` — auditoría de integridad.
    - `aggregate.py`: `aggregate`, `add_calculated_column` — agregaciones estadísticas y columnas calculadas.
    - `curation.py`: `remove_duplicates`, `clean_table_dataset`, `curate_pipeline` — pipeline de limpieza y deduplicación.
    - `charts.py`: `create_pivot_table`, `create_chart`, `create_summary_table` — visualización y tablas dinámicas.
  - Módulo `antigravity_excel_engine/engine.py`: `AntigravityExcelEngine` como coordinador delgado (< 280 líneas) con inyección de backend, auto-detección COM/headless y delegación total a `operations/*`. Expone alias `ExcelEngine`.
  - `antigravity_excel_engine/__init__.py` actualizado: exporta `AntigravityExcelEngine`, `ExcelEngine` y todas las operaciones públicas.
  - `antigravity_excel_core.py` actualizado: re-exporta `AntigravityExcelEngine` y `ExcelEngine` desde `engine.py`, garantizando compatibilidad retroactiva 100%.
  - Suite de pruebas: **55/55 pruebas aprobadas** sin regresiones (1.74s).
- **Fase 3 del Refactor de EXCEL_ENGINE (Backends: Base, Live COM y Headless):**
  - Subpaquete `antigravity_excel_engine.backends`:
    - `base.py`: Definición de la clase base abstracta `BaseExcelBackend` (ABC) consolidando el contrato unificado para backends de Excel con métodos abstractos y comunes (`open_workbook`, `close`, `save`, `get_sheet`, `read_range`, `write_range`, `get_cell_value`, `set_cell_value`, `supports_formula_recalc`, `is_busy`, `get_cell_ranges`, `set_cell_range`, `set_cells`, `get_range_as_csv`, `copy_paste_range`, `check_formula_errors`, `audit_data_quality`, `create_pivot_table`, `create_chart`, `aggregate`, `add_calculated_column`, `apply_style_preset`, `create_summary_table`, `remove_duplicates`, `detect_and_flag_outliers`, `clean_table_dataset`, `curate_pipeline`, `status`, `wait_until_ready`, `dispatch_com`, `_calculate_placement_coordinates`).
    - `live_com.py`: Extracción modular de `LiveExcelCOMBackend`, el decorador `@retry_on_excel_busy()`, el context manager `excel_fast_mode()` y la utilidad `get_com_clean_address()`. Preserva la arquitectura completa de integración con Microsoft Excel COM (acceso a ROT, enganche visual en `WinSta0\\Default` con `oleacc`, serialización y marshalling entre hilos de COM streams, optimizaciones matriciales de alto rendimiento, tablas dinámicas, gráficos, despacho COM dinámico `dispatch_com` y tolerancia adaptativa a bloqueos interactivos).
    - `headless.py`: Extracción modular de `HeadlessOpenPyXLBackend`. Preserva la carga de libros en disco con lectura concurrente ante archivos bloqueados vía `win32file.CreateFile`, generación y actualización de rangos, formateo de celdas y comentarios, tablas analíticas cruzadas, gráficos openpyxl y pipeline declarativo completo de curaduría de datos.
    - `__init__.py`: Exportación centralizada de `BaseExcelBackend`, `LiveExcelCOMBackend`, `HeadlessOpenPyXLBackend`, `retry_on_excel_busy`, `excel_fast_mode` y `get_com_clean_address`.
  - Refactorización de `antigravity_excel_core.py`: Reducción arquitectónica masiva (de ~4,400 líneas a ~340 líneas), delegando directamente en `antigravity_excel_engine.backends` y re-exportando el 100% de símbolos para compatibilidad retroactiva absoluta.
  - Exposición en la raíz del paquete `antigravity_excel_engine/__init__.py`.
  - Suite de pruebas unitarias dedicada: `tests/test_backends.py` con 12 nuevos casos de prueba, alcanzando 55/55 pruebas pasando al 100%.
- **Fase 2 del Refactor de EXCEL_ENGINE (Fórmulas y Formateo Modular):**
  - Subpaquete `antigravity_excel_engine.formatting`:
    - `presets.py`: Extracción y enriquecimiento de `FinancialStyleGuide` con formatos numéricos (`CURRENCY_FORMAT`, `CURRENCY_CLP`, `CURRENCY_USD`, `PERCENT`, `DATE`), colores corporativos Essbio/institucionales (`COLOR_PRIMARY`, `COLOR_SECONDARY`, `COLOR_INPUT`, `COLOR_ASSUMPTION_FILL`, `COLOR_HEADER_FILL`, `COLOR_ZEBRA_EVEN`, `COLOR_ZEBRA_ODD`, `COLOR_KPI_CARD_FILL`, `COLOR_DELTA_POS`, `COLOR_DELTA_NEG`, etc.), diccionario de definiciones declarativas `PRESETS` y método resiliente regional `apply_number_format_resilient`.
    - `styles.py`: Funciones universales de aplicación de formato y presets tanto para celdas de `openpyxl` como objetos `Range` de `win32com` (`apply_style_preset`, `apply_style_preset_openpyxl`, `apply_style_preset_com`, `apply_number_format`, `apply_borders_openpyxl`, `hex_to_rgb_int`, `hex_to_argb_hex`).
    - `__init__.py`: Exportación unificada de utilidades de diseño y estilos.
  - Subpaquete `antigravity_excel_engine.formulas`:
    - `sanitize.py`: Normalización de fórmulas y compatibilidad XML (`sanitize_formula_compatibility`, `strip_xlfn_prefix`) con catálogo ampliado de 36 funciones modernas (post-2007, matrices dinámicas de Excel 365/2021 como `XLOOKUP`, `UNIQUE`, `FILTER`, `SORT`, `SEQUENCE`, `LET`, `LAMBDA`, `VSTACK`, `HSTACK`, `TEXTSPLIT`).
    - `audit.py`: Auditoría y reconocimiento de errores de cálculo (`#VALUE!`, `#REF!`, `#DIV/0!`, `#N/A`, `#NAME?`, `#NUM!`, `#NULL!`, `#SPILL!`, `#CALC!`), mapa de códigos COM (`COM_EXCEL_ERROR_MAP`), verificadores (`is_formula_error_value`, `get_formula_error_name`) y extractores sobre hojas (`audit_openpyxl_formula_errors`, `audit_com_formula_errors`).
    - `compatibility.py`: Reemplazo de placeholders dinámicos (`replace_formula_placeholders` con `{r}`, `{row}`, `{c}`, `{col}`, `{table}` y variables libres), propagación con preservación estricta de referencias relativas/absolutas (`propagate_formula_row`) y constructores de referencias (`make_absolute_reference`, `make_relative_reference`, etc.).
    - `__init__.py`: Exportación unificada de utilidades de fórmulas.
  - Integración en `antigravity_excel_core.py`: Delegación limpia de `FinancialStyleGuide`, `sanitize_formula_compatibility`, `propagate_formula_row`, `check_formula_errors` y `apply_style_preset` hacia los nuevos módulos modulares con 100% de compatibilidad hacia atrás.
  - Suite de pruebas unitarias dedicada: `tests/test_formatting_formulas.py` con 10 nuevos casos de prueba, alcanzando 43/43 tests pasando al 100%.
- **Fase 1 del Refactor de EXCEL_ENGINE (Modelos, Errores y Utilidades Base):**
  - Creación del paquete raíz modular `antigravity_excel_engine/` con `__init__.py` y registro de autoría `Creado por Francisco Grandón Vergara`.
  - Módulo `antigravity_excel_engine.models`:
    - `errors.py`: Jerarquía de excepciones (`ExcelEngineError`, `ExcelLockError`, `WorkbookNotFoundError`, `FormulaValidationError`, `PipeConnectionError`, `UnsupportedOperationError`, `InvalidRangeError`, `ExcelBusyError`) con serialización a diccionario.
    - `requests.py`: Dataclasses tipadas para solicitudes (`DaemonRequest`, `PipeCommandRequest`, `CellWriteRequest`, `CuratePipelineRequest`).
    - `responses.py`: Estructuras de respuesta normalizadas (`BaseResponse`, `SuccessResponse`, `ErrorResponse`, `OperationStatus`).
    - `result_types.py`: Dataclasses de resultados (`CellWriteResult`, `AuditReport`, `CurateResult`, `AggregateResult`, `PivotResult`, `OutlierResult`).
  - Módulo `antigravity_excel_engine.utils`:
    - `logging.py`: Logger centralizado `antigravity_excel`, helpers `log_operation`, `log_exception` y context manager `log_duration`.
    - `ranges.py`: Utilidades de coordenadas y rangos (`coordinate_from_string`, `column_index_from_string`, `col_to_index`, `get_column_letter`, `index_to_col`, `range_boundaries`, `parse_range`, `chunk_cell_addresses`) con LRU cache.
    - `dates.py`: Conversión y serialización de fechas a seriales de Excel (`parse_excel_date_serial`, `coerce_date_to_excel_serial`, `excel_serial_to_datetime`, `excel_serial_to_date`).
    - `validation.py`: Validadores y sanitizadores de nombres de hojas (`validate_sheet_name`, `sanitize_sheet_name`), rangos (`validate_range_string`), existencia de libros (`validate_workbook_exists`), sintaxis de fórmulas (`validate_formula_text`) y detección de errores de cálculo (`is_formula_error`).
- **Biblioteca Modular de Habilidades para Agentes de IA (`excel-skills/`):**
  - Catálogo de habilidades de interacción inspirado en la arquitectura modular de `interaction-skills` de browser-use, adaptado al ecosistema de Microsoft Excel COM, Headless OpenPyXL y Daemon IPC de alta velocidad.
  - `excel-skills/README.md`: Visión general, catálogo e índice de habilidades, mapa de arquitectura y reglas de oro de latencia (<5ms vía Daemon Named Pipes), eficiencia de tokens y cero corrupción de libros.
  - `excel-skills/com-locks-and-readiness.md`: Sincronización adaptativa con `wait_until_ready`, decorador `@retry_on_excel_busy`, catálogo de HRESULTs de bloqueo interactivo (`0x8001010A`, `0x800AC472`, `0x80010001`), manejo seguro de `excel_fast_mode` y restauración obligatoria de `ScreenUpdating` y `Calculation`.
  - `excel-skills/formulas-dynamic-spill.md`: Inyección de fórmulas vectorizadas (`UNIQUE`, `FILTER`, `SORT`, `XLOOKUP`), notación estructurada `Table[@Columna]`, propiedad `.Formula2` vs `.Formula`, prevención del error `#SPILL!` y resolución de discrepancias regionales (`;` vs `,`).
  - `excel-skills/dates-and-serial-numbers.md`: Conversión bidireccional ultrarrápida entre fechas ISO-8601 y números seriales de Excel, época base computacional (30 de diciembre de 1899), resolución del bug bisiesto de Lotus 1-2-3 (Serial 60) y prevención de desfases de huso horario.
  - `excel-skills/dispatch-com-advanced.md`: Invocación reflexiva de la API profunda mediante `dispatch_com` (`antigravity_excel_cli.py dispatch-com`) para autofiltros dinámicos, protección de hojas con contraseña, listas desplegables (validación de datos) y exportación nativa a PDF, previniendo objetos COM huérfanos y procesos `EXCEL.EXE` zombis.
  - `excel-skills/benchmarking-harness.md`: Harness de evaluación estandarizado para agentes de IA sobre Excel utilizando el banco de pruebas de `PRUEBAS/`, con métricas de precisión de cálculo contra claves de verdad, ahorro de tokens de contexto (>60% en CSV vs JSON), latencia IPC (<5ms) y garantía de cero procesos huérfanos.
- **Primitivas de Sincronización y Despacho Dinámico COM (`wait_until_ready` y `dispatch_com`):**
  - Incorporadas en `BaseExcelBackend`, `LiveExcelCOMBackend`, `HeadlessOpenPyXLBackend` y la fachada `AntigravityExcelEngine`.
  - `wait_until_ready`: Sondeo adaptativo con timeout configurable que verifica `app.Ready` y `app.CalculationState == 0` (xlDone). Tolera bloqueos por edición interactiva del usuario mediante captura de HRESULTs COM (`0x8001010A`, `0x800AC472`, `0x80010001`). En backend Headless retorna disponibilidad inmediata sin esperas.
  - `dispatch_com`: Invocación reflexiva tolerante a fallos con `@retry_on_excel_busy()`. Resuelve destinos jerárquicos (`app`, `workbook`/`wb`, `activesheet`/`sheet`, rangos o expresiones vía `app.Evaluate` y propiedades encadenadas). Invoca métodos con argumentos o efectúa lectura/asignación de propiedades si no son invocables, serializando los resultados a tipos estándar Python y liberando referencias COM explícitamente (`del res`) para prevenir fugas de recursos.
  - Integración en `antigravity_excel_daemon.py` dentro de `handle_request` para canalizar peticiones `wait-until-ready` y `dispatch-com` a través del Named Pipe.
  - Nuevos subparsers en `antigravity_excel_cli.py` para `wait-until-ready` (`--timeout`, `--interval`) y `dispatch-com` (`--target`, `--member`, `--args-json`, `--kwargs-json`, `--stdin`, `--input-file`) con despacho por daemon y fallback transparente a standalone.
  - Cobertura de pruebas unitarias y de integración en `tests/test_excel_engine.py` y `tests/test_daemon_pipe.py` (24 unit tests y 9 integration tests aprobados al 100%).
- **Servidor Daemon Residente y Named Pipe IPC (`antigravity_excel_daemon.py` - Tarea T3.1):** Servidor persistente multihilo en Windows que mantiene en memoria (RAM) la sesión de Excel y `AntigravityExcelEngine` caliente, proveyendo un canal IPC duplex de ultra-baja latencia sobre el Named Pipe `\\.\pipe\antigravity_excel` (búferes de 16MB para transferencias masivas). Soporta los 18 comandos del motor (`status`, `get-csv`, `get-ranges`, `set-range`, `set-cells`, `copy-paste`, `check-errors`, `audit-quality`, `aggregate`, `add-column`, `create-pivot`, `create-chart`, `summary-table`, `apply-preset`, `dedup`, `flag-outliers`, `clean`, `curate`) más acciones de ping y shutdown.
- **Cliente Ultrarrápido (< 30ms) y Fallback Transparente en CLI (`antigravity_excel_cli.py`):** Integración de `_try_call_daemon` con timeout ultra-corto (50ms). Si el daemon está activo, despacha la petición por el pipe en ~5ms (latencia promedio ~5.25 ms); si está inactivo, cae de forma transparente e instantánea (< 1ms) al flujo standalone local sin errores.
- **Flags de Gestión del Daemon en CLI y Proceso Autónomo:** Nuevos modificadores `--daemon-start` (despacho en background desacoplado), `--daemon-stop` (señal de apagado ordenado) y `--daemon-status` (diagnóstico en tiempo real y latencia de ping) en `antigravity_excel_cli.py` y `antigravity_excel_daemon.py`.
- **Suite de Pruebas Unitarias y de Integración IPC (`tests/test_daemon_pipe.py`):** Cobertura completa de pruebas para latencia de ping, comandos sobre pipe (`status`, `set-range`, `get-csv`, `get-ranges`, `audit-quality`), control de errores y fallback transparente offline, elevando la suite total del proyecto a 27 pruebas aprobadas al 100%.
- **Pipeline Unificado de Curaduría y Vectorización en Memoria (`curate_pipeline` - Tarea T2.1):** Método de alto nivel en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que resuelve tareas compuestas de preparación de datos (deduplicación por fila/columnas clave, estandarización de texto, casteo estricto de tipos, coerción de fechas a números seriales y detección/marcado no destructivo de outliers) en UNA SOLA PASADA en memoria Python y UN SOLO VIAJE COM de lectura/escritura bajo `excel_fast_mode`, reduciendo drásticamente la latencia en escrituras y volcado matricial.
- **Vectorización Multi-Rango y Formateo Agrupado de Outliers (`_chunk_cell_addresses`):** Helper y optimización en `LiveExcelCOMBackend.detect_and_flag_outliers` y `curate_pipeline` que agrupa coordenadas no contiguas en cadenas de rango combinadas (`ws.Range("E325,E412,E689")`) respetando límites seguros de caracteres COM, aplicando formato visual (#FCE4D6 y texto rojo en negrita) en una sola llamada COM por bloque en vez de llamadas celda a celda.
- **Subcomando CLI `curate`:** Nuevo comando en `antigravity_excel_cli.py` con opciones `--source`, `--spec-json`, `--input-file`, `--stdin` y `--sheet` con salida JSON estructurada y soporte UTF-8 forzado en stdout de Windows.
- **Cobertura de Pruebas Unitarias para Pipeline Unificado:** Nuevo test `test_curate_pipeline_headless` en `tests/test_excel_engine.py` validando la ejecución compuesta de dedup + clean + outliers y redimensionamiento de tabla estructurada, alcanzando 19/19 pruebas aprobadas al 100%.
- **Conversión Universal de Fechas a Número de Serie de Excel (`parse_excel_date_serial`):** Función auxiliar en `antigravity_excel_core.py` que interpreta formatos ISO ('YYYY-MM-DD', 'YYYY/MM/DD'), fechas con slashes ('DD/MM/YYYY', 'MM/DD/YYYY'), marcas de hora y objetos datetime nativos, convirtiéndolos con precisión a números seriales de Excel (días transcurridos desde 1899-12-30).

- **Deduplicación Inteligente On-Demand (`remove_duplicates`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` para eliminar filas duplicadas por fila completa o por columnas clave especificadas (`key_columns`). En Live COM redimensiona automáticamente la tabla estructurada (`ListObject.Resize`) dentro de `excel_fast_mode` sin dejar filas vacías, y en Headless OpenPyXL actualiza `tbl.ref`.
- **Detección y Señalización Visual No Destructiva de Outliers (`detect_and_flag_outliers`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que evalúa valores atípicos mediante rango intercuartílico (IQR) y Z-score adaptativo ($z = 2.5$ para $N < 30$, $z = 3.0$ para $N \ge 30$). Aplica formato visual de advertencia no destructivo (fondo suave `#FCE4D6`, texto rojo oscuro `#C00000` negrita) e inserta comentarios de celda explicativos preservando los valores numéricos intactos.
- **Limpieza y Estandarización Integral de Datasets (`clean_table_dataset`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que procesa en una sola pasada en memoria: normalización de espacios y casing con diccionario de sinónimos/reemplazos (`standardize_text`), coerción de números en texto a tipos primitivos (`cast_types`), normalización de fechas a seriales con formato uniforme (`coerce_dates`) y limpieza de encabezados (`trim_headers`), volcando los resultados en bloque matricial bajo `excel_fast_mode`.
- **Subcomandos CLI `dedup`, `flag-outliers` y `clean`:** Nuevas operaciones de línea de comandos en `antigravity_excel_cli.py` con argumentos para origen (`--source`), columnas clave (`--key-cols`), columna métrica (`--col`), método (`--method`), supresión de comentarios (`--no-comment`) y especificación declarativa de reglas (`--rules-json`, `--input-file`, `--stdin`).
- **Ampliación de Cobertura de Pruebas Unitarias de Calidad y Curación:** Nuevos tests en `tests/test_excel_engine.py` (`test_parse_excel_date_serial`, `test_remove_duplicates_headless`, `test_detect_and_flag_outliers_headless`, `test_clean_table_dataset_headless`) alcanzando 18 pruebas aprobadas al 100%.
- **Sistema Declarativo Universal de Presets de Estilo (`apply_style_preset`):** Método implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. Permite aplicar estilos de modelado financiero estandarizados de manera declarativa sobre cualquier rango de celdas (`header`, `input_cell`, `total_row`, `currency`, `percent`, `delta_positive`, `delta_negative`).
- **Motor Genérico de Tablas de Resumen / Comparativas (`create_summary_table`):** Método implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` para estructurar automáticamente cuadros de mando ejecutivos con bloques de supuestos y parámetros vinculados, título de destaque, encabezados corporativos, filas con formato condicional de subtotales, totales contables dobles y cuadrícula activa.
- **Protección y Propagación de Referencias Absolutas vs. Relativas (`propagate_formula_row`):** Función en `antigravity_excel_core.py` que preserva intactas las referencias absolutas en fila o celda (`$A$1`, `'Params'!$C$4`, `$C$2`) mientras sustituye dinámicamente las referencias relativas (`A2`, `$B2`, `{r}`, `{row}`) en la propagación de fórmulas por fila en `add_calculated_column`.
- **Subcomandos CLI `apply-preset` y `summary-table`:** Nuevas operaciones expuestas en `antigravity_excel_cli.py` con opciones de configuración vía argumentos de línea de comandos, archivos JSON externos o canal estándar stdin.
- **Cobertura de Pruebas Automatizadas para Presets, Tablas Resumen y Referencias Absolutas:** Incorporación de `test_apply_style_preset_headless`, `test_create_summary_table_headless` y `test_add_calculated_column_absolute_ref` en `tests/test_excel_engine.py` (14 pruebas pasando al 100%).
- **Guía de Estilo y Modelado Financiero Profesional (`FinancialStyleGuide`):** Clase de utilidades y directrices de modelado financiero estándar en `antigravity_excel_core.py`. Centraliza máscaras contables de moneda (`$ #,##0.00;($ #,##0.00);"-"`), porcentaje (`0.0%`) y paleta de colores para celdas de entrada (`#1F4E79`), fórmulas (`#000000`) y supuestos/escenarios (`#FFF2CC`).
- **Sanitización de Compatibilidad de Fórmulas Post-2007 (`sanitize_formula_compatibility`):** Función auxiliar que identifica funciones modernas (`TEXTJOIN`, `CONCAT`, `IFS`, `SWITCH`, `MAXIFS`, `MINIFS`) e inyecta el prefijo de compatibilidad XML `_xlfn.` en backends como openpyxl para prevenir fórmulas corruptas al abrir libros en versiones recientes de Excel.
- **Integración de Sanitización Vectorizada y Declarativa:** Incorporación de `sanitize_formula_compatibility` en `HeadlessOpenPyXLBackend.add_calculated_column`, `set_cell_range` y `set_cells`.
- **Cobertura de Pruebas para Modelado Financiero y Compatibilidad:** Nuevos tests unitarios en `tests/test_excel_engine.py` validando la inyección de `_xlfn.`, preservación de operadores y asignación de formatos contables por defecto en `create_pivot_table` y `add_calculated_column` (11 pruebas pasando al 100%).
- **Conector Nativo y Marshalling Inter-Thread vía `oleacc` y COM Streams:** Integración en `LiveExcelCOMBackend._try_get_live_desktop_excel` con puente entre escritorios de Windows (`WinSta0\Default`) y serialización de punteros `IDispatch` mediante `CoMarshalInterThreadInterfaceInStream` / `CoGetInterfaceAndReleaseStream`. Permite engancharse en tiempo real a la ventana interactiva del usuario sin errores de threading.
- **Resolución Inteligente de Rutas Sincronizadas OneDrive / SharePoint:** Mecanismo en `_init_workbook` que reconoce indistintamente rutas de disco locales (`D:\...`) y URLs remotas de SharePoint (`https://...sharepoint.com/...`), resolviendo libros por coincidencia exacta, decodificación URL o nombre base de archivo.
- **Autoselección Transparente de Modo Live en `mode="auto"`:** `AntigravityExcelEngine` ahora detecta de forma autónoma si el archivo solicitado se encuentra abierto en cualquier ventana de Excel en el escritorio y selecciona de inmediato `LiveExcelCOMBackend` sin requerir flags manuales.
- **Inyección de Columnas Calculadas y Auto-Resize de Tablas (`add_calculated_column`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que inserta una nueva columna con fórmulas relativas propagadas para todas las filas de datos, aplica formato numérico y redimensiona automáticamente la tabla estructurada (`ListObject` / `Table`) si la hoja contiene tablas oficiales.
- **Comando CLI `add-column`:** Subcomando en `antigravity_excel_cli.py` con opciones `--header`, `--formula`, `--number-format`, `--sheet` y `--no-autofit`.
- **Módulo de Agregación Analítica y Pivot Nativo (`aggregate`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que permite agrupar y tabular métricas numéricas (`sum`, `avg`/`mean`, `count`, `min`, `max`) con tolerancia a mayúsculas y espacios invisibles en nombres de columnas.
- **Rastreo Automático de Celdas de Citación:** El método `aggregate` reporta de forma nativa la celda del valor máximo (`max_cell`), mínimo (`min_cell`) y una muestra representativa de celdas (`sample_citation_cells`) para sustentar respuestas analíticas de LLMs sin alucinación.
- **Comando CLI `aggregate`:** Nuevo subcomando en `antigravity_excel_cli.py` con opciones `--group-by`, `--metric`, `--func`, `--top` y `--ascending`.
- **Lectura Resiliente de Archivos Bloqueados en Windows:** Mecanismo en `HeadlessOpenPyXLBackend._load_workbook_resilient` que detecta bloqueos exclusivos de escritura (`PermissionError` / `Errno 13`) provocados por instancias de Excel abiertas y lee los bytes concurrentemente mediante `win32file.CreateFile` con flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`.
- **Creación de Tablas Dinámicas Nativas y Tabulaciones Cruzadas (`create_pivot_table`):** Método implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend`. En Live COM genera `PivotCache` y `PivotTable` con campos de filas, columnas y agregación (`sum`, `count`, `average`, `min`, `max`), formato numérico y visualización de cuadrícula (`DisplayGridlines`). En Headless genera una hoja analítica con tabla de doble entrada cruzada (*cross-tabulation*), totales generales y formato corporativo.
- **Inserción de Gráficos Estadísticos No Invasivos (`create_chart`):** Método implementado en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que inserta gráficos (`column_clustered`, `bar_clustered`, `line`, `pie`, `area`) enlazados a la tabla dinámica o rango fuente, con título y dimensiones configurables.
- **Motor de Posicionamiento Geométrico Anti-Solapamiento (`_calculate_placement_coordinates`):** Helper en `BaseExcelBackend` que evalúa las dimensiones de la tabla de datos y ubica automáticamente el gráfico a la derecha ($\le 8$ columnas) o abajo ($> 8$ columnas) con margen de resguardo de 2 celdas, o en celdas de anclaje personalizadas.
- **Subcomandos CLI `create-pivot` y `create-chart`:** Expuestos en `antigravity_excel_cli.py` con soporte completo de flags (`--source`, `--rows`, `--cols`, `--value-field`, `--value-func`, `--value-format`, `--dest-sheet`, `--dest-cell`, `--table-name`, `--chart-type`, `--title`, `--placement`, `--anchor`).
- **Ampliación de Suite de Pruebas Unitarias:** Cobertura de tests unitarios para `_calculate_placement_coordinates`, `create_pivot_table` y `create_chart` en modo headless en `tests/test_excel_engine.py` (9 pruebas pasando al 100%).
- **Suite de Pruebas Automatizadas:** Cobertura de pruebas unitarias para `audit_data_quality`, `aggregate` y `add_calculated_column` en `tests/test_excel_engine.py` (6 pruebas pasando al 100%).

### Cambiado
- **Lazy Loading y Reducción Radical de Tiempos de Importación (Tarea T1.1):**
  - En `antigravity_excel_core.py`: Se eliminó la importación ansiosa a nivel de módulo de bibliotecas pesadas (`openpyxl`, `openpyxl.styles`, `openpyxl.comments`) y módulos secundarios de Windows/COM (`win32gui`, `win32service`, `win32file`, `win32con`, `ctypes`, `uuid`). Se implementaron utilidades de celda y coordenadas puras con expresiones regulares optimizadas y cacheo LRU (`coordinate_from_string`, `column_index_from_string`, `get_column_letter`, `range_boundaries`), compatibles al 100% con `openpyxl` pero sin coste de inicialización. Se añadieron lazy loaders (`_get_openpyxl()`, `_get_styles()`, etc.) y resolución transparente vía `__getattr__` a nivel de módulo.
  - En `antigravity_excel_cli.py`: Se postergó la importación de `AntigravityExcelEngine` y `antigravity_excel_core` dentro de los handlers de ejecución de `main()`, evitando la inicialización de motores durante la lectura de argumentos y subcomandos de `argparse`.
  - **Impacto Cuantitativo:** Reducción del tiempo de importación en frío de `antigravity_excel_core` de ~450–726 ms a **57.14 ms** (~87% más rápido) y de `antigravity_excel_cli` de ~380 ms a **22.01 ms** (~94% más rápido), superando holgadamente la meta de < 150 ms manteniendo 18/18 pruebas unitarias pasando al 100%.
- **Protección de Referencias Absolutas y Placeholders en `add_calculated_column`:** Se integró la función `propagate_formula_row` tanto en `HeadlessOpenPyXLBackend` como en `LiveExcelCOMBackend` para garantizar que las referencias con signo de dólar (`$`) permanezcan inalteradas mientras las referencias relativas y los placeholders `{r}` / `{row}` se sustituyen dinámicamente con el índice de fila de datos.
- **Formato Numérico por Defecto en Tablas Dinámicas y Columnas Calculadas:** `create_pivot_table` y `add_calculated_column` ahora adoptan por defecto el estándar contable de moneda `FinancialStyleGuide.CURRENCY_FORMAT` con negativos entre paréntesis y ceros en guion.
- **Valor por Defecto de `--value-format` en CLI `create-pivot`:** Configurado por defecto a `FinancialStyleGuide.CURRENCY_FORMAT` en `antigravity_excel_cli.py`.
- **Flexibilidad de Argumentos en CLI (`antigravity_excel_cli.py`):** Las opciones globales `--file`, `--mode` y `--json` ahora están disponibles tanto antes como después del subcomando (ej. `python antigravity_excel_cli.py audit-quality --file data.xlsx --json`).
