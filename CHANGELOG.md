# Changelog

Todas las modificaciones notables de este proyecto serán documentadas en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Añadido
- **Conector Nativo a Ventanas Interactivas de Excel vía Windows Accessibility API (`oleacc`):** Implementación en `LiveExcelCOMBackend._try_get_live_desktop_excel` que localiza la ventana real del usuario (`XLMAIN` y control hijo `EXCEL7`), enviando `WM_GETOBJECT` y resolviendo el puntero `IDispatch` mediante `oleacc.ObjectFromLresult`. Permite controlar la ventana viva de Excel incluso cuando la aplicación no está registrada en la ROT (Running Object Table) de COM.
- **Inyección de Columnas Calculadas y Auto-Resize de Tablas (`add_calculated_column`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que inserta una nueva columna con fórmulas relativas propagadas para todas las filas de datos, aplica formato numérico y redimensiona automáticamente la tabla estructurada (`ListObject` / `Table`) si la hoja contiene tablas oficiales.
- **Comando CLI `add-column`:** Subcomando en `antigravity_excel_cli.py` con opciones `--header`, `--formula`, `--number-format`, `--sheet` y `--no-autofit`.
- **Módulo de Agregación Analítica y Pivot Nativo (`aggregate`):** Método en `BaseExcelBackend`, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` que permite agrupar y tabular métricas numéricas (`sum`, `avg`/`mean`, `count`, `min`, `max`) con tolerancia a mayúsculas y espacios invisibles en nombres de columnas.
- **Rastreo Automático de Celdas de Citación:** El método `aggregate` reporta de forma nativa la celda del valor máximo (`max_cell`), mínimo (`min_cell`) y una muestra representativa de celdas (`sample_citation_cells`) para sustentar respuestas analíticas de LLMs sin alucinación.
- **Comando CLI `aggregate`:** Nuevo subcomando en `antigravity_excel_cli.py` con opciones `--group-by`, `--metric`, `--func`, `--top` y `--ascending`.
- **Lectura Resiliente de Archivos Bloqueados en Windows:** Mecanismo en `HeadlessOpenPyXLBackend._load_workbook_resilient` que detecta bloqueos exclusivos de escritura (`PermissionError` / `Errno 13`) provocados por instancias de Excel abiertas y lee los bytes concurrentemente mediante `win32file.CreateFile` con flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`.
- **Módulo de Auditoría de Higiene y Calidad de Datos:** Método `audit_data_quality(range_str, sheet)` en la interfaz base, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` para detectar espacios residuales en encabezados, números tipados como texto, inconsistencia de tipos en columnas y conteo de fórmulas dinámicas.
- **Comando CLI `audit-quality`:** Exposición del reporte de calidad con salida formateada o JSON determinista para agentes de IA.
- **Suite de Pruebas Automatizadas:** Cobertura de pruebas unitarias para `audit_data_quality`, `aggregate` y `add_calculated_column` en `tests/test_excel_engine.py` (6 pruebas pasando al 100%).

### Cambiado
- **Flexibilidad de Argumentos en CLI (`antigravity_excel_cli.py`):** Las opciones globales `--file`, `--mode` y `--json` ahora están disponibles tanto antes como después del subcomando (ej. `python antigravity_excel_cli.py audit-quality --file data.xlsx --json`).
