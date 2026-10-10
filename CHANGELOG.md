# Changelog

Todas las modificaciones notables de este proyecto serán documentadas en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.0.0/),
y este proyecto adhiere a [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Añadido
- **Lectura Resiliente de Archivos Bloqueados en Windows:** Mecanismo en `HeadlessOpenPyXLBackend._load_workbook_resilient` que detecta bloqueos exclusivos de escritura (`PermissionError` / `Errno 13`) provocados por instancias de Excel abiertas y lee los bytes concurrentemente mediante `win32file.CreateFile` con flags `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE`.
- **Módulo de Auditoría de Higiene y Calidad de Datos:** Método `audit_data_quality(range_str, sheet)` en la interfaz base, `LiveExcelCOMBackend` y `HeadlessOpenPyXLBackend` para detectar espacios residuales en encabezados, números tipados como texto, inconsistencia de tipos en columnas y conteo de fórmulas dinámicas.
- **Comando CLI `audit-quality`:** Exposición del reporte de calidad con salida formateada o JSON determinista para agentes de IA.
- **Suite de Pruebas Automatizadas:** Cobertura de pruebas unitarias para `audit_data_quality` en `tests/test_excel_engine.py`.

### Cambiado
- **Flexibilidad de Argumentos en CLI (`antigravity_excel_cli.py`):** Las opciones globales `--file`, `--mode` y `--json` ahora están disponibles tanto antes como después del subcomando (ej. `python antigravity_excel_cli.py audit-quality --file data.xlsx --json`).
