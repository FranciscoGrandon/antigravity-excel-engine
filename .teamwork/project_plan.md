# Plan de Proyecto, Arquitectura y Descomposición de Hitos (Project Plan)

> **Documento Oficial de Arquitectura — Teamwork Custom**  
> **Proyecto:** `ExcelEngine_Performance_Optimization`  
> **Modo:** `benchmark`

---

## 1. Arquitectura y Descomposición por Hitos

### Hito 1 (M1): Optimización de Carga Perezosa (Lazy Imports) y Benchmark Baseline
- **Objetivo:** Reducir el tiempo de importación inicial de 726 ms a < 150 ms.
- **Entregables:** Refactorización de imports diferidos en `antigravity_excel_core.py` y `antigravity_excel_cli.py`.
- **Archivos Exclusivos:** `antigravity_excel_core.py`, `antigravity_excel_cli.py`.

### Hito 2 (M2): Pipeline Unificado de Curación y Vectorización Multi-Rango
- **Objetivo:** Ejecutar tareas de múltiples pasos en un solo viaje COM y formatear celdas dispersas en bloque.
- **Entregables:** Método `curate_pipeline`, soporte de rangos agrupados en `flag_outliers` y subcomando CLI `curate`.
- **Archivos Exclusivos:** `antigravity_excel_core.py`, `antigravity_excel_cli.py`.

### Hito 3 (M3): Servidor Daemon Residente en Memoria (Zero Cold Start)
- **Objetivo:** Eliminar por completo el arranque en frío de Python para agentes y CLI (< 50 ms por respuesta).
- **Entregables:** Módulo `antigravity_excel_daemon.py` con Named Pipe Windows y cliente rápido integrado en el CLI con fallback automático.
- **Archivos Exclusivos:** `antigravity_excel_daemon.py`, `antigravity_excel_cli.py`.

### Hito 4 (M4): Auditoría de Calidad (Critic), Benchmark y Cierre
- **Objetivo:** Certificar 100% de tests aprobados, medir latencias antes vs. después y documentar en `CHANGELOG.md` y `BITACORA_PRUEBAS.md`.
- **Entregables:** Script `tests/benchmark_speed.py`, suite `pytest`, `CHANGELOG.md`, `PRUEBAS/BITACORA_PRUEBAS.md`.
- **Archivos Exclusivos:** `tests/benchmark_speed.py`, `CHANGELOG.md`, `PRUEBAS/BITACORA_PRUEBAS.md`.
