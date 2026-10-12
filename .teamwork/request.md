# Especificación de Requerimientos y Alineación del Proyecto: Optimización de Velocidad

> **Documento Oficial de Alineación — Fase 1 Teamwork Custom**  
> **Modo de Integridad:** `benchmark`  
> **Directorio:** `D:\Users\francisco.grandon\OneDrive - essbio\Documentos\ANTIGRAVITY\EXCEL_ENGINE`  
> **Autor:** Creado por Francisco Grandón Vergara

---

## 1. Información General y Metadatos
- **Nombre del Proyecto:** `ExcelEngine_Performance_Optimization`
- **Fecha:** `2026-10-10`
- **Sentinel:** Coordinador Antigravity (Hilo Principal)
- **Modo de Integridad:** `benchmark`
- **Estado:** PENDIENTE DE APROBACIÓN DEL USUARIO

---

## 2. Declaración del Problema y Objetivo
- **Contexto:** Las operaciones de automatización y respuesta de `EXCEL_ENGINE` en tareas interactivas toman actualmente entre 1 y 4,7 segundos por comando. El 96% de ese tiempo corresponde a la sobrecarga de arranque del intérprete Python en Windows (arranque en frío ~400 ms), importación masiva estática de módulos pesados (`~726 ms`) y múltiples comandos secuenciales para tareas compuestas.
- **Objetivo:** Optimizar drásticamente la velocidad de respuesta para que cualquier consulta o modificación ejecute en **< 50 ms** en modo residente y en **< 200 ms** en modo CLI sin servidor, introduciendo lazy loading, vectorización multi-rango, pipeline de curación unificado y un servidor daemon local en memoria.

---

## 3. Requisitos Funcionales y Criterios de Aceptación
1. **RF-01 (Carga Perezosa de Módulos):** `antigravity_excel_core.py` no debe importar librerías pesadas (ej. `openpyxl` o subsistemas COM no requeridos) hasta que el método específico sea ejecutado.
2. **RF-02 (Pipeline Unificado `curate`):** Capacidad de ejecutar deduplicación, estandarización de texto, conversión de tipos, parseo de fechas y marcado de outliers en una sola pasada en memoria y una sola transferencia COM.
3. **RF-03 (Vectorización Multi-Rango No Contigua):** Formatear y marcar múltiples celdas disjuntas en una sola llamada COM agrupada (`Range("E325,E412,E689")`).
4. **RF-04 (Servidor Daemon Residente en Memoria):** Subproceso en memoria con Named Pipe Windows (`\\.\pipe\antigravity_excel`) que mantenga caliente la sesión COM y responda en milisegundos.
5. **RF-05 (Compatibilidad Transparente):** Si el daemon no está corriendo, el CLI ejecuta de inmediato en modo standalone sin fallar.
6. **RF-06 (100% Tests Pasando):** La suite completa de `pytest` debe ejecutarse con salida `0` y sin regresiones.
