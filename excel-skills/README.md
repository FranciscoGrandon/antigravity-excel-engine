<!-- Creado por Francisco Grandón Vergara -->
# Excel Skills: Catálogo de Habilidades de Interacción para Agentes Autónomos de IA

> **Inspirado en la arquitectura modular de `interaction-skills` de browser-use**, adaptado al ecosistema de Microsoft Excel y guiado por la filosofía de **máxima velocidad, consistencia transaccional y cero corrupción** de `excel_engine`.

---

## 1. Visión General y Filosofía de Diseño

Cuando un modelo de lenguaje (LLM) o un agente autónomo interactúa con hojas de cálculo en producción, enfrenta desafíos análogos a los de la automatización web moderna: bloqueos de concurrencia, estados dinámicos no deterministas, desfases regionales de sintaxis y penalizaciones masivas de latencia si no se optimiza el canal de comunicación.

La biblioteca **Excel Skills** formaliza el conocimiento operativo, patrones de ejecución, recetas comprobadas y trampas críticas (*traps*) necesarias para que un agente de IA manipule libros de Excel en vivo (`Live COM`) o en segundo plano (`Headless OpenPyXL`) de manera confiable, ultrarrápida y segura.

```
                      ┌───────────────────────────────────────┐
                      │    Agente de IA / Subagente LLM       │
                      └──────────────────┬────────────────────┘
                                         │ JSON IPC / CLI
                                         ▼
                      ┌───────────────────────────────────────┐
                      │     Daemon IPC (Named Pipe)           │ < 5 ms
                      │     \\.\pipe\antigravity_excel        │
                      └──────────────────┬────────────────────┘
                                         │ En memoria
                         ┌───────────────┴───────────────┐
                         ▼                               ▼
              ┌─────────────────────┐         ┌─────────────────────┐
              │ LiveExcelCOMBackend │         │HeadlessOpenPyXLBack.│
              │  (Win32 COM / OLE)  │         │  (Memoria / XML)    │
              └──────────┬──────────┘         └──────────┬──────────┘
                         │                               │
                         ▼                               ▼
                 Microsoft Excel                  Disco / Red Local
               (Sesión de usuario)               (SGL / Parquet / FS)
```

---

## 2. Las Tres Reglas de Oro

### Regla 1: Latencia Sub-5ms vía Daemon Named Pipes
- **Prohibido el arranque en frío iterativo:** Invocar `python antigravity_excel_cli.py` en cada paso sin daemon penaliza al agente con 800ms – 1.500ms por comando debido al costo de importación de módulos y enganche COM.
- **Canal IPC Residente:** Con `antigravity_excel_daemon.py` activo, las peticiones viajan por el Named Pipe `\\.\pipe\antigravity_excel`, ejecutándose en memoria en **~3ms a ~5ms** sin recargar bibliotecas ni reinicializar punteros COM.

### Regla 2: Cero Corrupción y Resguardo Transaccional
- **Sin huérfanos ni celdas rotas:** Toda inyección o transformación sobre tablas estructuradas (`ListObjects`) debe auto-ajustar su rango geométrico (`ListObject.Resize`), sanitizar prefijos XML (`_xlfn.`) y preservar estilos preexistentes.
- **Aislamiento de Escritura:** Emplear `FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE` en lecturas de archivos abiertos para evitar excepciones de `PermissionError [Errno 13]`.

### Regla 3: Ahorro Radical de Tokens de Contexto
- **CSV Matricial sobre JSON:** Volcar o leer matrices en formato CSV crudo ahorra más del 65% de tokens en comparación con estructuras JSON anidadas, permitiendo al agente procesar libros de miles de filas sin saturar la ventana de contexto.

---

## 3. Catálogo e Índice de Habilidades

| Skill | Archivo | Propósito Central | Cuándo Usarla (Gatillos) |
|---|---|---|---|
| **COM Locks & Readiness** | [`com-locks-and-readiness.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/excel-skills/com-locks-and-readiness.md) | Sincronización adaptativa, gestión de `RPC_E_SERVERCALL_RETRYLATER` y control de `excel_fast_mode`. | El usuario tiene Excel abierto, edita celdas activas, o se ejecutan operaciones masivas que requieren suspender recálculos. |
| **Formulas & Dynamic Spill** | [`formulas-dynamic-spill.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/excel-skills/formulas-dynamic-spill.md) | Inyección de matrices dinámicas (`UNIQUE`, `FILTER`, `XLOOKUP`), notación `@` y prevención de error `#SPILL!`. | Inserción de fórmulas vectorizadas, columnas calculadas en tablas estructuradas, o internacionalización de fórmulas (`;` vs `,`). |
| **Dates & Serial Numbers** | [`dates-and-serial-numbers.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/excel-skills/dates-and-serial-numbers.md) | Conversión bidireccional ultrarrápida ISO-8601 $\leftrightarrow$ Serial Excel y manejo de anomalías temporales. | Coerción de fechas en texto, filtros temporales, cálculo de vencimientos o importación de datasets con timestamps. |
| **Dispatch COM Advanced** | [`dispatch-com-advanced.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/excel-skills/dispatch-com-advanced.md) | Invocación reflexiva de la API COM profunda para autofiltros, protección, validación de datos y exportación PDF. | Tareas avanzadas fuera del CLI estándar: proteger hojas con contraseña, agregar listas desplegables, exportar reportes ejecutivos. |
| **Benchmarking Harness** | [`benchmarking-harness.md`](file:///d:/Users/francisco.grandon/OneDrive%20-%20essbio/Documentos/ANTIGRAVITY/EXCEL_ENGINE/excel-skills/benchmarking-harness.md) | Evaluación estandarizada de capacidades de agentes de IA sobre Excel usando el banco de pruebas de `PRUEBAS/`. | Validación de modelos LLM, regresión de rendimiento, auditoría de precisión de cálculo y verificación de cero procesos zombis. |

---

## 4. Cómo Consumir las Skills en Agentes de IA

1. **Invocación Modular:** Cuando un agente detecta un problema específico (por ejemplo, Excel no responde porque el usuario está editando una celda), debe consultar de inmediato el archivo correspondiente en `excel-skills/` para aplicar la receta canónica y evitar trampas documentadas.
2. **Prioridad CLI y Daemon:** Las recetas prefieren siempre la invocación mediante `antigravity_excel_cli.py` con daemon activo antes que scripts ad-hoc en Python, garantizando latencias <5ms y telemetría estructurada.
3. **Manejo de Errores Determinista:** Cada habilidad documenta las excepciones COM canónicas (`0x8001010A`, `0x800AC472`, `0x80010001`) y el árbol de decisión de resiliencia asociado.
