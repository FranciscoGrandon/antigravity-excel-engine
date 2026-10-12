<!-- Creado por Francisco Grandón Vergara -->
# Skill: Benchmarking Harness (Banco de Pruebas y Evaluación de Agentes)

> **Misión:** Establecer un entorno de evaluación riguroso, reproducible y automatizado para cuantificar la competencia de agentes de IA sobre Microsoft Excel utilizando el banco de pruebas canónico en `PRUEBAS/`, evaluando precisión matemática, ahorro de tokens de contexto, latencia IPC y limpieza de procesos en el sistema operativo.

---

## 1. Arquitectura del Banco de Pruebas (`PRUEBAS/`)

El directorio `PRUEBAS/` en la raíz de `EXCEL_ENGINE` contiene una suite de 8 desafíos funcionales diseñados para someter a prueba las capacidades de agentes autónomos bajo condiciones reales de producción:

```
PRUEBAS/
├── 1_Comprension/         # Lectura compartida con bloqueo de usuario, tipado y calidad
├── 2_Preguntas_con_cita/  # Preguntas analíticas con citas exactas de celda (anti-alucinación)
├── 3_Formulas/            # Inyección de columnas calculadas y auto-resize de tablas
├── 4_Tabla_dinamica/      # Creación de tablas dinámicas nativas y gráficos vinculados
├── 5_Escenario/           # Modelado financiero, supuestos ejecutivos y presets contables
├── 6_Limpieza/            # Pipeline de curaduría (dedup, casing, seriales, outliers)
├── 7_Auditoria/           # Detección de errores, fórmulas dinámicas y espacios invisibles
├── 8_Limites_y_seguridad/ # Concurrencia de archivos bloqueados y límites de escala
├── CLAVES/                # Respuestas canónicas y claves de verdad (Ground Truth)
└── BITACORA_PRUEBAS.md    # Registro histórico de auditoría y resultados verificables
```

> [!IMPORTANT]
> **Aislamiento de Claves:** Durante la evaluación ciega de un agente, el directorio `PRUEBAS/CLAVES/` debe permanecer estrictamente inaccesible para el modelo, actuando como oráculo de verificación posterior (*Post-Evaluation Ground Truth*).

---

## 2. Las Cuatro Métricas Esenciales de Evaluación

| Métrica | Definición y Umbral Óptimo | Fórmula / Criterio de Medición |
|---|---|---|
| **1. Precisión de Cálculo** | $\mathbf{100\%}$ de coincidencia numérica exacta contra las claves de verdad. | $\text{Error Absoluto} = |V_{\text{agente}} - V_{\text{ground\_truth}}| = 0$ |
| **2. Eficiencia de Tokens** | Ahorro $\mathbf{> 60\%}$ de tokens en la ventana de contexto usando matrices CSV. | $\text{Ahorro} = 1 - \frac{\text{Tokens}_{\text{CSV}}}{\text{Tokens}_{\text{JSON}}}$ |
| **3. Latencia IPC** | Tiempo de respuesta $\mathbf{< 5\text{ ms}}$ en el canal de comunicación con Excel. | Ping y roundtrip sobre Named Pipe `\\.\pipe\antigravity_excel` |
| **4. Cero Procesos Zombis** | $\mathbf{0}$ instancias huérfanas de `EXCEL.EXE` tras finalizar la sesión. | Conteo de PIDs de Excel antes vs después de la ejecución |

---

## 3. Receta 1: Medición de Precisión Matemática y Citas de Celda

Los LLMs tienden a alucinar agregaciones cuando leen tablas grandes. La prueba `2_Preguntas_con_cita` exige que cualquier afirmación analítica esté respaldada por las coordenadas exactas de la celda de origen.

### Verificación Automatizada vía CLI
```bash
# Calcular ventas agregadas por país con extracción de celdas de citación
python antigravity_excel_cli.py aggregate \
  --file "PRUEBAS/1_Comprension/Financial_Sample.xlsx" \
  --group-by "Country" \
  --metric "Gross Sales" \
  --func "sum" \
  --json
```

**Salida con Respaldo de Citación:**
```json
{
  "metric": "Gross Sales",
  "func": "sum",
  "grouped": {
    "Canada": 24887654.5,
    "Germany": 23505340.0
  },
  "max_cell": "H325",
  "min_cell": "H68",
  "sample_citation_cells": ["H2", "H3", "H4", "H5", "H6"]
}
```
El evaluador compara los valores `grouped` y las celdas reportadas contra `PRUEBAS/CLAVES/` para certificar cero alucinaciones.

---

## 4. Receta 2: Medición de Eficiencia de Tokens (CSV vs JSON)

Para transferir datos matriciales a la ventana de contexto del LLM, el uso de CSV ahorra drásticamente tokens en comparación con JSON:

```bash
# 1. Extracción en formato CSV compacto
python antigravity_excel_cli.py get-csv --range "A1:P20" --sheet "Sheet1"

# 2. Extracción equivalente en JSON anidado
python antigravity_excel_cli.py get-ranges --ranges "A1:P20" --sheet "Sheet1" --json
```

### Tabla Comparativa de Consumo de Tokens (Muestra de 100 Filas $\times$ 16 Columnas)

| Formato | Caracteres Aproximados | Tokens Estimados (tiktoken / cl100k) | Ahorro Relativo |
|---|---|---|---|
| **JSON Anidado** | ~48.200 caracteres | ~11.500 tokens | Base (0%) |
| **CSV Matricial** | ~14.800 caracteres | ~3.850 tokens | **66.5% Ahorro** |

---

## 5. Receta 3: Medición de Latencia IPC y Rendimiento del Daemon

Un agente interactivo no puede esperar 1.2 segundos por cada consulta de celda. El harness evalúa la velocidad del Named Pipe:

```bash
# 1. Comprobar estado y latencia de ping del daemon
python antigravity_excel_cli.py --daemon-status
```

**Diagnóstico emitido:**
```
[DAEMON] Servidor residente ACTIVO en \\.\pipe\antigravity_excel
[DAEMON] Ping exitoso. Latencia IPC: 3.42 ms
```

### Benchmark de Latencia Comparativa (100 Iteraciones)

```
Standalone (Arranque en frío por comando):  ████████████████████ 1.150 ms
Daemon Named Pipe IPC (Memoria RAM):        █ 4,8 ms (240x más rápido)
```

---

## 6. Receta 4: Auditoría de Cero Procesos Zombis

Al automatizar Excel mediante COM en Windows, un error común es olvidar invocar `CoUninitialize()` o dejar referencias circulares que mantienen a `EXCEL.EXE` oculto en memoria.

### Script de Auditoría de Procesos Huérfanos

```powershell
# Verificar procesos de Excel en ejecución antes de la prueba
$pids_iniciales = (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue).Id

# [Ejecución de la suite de pruebas del agente]

# Verificar procesos tras finalizar la prueba
$pids_finales = (Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue).Id
$zombies = $pids_finales | Where-Object { $pids_iniciales -notcontains $_ }

if ($zombies.Count -eq 0) {
    Write-Host "[OK] Tasa de procesos zombis: 0 (Cero fugas de procesos)" -ForegroundColor Green
} else {
    Write-Host "[ALERTA] Se detectaron procesos huérfanos: $zombies" -ForegroundColor Red
}
```

---

## 7. Protocolo de Ejecución del Harness para Evaluación

1. **Paso 1 (Sanitización del Entorno):** Asegurar que no existan instancias huérfanas de `EXCEL.EXE`.
2. **Paso 2 (Inicialización del Daemon):** Levantar el daemon con `python antigravity_excel_cli.py --daemon-start`.
3. **Paso 3 (Despacho de Pruebas 1 a 8):** Ejecutar las tareas de comprensión, fórmulas, pivots, escenarios y curaduría sobre copias en `PRUEBAS/`.
4. **Paso 4 (Auditoría Ciega contra Claves):** Comparar los resultados con los archivos de `PRUEBAS/CLAVES/`.
5. **Paso 5 (Cierre y Telemetría):** Registrar los tiempos de respuesta, consumo de tokens y estado final en la bitácora.
