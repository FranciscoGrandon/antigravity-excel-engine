<!-- Creado por Francisco Grandón Vergara -->
# Skill: Dispatch COM Advanced (Invocación Dinámica de la API Profunda)

> **Misión:** Proporcionar a los agentes de IA un canal universal y reflexivo para ejecutar cualquier método o propiedad del Modelo de Objetos de Microsoft Excel sin requerir nuevos comandos de CLI compilados en el motor, optimizando la gestión de memoria COM y previniendo procesos zombis.

---

## 1. Arquitectura de `dispatch_com`

La primitiva `dispatch_com` es un puente de invocación dinámico expuesto tanto en la fachada `AntigravityExcelEngine` como en el CLI (`dispatch-com`) y el servidor IPC Daemon. Permite navegar jerarquías de objetos, ejecutar métodos con argumentos posicionales o por nombre, y leer o escribir atributos arbitrarios.

```
 Agente (CLI / IPC) 
   │
   ├─► target: "app" | "wb" | "sheet:<nombre>" | "range:<coordenada>" | "prop.subprop"
   ├─► member: "NombreMetodo" | "NombrePropiedad"
   └─► args / kwargs: Parámetros serializados en JSON
           │
           ▼
   [LiveExcelCOMBackend._resolve_target]
           │
           ▼
   [Invocación reflexiva con @retry_on_excel_busy]
           │
           ▼
   [_serialize_com_result]  ──► Serializa tipos COM a primitivas nativas (JSON seguro)
```

### Tabla de Resolución de Destinos (`target`)

| Sintaxis de `target` | Objeto COM Resuelto |
|---|---|
| `"app"` o `"excel"` | `Excel.Application` |
| `"wb"` o `"workbook"` | `Workbook` activo |
| `"sheet"` o `"activesheet"` | `Worksheet` activo |
| `"sheet:NombreHoja"` | `wb.Sheets("NombreHoja")` |
| `"range:A1:D10"` | `ActiveSheet.Range("A1:D10")` |
| `"wb.Worksheets('Data').PageSetup"` | Resolución encadenada de propiedades COM |

---

## 2. Receta 1: Autofiltros Dinámicos (AutoFilter)

El autofiltro permite al agente aislar filas específicas programáticamente sin alterar los datos del libro.

### Vía CLI (Named Pipe Daemon <5ms)
```bash
# Aplicar autofiltro en la columna 1 (Segment) para el valor "Government"
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Sheet1" \
  --member "Range('A1:P701').AutoFilter" \
  --args-json '[1, "Government"]' \
  --json
```

### Quitar Filtros Activos
```bash
# Mostrar todos los datos sin eliminar las flechas de filtro
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Sheet1" \
  --member "ShowAllData" \
  --json
```

---

## 3. Receta 2: Protección y Desprotección de Hojas con Contraseña

Permite al agente bloquear fórmulas y estructuras para evitar alteraciones accidentales por parte de usuarios.

### Proteger Hoja
```bash
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Sheet1" \
  --member "Protect" \
  --kwargs-json '{"Password": "MasterPass2026", "DrawingObjects": true, "Contents": true, "Scenarios": true, "AllowFiltering": true}' \
  --json
```

### Desproteger Hoja
```bash
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Sheet1" \
  --member "Unprotect" \
  --kwargs-json '{"Password": "MasterPass2026"}' \
  --json
```

---

## 4. Receta 3: Validación de Datos y Listas Desplegables

Insertar reglas de validación en celdas para forzar que el usuario sólo seleccione valores de una lista predefinida:

### Vía Python COM con Limpieza Transaccional
```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="live")
sheet = engine.backend.wb.Sheets("Sheet1")
target_range = sheet.Range("D2:D50")

# 1. Eliminar validaciones previas para evitar conflictos COM
target_range.Validation.Delete()

# 2. Agregar validación de lista (xlValidateList = 3, xlValidAlertStop = 1)
target_range.Validation.Add(
    Type=3,           # xlValidateList
    AlertStyle=1,     # xlValidAlertStop
    Operator=1,       # xlBetween
    Formula1="High,Medium,Low,None"
)
target_range.Validation.IgnoreBlank = True
target_range.Validation.InCellDropdown = True
```

---

## 5. Receta 4: Exportación Nativa a Formato PDF

Exportar hojas ejecutivas o informes de auditoría a PDF sin depender de impresoras virtuales de terceros:

### Vía CLI de Alta Velocidad
```bash
# Exportar la hoja activa a PDF usando el motor nativo de Office (xlTypePDF = 0)
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Sheet1" \
  --member "ExportAsFixedFormat" \
  --kwargs-json '{"Type": 0, "Filename": "D:\\Users\\francisco.grandon\\Documentos\\Reporte_Financiero.pdf", "Quality": 0, "IncludeDocProperties": true, "IgnorePrintAreas": false, "OpenAfterPublish": false}' \
  --json
```

**Respuesta JSON esperada:**
```json
{
  "success": true,
  "result": null,
  "target": "sheet:Sheet1",
  "member": "ExportAsFixedFormat"
}
```

---

## 6. Catálogo de Trampas Comunes (*Traps*) y Remedios

### Trampa 1: El Objeto COM Huérfano y el Proceso `EXCEL.EXE` Zombi
* **Síntoma:** El script termina, pero en el Administrador de Tareas de Windows quedan múltiples procesos `EXCEL.EXE` consumiendo memoria y bloqueando archivos.
* **Causa Raíz:** En Python, `win32com` crea punteros `IDispatch` envueltos en objetos Python. Si un agente asigna `rng = sheet.Range("A1").Font.Color` o encadena accesos de puntos sin liberar las referencias intermedias, el contador de referencias de Windows no llega a cero y el proceso de Excel se niega a cerrarse.
* **Solución de Arquitectura:**
  1. `dispatch_com` serializa inmediatamente los resultados a tipos primitivos nativos (`str`, `int`, `float`, `list`, `dict`) y elimina referencias intermedias con `del res`.
  2. Al cerrar el motor, ejecutar recolección forzada:
  ```python
  import gc
  gc.collect() # Dispara el decremento de referencias IDispatch
  ```

### Trampa 2: Nombres de Hoja con Espacios o Caracteres Reservados
* **Síntoma:** Al invocar `--target "sheet:Ventas 2026"` o evaluar expresiones con espacios, Excel lanza error `0x800A03EC`.
* **Causa:** En la sintaxis de rangos o evaluación de Excel, las hojas con espacios deben estar entre comillas simples (`'Ventas 2026'!A1`).
* **Solución:**
  - `LiveExcelCOMBackend._resolve_target` maneja automáticamente la extracción limpia del nombre tras el prefijo `sheet:` usando `wb.Sheets("Ventas 2026")`, que acepta espacios sin requerir comillas simples.
  - Al inyectar fórmulas dinámicas cruzadas, envolver siempre el nombre de la hoja entre comillas simples:
    `='Ventas 2026'!$B$5`

### Trampa 3: Incompatibilidad de Retorno de Tipos COM Variantes
* **Síntoma:** Al leer una propiedad como `Range.Value` sobre un rango de múltiples celdas, COM retorna una tupla bidireccional de tuplas `((1, 2), (3, 4))` que no es directamente convertible por algunos serializadores JSON si contiene fechas (`pywintypes.TimeType`).
* **Solución:**
  - El helper interno `_serialize_com_result` convierte recursivamente todas las tuplas a listas estándar de Python y transforma cualquier objeto fecha o `TimeType` a formato canónico ISO-8601 (`YYYY-MM-DDTHH:MM:SS`), garantizando salidas JSON deterministas.

---

## 7. Checklist de Seguridad para Invocaciones COM Dinámicas

- [ ] ¿Se verificó si el objetivo (`target`) existe antes de invocar métodos que mutan el estado?
- [ ] ¿Se especificaron todos los parámetros booleanos y constantes como enteros o booleanos nativos de Python (`true` / `false` en JSON)?
- [ ] ¿Se liberaron referencias a objetos COM intermedios para prevenir procesos `EXCEL.EXE` huérfanos?
- [ ] ¿Se envolvió la llamada con `@retry_on_excel_busy` para tolerar interacciones simultáneas del usuario?
