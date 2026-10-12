<!-- Creado por Francisco Grandón Vergara -->
# Skill: Formulas & Dynamic Spill (Fórmulas Vectorizadas y Matrices Dinámicas)

> **Misión:** Guiar la inyección, propagación y evaluación de fórmulas modernas de matriz dinámica (`UNIQUE`, `FILTER`, `SORT`, `XLOOKUP`), notación estructurada de tablas y operadores de rango desbordado (`#`), evitando bloqueos por error `#SPILL!` e inconsistencias idiomáticas.

---

## 1. La Revolución de las Matrices Dinámicas en Excel 365+

En versiones modernas de Microsoft Excel, las fórmulas ya no están confinadas a una sola celda. Una fórmula que devuelve una colección de valores desborda (*spills*) automáticamente sobre las celdas contiguas adyacentes.

### Principales Funciones de Matriz Dinámica

| Función | Propósito | Ejemplo Canónico |
|---|---|---|
| `UNIQUE` | Extrae valores únicos o distintos sin duplicados. | `=UNIQUE(financials[Country])` |
| `SORT` | Ordena un rango o matriz por una o más columnas. | `=SORT(UNIQUE(financials[Segment]))` |
| `FILTER` | Filtra una matriz con base en condiciones booleanas. | `=FILTER(financials[Sales], financials[Country]="Canada")` |
| `XLOOKUP` | Búsqueda bidireccional exacta sin limitación de columna izquierda. | `=XLOOKUP(E2, Products[ID], Products[Price], 0)` |
| `SEQUENCE` | Genera una secuencia matricial de enteros. | `=SEQUENCE(12, 1, 1, 1)` |

### Operadores Fundamentales: `@` y `#`

1. **Operador de Rango Desbordado (`#`):**  
   Permite hacer referencia a la totalidad de una matriz dinámica viva dondequiera que se expanda.
   - Si la celda `H2` contiene `=UNIQUE(A2:A500)`, la fórmula `=COUNTA(H2#)` o `=SUM(H2#)` referenciará automáticamente todo el rango resultante, sin importar cuántas filas ocupe.
2. **Operador de Intersección Implícita (`@`):**  
   Fuerza la evaluación a un valor escalar único para la fila actual.
   - `Table[@Units Sold]`: Valor de la columna *Units Sold* en la fila actual de la tabla estructurada.
   - Si Excel 365 inserta un `@` indeseado al inicio de una fórmula (`=@UNIQUE(...)`), indica que la fórmula fue inyectada mediante la propiedad heredada `.Formula` en lugar de `.Formula2`.

---

## 2. Receta 1: Inyección Segura de Matrices Dinámicas con `Formula2`

En la API COM de Microsoft Excel:
- `.Formula`: Modo heredado pre-2018. Si se asigna una fórmula que genera una matriz, Excel asume compatibilidad retrospectiva e inyecta silenciosamente el operador `@` (`=@UNIQUE(...)`), impidiendo el desbordamiento.
- `.Formula2`: Propiedad nativa de matrices dinámicas en Excel 365. Respeta el desbordamiento y evalúa la fórmula como un vector vivo.

### Código Python / COM Canónico

```python
# Inyectar una matriz dinámica que desborde de forma natural
cell = sheet.Range("J2")
# CRÍTICO: Usar Formula2 para permitir desbordamiento (spill)
cell.Formula2 = "=SORT(UNIQUE(financials[Segment]))"

# Referenciar el rango desbordado completo en otra celda
sheet.Range("L1").Formula2 = "=COUNTA(J2#)"
```

### Vía CLI de Antigravity
```bash
# Inyección vía dispatch-com de alta velocidad
python antigravity_excel_cli.py dispatch-com \
  --target "sheet:Resumen" \
  --member "Range('B5').Formula2" \
  --args-json "[\"=UNIQUE(financials[Country])\"]"
```

---

## 3. Receta 2: Fórmulas en Tablas Estructuradas (`ListObjects`)

Al trabajar dentro de una tabla de Excel con nombre (`financials`, `Ventas2026`), las fórmulas deben emplear **Notación Estructurada** en lugar de coordenadas fijas (`A2:A701`):

```excel
-- Referencia a columna completa (vectorial):
=SUM(financials[Gross Sales])

-- Referencia a la celda de la fila actual:
=[@[Sale Price]] * [@[Units Sold]]

-- Referencia condicional combinada:
=IF([@Discounts] > 0, [@[Gross Sales]] - [@Discounts], [@[Gross Sales]])
```

### Inyección de Columnas Calculadas con Auto-Resize
Al agregar columnas a una tabla mediante `add_calculated_column`:
```bash
python antigravity_excel_cli.py add-column \
  --header "Margin_Pct" \
  --formula "=[@Profit] / [@[Sales]]" \
  --number-format "0.0%" \
  --sheet "Sheet1"
```
El motor inserta la fórmula en el encabezado expandido, Excel propaga automáticamente la fórmula calculada a todas las filas de la tabla estructurada y redimensiona su rango sin dejar huecos.

---

## 4. Catálogo de Trampas Comunes (*Traps*) y Remedios

### Trampa 1: El Error `#SPILL!` (#¡DESBORDAMIENTO!)
* **Síntoma:** La celda muestra `#SPILL!` y una línea punteada que rodea un rango vacío contiguo.
* **Causa Raíz:** Hay al menos una celda **no vacía** en el área geométrica que la matriz necesita para expandirse. Incluso un espacio en blanco (`" "`), una cadena de longitud cero (`""`) o una celda con formato residual impedirá el desbordamiento.
* **Diagnóstico y Solución:**
  1. Identificar el tamaño potencial de la matriz.
  2. Limpiar el rango de expansión antes de escribir la fórmula:
  ```python
  # Limpiar rango contiguo antes de inyectar matriz dinámica
  sheet.Range("H2:H100").ClearContents()
  sheet.Range("H2").Formula2 = "=UNIQUE(A2:A500)"
  ```

### Trampa 2: La Torre de Babel de Fórmulas (`.Formula` vs `.FormulaLocal`)
* **Síntoma:** Asignar `=BUSCARX(A2, B:B, C:C)` provoca error COM o celda con `#¿NOMBRE?`. Asignar `=IF(A2>0; "SI"; "NO")` en `.Formula` falla con error de sintaxis COM.
* **Causa Raíz:**
  - `.Formula` y `.Formula2` **exigen siempre nombres en inglés y coma (`,`) como separador de argumentos**, sin importar el idioma del sistema operativo.
  - `.FormulaLocal` exige los nombres traducidos del idioma regional del Excel instalado (ej. `BUSCARX` en español) y punto y coma (`;`) si la configuración regional de Windows usa coma decimal.
* **Regla de Oro para Agentes de IA:**
  - **Escribir SIEMPRE fórmulas en INGLÉS estándar usando coma (`,`)** y asignarlas a `.Formula2` o `.Formula`.
  - Excel traducirá automáticamente la interfaz visual para el usuario humano (un usuario con Excel en español verá `=BUSCARV` aunque el agente haya inyectado `=VLOOKUP`).

```python
# CORRECTO (Universal e independiente del idioma de Windows)
cell.Formula2 = "=IF(A2>100, \"Alto\", \"Bajo\")"

# INCORRECTO (Rompe en Excel en inglés o configuraciones con coma decimal)
# cell.Formula = "=SI(A2>100; \"Alto\"; \"Bajo\")"
```

### Trampa 3: Fórmulas Dependientes No Recalculadas
* **Síntoma:** El agente escribió fórmulas, luego leyó los valores vía `get-csv` o `get-ranges` y los valores aparecen en blanco, como `0`, o conservan el resultado anterior.
* **Causa:** En modo `Live COM` bajo `excel_fast_mode`, `app.Calculation = xlCalculationManual` suspende el motor de dependencias hasta salir del bloque. En backend `Headless OpenPyXL`, **OpenPyXL no tiene motor de cálculo**: no puede evaluar fórmulas y devolverá `None` en `.value` a menos que se use `data_only=False` (que retorna el string de la fórmula) o se ejecute en `Live COM`.
* **Solución:**
  - Para leer valores calculados de fórmulas vivas, utilizar obligatoriamente el modo `Live COM` con `app.Calculate()`:
  ```python
  app.CalculateFull() # Fuerza recálculo completo de la cadena de dependencias
  ```

---

## 5. Resumen de Prácticas Óptimas

1. Usa siempre `.Formula2` para habilitar el desbordamiento dinámico sin `@` forzado.
2. Limpia el rango de destino con `.ClearContents()` antes de asignar matrices expansibles.
3. Escribe fórmulas en **inglés y con comas** (`=FILTER(A:A, B:B="X")`), dejando que Excel maneje la localización visual.
4. Usa notación estructurada `Table[@Columna]` al operar dentro de tablas para mantener robustez ante inserciones de filas o columnas.
