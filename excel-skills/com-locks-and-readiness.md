<!-- Creado por Francisco Grandón Vergara -->
# Skill: COM Locks & Readiness (Sincronización y Tolerancia a Bloqueos)

> **Misión:** Garantizar que los agentes de IA interactúen con Microsoft Excel sin colapsar por bloqueos interactivos de usuario, diálogos modales o cálculos pesados en segundo plano, maximizando el rendimiento mediante `excel_fast_mode` sin corromper el estado de la aplicación.

---

## 1. El Problema de la Concurrencia Humano-Agente

A diferencia de las bases de datos transaccionales tradicionales, Microsoft Excel en modo interactivo (`Live COM`) es un entorno mono-hilo con una interfaz de usuario compartida. Cuando un usuario humano:
1. **Tiene el cursor dentro de la barra de fórmulas o editando una celda** (modo *Edit/Point*).
2. **Mantiene abierto un cuadro de diálogo modal** (ej. *Formato de Celdas*, *Buscar y Reemplazar*, un mensaje de alerta de VBA).
3. **Ejecuta un recálculo masivo o una consulta Power Query**.

El servidor OLE/COM de Excel rechaza inmediatamente cualquier llamada entrante desde la automatización externa lanzando excepciones RPC nativas de Windows.

### Tabla Canónica de HRESULTs de Bloqueo

| HRESULT (Hex) | HRESULT (Dec) | Nombre Simbólico | Causa Raíz | Acción del Agente |
|---|---|---|---|---|
| `0x8001010A` | `-2147417846` | `RPC_E_SERVERCALL_RETRYLATER` | Usuario editando celda activa o barra de fórmulas. | Reintentar con backoff exponencial o invocar `wait_until_ready`. |
| `0x800AC472` | `-2146777998` | `VBA_E_IGNORE` | Excel ocupado redibujando pantalla o en ciclo interno. | Pausa breve (200-500ms) y reintento. |
| `0x80010001` | `-2147418111` | `RPC_E_CALL_REJECTED` | Diálogo modal activo o servidor en suspensión. | Esperar cierre de ventana o notificar al usuario. |

---

## 2. Receta 1: Sincronización Previa con `wait_until_ready`

Antes de iniciar cualquier secuencia crítica de lecturas o escrituras, el agente debe verificar el estado de preparación de Excel.

### Vía CLI (Recomendado para Agentes)
```bash
# Sondeo con timeout de 10 segundos e intervalo de 250ms
python antigravity_excel_cli.py wait-until-ready --timeout 10.0 --interval 0.25 --json
```

**Respuesta JSON esperada:**
```json
{
  "ready": true,
  "elapsed_sec": 0.0124,
  "calculation_state": "done"
}
```

Si el usuario permanece editando y expira el tiempo:
```json
{
  "ready": false,
  "timeout": true,
  "elapsed_sec": 10.001
}
```

### Vía Python Core
```python
from antigravity_excel_core import AntigravityExcelEngine

engine = AntigravityExcelEngine(mode="live")
status = engine.wait_until_ready(timeout_sec=5.0, retry_interval=0.2)
if not status["ready"]:
    # Informar al usuario o diferir tarea sin abortar con traceback
    print("Excel se encuentra ocupado por interacción del usuario.")
```

---

## 3. Receta 2: Decorador de Resiliencia `@retry_on_excel_busy`

Para proteger métodos que ejecutan operaciones COM directas contra interrupciones instantáneas mientras el usuario tipea o navega en la hoja:

```python
import functools
import time
import win32com.client
import pywintypes

BUSY_HRESULTS = (-2147417846, -2147418111, -2146777998, 0x8001010A, 0x80010001, 0x800AC472)

def retry_on_excel_busy(max_retries: int = 10, delay_sec: float = 0.5):
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            retries = 0
            while retries < max_retries:
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    is_busy = False
                    if pywintypes and isinstance(e, pywintypes.com_error):
                        hresult = getattr(e, 'hresult', None) or (e.args[0] if e.args else None)
                        if hresult in BUSY_HRESULTS:
                            is_busy = True
                    if is_busy:
                        retries += 1
                        time.sleep(delay_sec)
                        continue
                    raise
            raise RuntimeError(
                f"Excel permaneció bloqueado en modo edición tras {max_retries * delay_sec:.1f}s. "
                f"Presiona [ENTER] o [ESC] en Excel para liberar el cursor."
            )
        return wrapper
    return decorator
```

---

## 4. Receta 3: Modo de Alto Rendimiento con `excel_fast_mode`

Cada actualización visual de celda o recálculo intermedio genera un costo de CPU inmenso. `excel_fast_mode` es un context manager que congela la interfaz y suspende los cálculos automáticos durante operaciones por lotes, restaurando el entorno de forma transaccional al finalizar.

### Implementación Canónica

```python
from contextlib import contextmanager

@contextmanager
def excel_fast_mode(app):
    if not app:
        yield
        return

    # Capturar estados previos para restauración precisa
    prev_screen_updating = getattr(app, "ScreenUpdating", True)
    prev_display_alerts = getattr(app, "DisplayAlerts", True)
    prev_enable_events = getattr(app, "EnableEvents", True)
    prev_calc = getattr(app, "Calculation", -4105) # xlCalculationAutomatic

    try:
        app.ScreenUpdating = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.Calculation = -4135 # xlCalculationManual
    except Exception:
        pass

    try:
        yield
    finally:
        try:
            # Restauración obligatoria y forzado de cálculo consistente
            app.Calculation = prev_calc
            app.EnableEvents = prev_enable_events
            app.DisplayAlerts = prev_display_alerts
            app.ScreenUpdating = prev_screen_updating
            app.Calculate() # Dispara el recálculo íntegro del libro
        except Exception:
            pass
```

---

## 5. Catálogo de Trampas Comunes (*Traps*) y Remedios

### Trampa 1: El Síndrome de la Pantalla Congelada (Fuga de `ScreenUpdating`)
* **Síntoma:** El agente ejecutó un script, falló con una excepción no controlada, y ahora el usuario ve que Excel no responde visualmente a clics ni actualiza la pantalla al escribir.
* **Causa:** Se asignó `app.ScreenUpdating = False` sin bloque `try...finally`. Al romper la ejecución, Excel queda congelado de por vida hasta reiniciar el proceso.
* **Remedio:** **NUNCA** modificar `ScreenUpdating` o `Calculation` directamente en código plano. Usar siempre `with excel_fast_mode(self.app):`. Si ya ocurrió, ejecutar desde CLI:
  ```bash
  python antigravity_excel_cli.py dispatch-com --target app --member ScreenUpdating --args-json "[true]"
  python antigravity_excel_cli.py dispatch-com --target app --member Calculation --args-json "[-4105]"
  ```

### Trampa 2: La Celda Bloqueada por el Cursor Humano
* **Síntoma:** Llamadas COM fallan instantáneamente con `pywintypes.com_error: (-2147417846, 'La llamada fue rechazada por el destinatario.', None, None)`.
* **Causa:** El usuario hizo doble clic en una celda y dejó el cursor parpadeando. Excel entra en estado modal y bloquea la API COM.
* **Remedio:** No reintentar en bucle infinito sin pausa. `wait_until_ready` sondea suavemente. Si tras 10 segundos persiste, el agente debe emitir un mensaje amigable:
  > *"Excel está en modo edición de celda. Por favor presiona Enter o Escape para continuar la automatización."*

### Trampa 3: Diálogos Modales Abiertos (ej. Cuadro 'Guardar como')
* **Síntoma:** Llamadas COM retornan error de servidor ocupado indefinidamente.
* **Causa:** Hay un diálogo modal de Windows abierto en la ventana de Excel.
* **Remedio:** Detectar si el proceso principal tiene ventanas secundarias modales activas mediante Win32 API (`win32gui.GetWindow`), o notificar al usuario para que cierre el diálogo pendiente.

---

## 6. Lista de Verificación para Agentes Autónomos

- [ ] ¿Invoqué `wait_until_ready` antes de una ráfaga de modificaciones?
- [ ] ¿Toda mutación masiva está envuelta en `with excel_fast_mode(app):`?
- [ ] ¿Está garantizada la restauración de `Calculation = xlCalculationAutomatic` y `ScreenUpdating = True` en el bloque `finally`?
- [ ] ¿Se capturan los HRESULTs `0x8001010A` y `0x800AC472` mediante reintentos con retraso antes de declarar una falla fatal?
