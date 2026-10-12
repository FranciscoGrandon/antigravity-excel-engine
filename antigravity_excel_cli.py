# Creado por Francisco Grandón Vergara
"""
Shim de compatibilidad retroactiva para antigravity_excel_cli.
Re-exporta el CLI modular desde antigravity_excel_engine.cli.

Los scripts que llamen a este archivo seguirán funcionando.
El punto de entrada antigravity-excel registrado en pyproject.toml
también continúa funcional.
"""

from antigravity_excel_engine.cli import main, build_parser, load_json_or_file, _try_daemon

# Alias de compatibilidad retroactiva para tests y scripts existentes
_try_call_daemon = _try_daemon

__all__ = ["main", "build_parser", "load_json_or_file", "_try_call_daemon", "_try_daemon"]

if __name__ == "__main__":
    main()
