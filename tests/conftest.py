# Creado por Francisco Grandón Vergara
"""
Configuración global de fixtures y entorno para pruebas unitarias de Antigravity Excel Engine.
Aisla la suite para evitar engancharse a instancias reales de Excel del usuario en el escritorio.
"""

import os
import pytest


@pytest.fixture(autouse=True)
def isolate_excel_suite(request):
    """
    Activa AGY_TEST_ISOLATED=1 durante toda la suite de pruebas para prevenir que
    LiveExcelCOMBackend intente engancharse a ventanas abiertas de Excel en el escritorio.
    Si una prueba declara explícitamente el marcador @pytest.mark.real_com, el aislamiento se desactiva.
    """
    if request.node.get_closest_marker("real_com"):
        old_val = os.environ.get("AGY_TEST_ISOLATED")
        if "AGY_TEST_ISOLATED" in os.environ:
            del os.environ["AGY_TEST_ISOLATED"]
        yield
        if old_val is not None:
            os.environ["AGY_TEST_ISOLATED"] = old_val
    else:
        old_val = os.environ.get("AGY_TEST_ISOLATED")
        os.environ["AGY_TEST_ISOLATED"] = "1"
        yield
        if old_val is not None:
            os.environ["AGY_TEST_ISOLATED"] = old_val
        else:
            os.environ.pop("AGY_TEST_ISOLATED", None)
