# Creado por Francisco Grandón Vergara
"""
Script para sincronizar y commitear todos los archivos modificados y nuevos
hacia GitHub usando la API Git Data (Blobs -> Trees -> Commit -> Ref Update).
"""

import os
import json
import base64
import urllib.request
from pathlib import Path
import keyring

REPO = "FranciscoGrandon/antigravity-excel-engine"
BRANCH = "feature/modular-engine-refactor-and-tests"
API_BASE = f"https://api.github.com/repos/{REPO}"

token = keyring.get_password("EnvVarManagerPro", "GITHUB_TOKEN")
HEADERS = {
    "Authorization": f"token {token}",
    "Content-Type": "application/json",
    "User-Agent": "Antigravity",
    "Accept": "application/vnd.github.v3+json",
}

def api_call(endpoint, data=None, method="GET"):
    url = f"{API_BASE}/{endpoint}"
    payload = json.dumps(data).encode("utf-8") if data is not None else None
    req = urllib.request.Request(url, data=payload, headers=HEADERS, method=method)
    with urllib.request.urlopen(req) as resp:
        return json.loads(resp.read().decode("utf-8"))

def collect_files():
    ignored_dirs = {'.git', '__pycache__', '.pytest_cache', '.ruff_cache', '.mypy_cache', 'venv', 'PRUEBAS', '.idea', '.vscode'}
    root = Path(".")
    files_to_sync = []
    
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if any(part in ignored_dirs for part in path.parts):
            continue
        rel = path.relative_to(root).as_posix()
        # Filtrar temporales o binarios locales
        if rel in {".coverage", "import_core.txt"} or rel.endswith((".pyc", ".xlsx", ".xlsb", ".log")):
            continue
        files_to_sync.append(rel)
    
    return sorted(files_to_sync)

def main():
    print(f"Obteniendo último commit de la rama {BRANCH}...")
    ref_data = api_call(f"git/ref/heads/{BRANCH}")
    latest_commit_sha = ref_data["object"]["sha"]
    print(f"Último commit SHA: {latest_commit_sha}")
    
    commit_data = api_call(f"git/commits/{latest_commit_sha}")
    base_tree_sha = commit_data["tree"]["sha"]
    print(f"Base tree SHA: {base_tree_sha}")

    files = collect_files()
    print(f"Total archivos a procesar: {len(files)}")
    
    tree_items = []
    for idx, rel_path in enumerate(files, 1):
        full_path = Path(rel_path)
        try:
            with open(full_path, "rb") as f:
                content_bytes = f.read()
            b64_content = base64.b64encode(content_bytes).decode("utf-8")
            
            # Crear blob
            blob_resp = api_call("git/blobs", {"content": b64_content, "encoding": "base64"}, method="POST")
            blob_sha = blob_resp["sha"]
            
            tree_items.append({
                "path": rel_path,
                "mode": "100644",
                "type": "blob",
                "sha": blob_sha
            })
            if idx % 10 == 0 or idx == len(files):
                print(f"[{idx}/{len(files)}] Blob creado para {rel_path} -> {blob_sha[:8]}")
        except Exception as e:
            print(f"Error en {rel_path}: {e}")

    print(f"\nCreando nuevo Git Tree con {len(tree_items)} elementos...")
    new_tree = api_call("git/trees", {"base_tree": base_tree_sha, "tree": tree_items}, method="POST")
    new_tree_sha = new_tree["sha"]
    print(f"Nuevo Tree SHA: {new_tree_sha}")

    commit_message = """feat: refactor modular antigravity_excel_engine, optimizaciones matriciales y suite de 87 tests

- Modularización de arquitectura en 6 subpaquetes: backends, formatting, formulas, ipc, models, operations y utils.
- Aceleración matricial 2D en Live COM para auditoría de calidad de datos (>300x de velocidad).
- Incorporación de presets financieros y diseño adaptativo sin solapamiento para tablas dinámicas y gráficos.
- Serialización segura Decimal/datetime en IPC Named Pipe y CLI.
- Suite de pruebas unitarias expandida a 87 tests con 100% de aprobación.
- Actualización de documentación (README.md y CHANGELOG.md).
"""
    print("Creando commit...")
    new_commit = api_call("git/commits", {
        "message": commit_message,
        "tree": new_tree_sha,
        "parents": [latest_commit_sha]
    }, method="POST")
    new_commit_sha = new_commit["sha"]
    print(f"Nuevo Commit SHA: {new_commit_sha}")

    print(f"Actualizando referencia de la rama {BRANCH}...")
    api_call(f"git/refs/heads/{BRANCH}", {"sha": new_commit_sha}, method="PATCH")
    print(f"¡Rama {BRANCH} actualizada exitosamente!")

if __name__ == "__main__":
    main()
