# backend/run_dev.py
"""
Arranque standalone del backend para desarrollo, sin Electron.

Uso:
    python run_dev.py           # puerto por defecto 8765
    python run_dev.py 9000      # puerto personalizado

Con reload=True para que se refresque al guardar archivos.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Asegurar que backend/ esté en sys.path
BASE_DIR = Path(__file__).resolve().parent
if str(BASE_DIR) not in sys.path:
    sys.path.insert(0, str(BASE_DIR))

import uvicorn  # noqa: E402


DEFAULT_PORT = 8765


def main() -> None:
    puerto = DEFAULT_PORT
    if len(sys.argv) >= 2:
        try:
            puerto = int(sys.argv[1])
        except ValueError:
            print(f"Puerto invalido: {sys.argv[1]}. Usando {DEFAULT_PORT}.")
            puerto = DEFAULT_PORT

    print("=" * 60)
    print(f"  SFA / Facturacion - Backend dev")
    print(f"  http://127.0.0.1:{puerto}")
    print(f"  Docs interactivos: http://127.0.0.1:{puerto}/docs")
    print("  Ctrl+C para detener.")
    print("=" * 60)
    print()

    uvicorn.run(
        "api.server:app",
        host="127.0.0.1",
        port=puerto,
        reload=True,
        log_level="info",
    )


if __name__ == "__main__":
    main()
