# backend/api/server.py
"""
Aplicación FastAPI del backend SFA / Facturación.

Expone los endpoints que consume la UI (Electron / React).
Bindea solo a 127.0.0.1, no es accesible desde red externa.

Mantiene intacto el código de `core/` (business logic original).
"""

from __future__ import annotations

import logging
import sys
from pathlib import Path

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

# --------------------------------------------------------------------
# sys.path: permitir imports sin prefijo desde core/ (preserva imports
# existentes como `from utils_sfa import X` sin tener que modificar
# los archivos originales de business logic).
# --------------------------------------------------------------------
BASE_DIR = Path(__file__).resolve().parent.parent  # .../backend
CORE_DIR = BASE_DIR / "core"

for extra in (BASE_DIR, CORE_DIR):
    p = str(extra)
    if p not in sys.path:
        sys.path.insert(0, p)

# --------------------------------------------------------------------
# Imports internos (DESPUÉS de configurar sys.path)
# --------------------------------------------------------------------
from api.auth import AuthMiddleware  # noqa: E402
from api.routes import ping  # noqa: E402

# --------------------------------------------------------------------
# Logging
# --------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(levelname)s %(name)s: %(message)s",
)
logger = logging.getLogger("sfa.api")


def create_app() -> FastAPI:
    app = FastAPI(
        title="SFA / Facturación API",
        version="0.1.0",
        description="Backend local para la app Electron.",
    )

    # CORS abierto: el cliente es siempre la propia app Electron
    # (file:// o http://localhost en dev). No es un servicio público.
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=False,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Autenticación por token (si SFA_AUTH_TOKEN está seteado)
    app.add_middleware(AuthMiddleware)

    # Routers
    app.include_router(ping.router, prefix="/api", tags=["ping"])

    logger.info("FastAPI app creada correctamente.")
    return app


app = create_app()
