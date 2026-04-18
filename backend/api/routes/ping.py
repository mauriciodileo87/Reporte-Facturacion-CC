# backend/api/routes/ping.py
"""
Endpoint de healthcheck.

Electron lo usa para saber cuándo el backend terminó de arrancar y
está listo para recibir requests. Además verifica que los módulos
core/ se importen correctamente (en esta fase, solo app_state).
"""

from __future__ import annotations

from datetime import datetime

from fastapi import APIRouter
from pydantic import BaseModel

router = APIRouter()


class PingResponse(BaseModel):
    status: str
    version: str
    timestamp: str
    python_ok: bool
    app_state_ok: bool
    core_modules_checked: list[str]


@router.get("/ping", response_model=PingResponse)
def ping() -> PingResponse:
    """Healthcheck simple."""
    checked: list[str] = []
    app_state_ok = False

    # Intentamos importar los módulos core que ya deberían estar disponibles
    # en Fase 0 (solo app_state). En fases siguientes agregamos más.
    try:
        import app_state  # type: ignore  # noqa: F401
        app_state_ok = True
        checked.append("app_state")
    except Exception as exc:
        checked.append(f"app_state: FAIL ({exc})")

    return PingResponse(
        status="ok",
        version="0.1.0",
        timestamp=datetime.now().isoformat(timespec="seconds"),
        python_ok=True,
        app_state_ok=app_state_ok,
        core_modules_checked=checked,
    )
