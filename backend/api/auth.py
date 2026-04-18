# backend/api/auth.py
"""
Middleware de autenticación por token simple.

Comportamiento:
- Si la variable de entorno SFA_AUTH_TOKEN está seteada, exige que
  todos los requests traigan el header 'X-Auth-Token' con ese valor.
- Si NO está seteada (modo desarrollo), no exige nada y deja pasar todo.

En producción, Electron genera un token random al arrancar, lo setea
como env var del proceso Python, y lo usa en todos sus requests.
Así garantizamos que solo Electron pueda hablar con el backend aunque
el puerto esté abierto en localhost.
"""

from __future__ import annotations

import os

from fastapi import Request
from fastapi.responses import JSONResponse
from starlette.middleware.base import BaseHTTPMiddleware


# Rutas que siempre pasan sin exigir token (healthcheck y docs de dev).
RUTAS_ABIERTAS = {
    "/api/ping",
    "/docs",
    "/redoc",
    "/openapi.json",
}


class AuthMiddleware(BaseHTTPMiddleware):
    async def dispatch(self, request: Request, call_next):
        expected_token = os.environ.get("SFA_AUTH_TOKEN")

        # Modo dev: sin token configurado → dejar pasar todo
        if not expected_token:
            return await call_next(request)

        # Rutas abiertas siempre
        if request.url.path in RUTAS_ABIERTAS:
            return await call_next(request)

        received = request.headers.get("X-Auth-Token")
        if received != expected_token:
            return JSONResponse(
                status_code=401,
                content={"detail": "Token invalido o ausente."},
            )

        return await call_next(request)
