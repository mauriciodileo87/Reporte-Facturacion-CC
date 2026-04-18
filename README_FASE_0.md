# Fase 0 — Preparación del backend

Objetivo: dejar el backend Python corriendo como servidor HTTP local,
con tu código actual de `core/` intacto.

---

## 1. Crear la carpeta del proyecto nuevo

Elegí una ubicación en tu máquina (ejemplo: `C:\proyectos\sfa-facturacion\`)
y creá ahí la estructura inicial. En una terminal:

```bat
mkdir C:\proyectos\sfa-facturacion
cd C:\proyectos\sfa-facturacion
git init
```

---

## 2. Copiar los archivos que te entregué

Copiá los archivos que vienen en esta entrega respetando la jerarquía:

```
sfa-facturacion/
├── .gitignore
└── backend/
    ├── requirements.txt
    ├── run_dev.py
    ├── start_backend.bat
    ├── api/
    │   ├── __init__.py
    │   ├── server.py
    │   ├── auth.py
    │   ├── schemas.py
    │   └── routes/
    │       ├── __init__.py
    │       └── ping.py
    └── core/
        └── __init__.py
```

---

## 3. Copiar tu código actual a `backend/core/` (SIN MODIFICAR)

Esta es la parte crítica: **tu business logic NO se toca**.

Copiá tus archivos actuales a `backend/core/` tal cual están:

### Archivos top-level del proyecto viejo → `backend/core/`
- `app_state.py`
- `compare_sfa.py`
- `facuni_sfa.py`
- `json_sfa.py`
- `txt_sfa.py`
- `utils_sfa.py`
- `utils_excel.py`
- `tree_filters.py` *(por ahora queda, se elimina en fases siguientes)*
- `main_sfa.py` *(por ahora queda, se elimina en Fase 1)*
- Todas las `tab_*.py`

### Archivos de la carpeta `tabs/` vieja → `backend/core/` (planos, sin subcarpeta)
- `planilla_facturacion_storage.py`
- `planilla_area_recaudacion.py`
- `planilla_prescripciones.py`
- `planilla_agencia_amiga.py`
- `planilla_anticipos_topes.py`
- `planilla_control_cio.py`
- `planilla_totales.py`
- `planilla_clipboard.py`
- `scroll_utils.py`
- `filter_combobox_style.py`
- `tab_planilla_facturacion.py`

> **IMPORTANTE:** Cuando en fases siguientes los archivos que hoy importan
> `from tabs.planilla_clipboard import X` se migren, esos imports van a
> desaparecer porque en Electron el clipboard lo maneja el navegador.
> Por ahora no los toques, no molestan.

### Archivos JSON → `backend/data/` (creá esa carpeta)

```
backend/data/
├── planilla_facturacion_guardada.json
└── planilla_facturacion_bundle.json
```

---

## 4. Arrancar el backend

Doble click en `backend/start_backend.bat` (o desde terminal):

```bat
cd backend
start_backend.bat
```

La primera vez va a tardar porque:
1. Crea el venv en `backend/.venv/`
2. Instala FastAPI + Uvicorn + openpyxl + Pydantic
3. Arranca el servidor en `http://127.0.0.1:8765`

Cuando veas en consola algo como:

```
INFO:     Uvicorn running on http://127.0.0.1:8765
INFO:     Application startup complete.
```

está listo.

---

## 5. Probar que funciona

Abrí un navegador (o usá curl) y entrá a:

```
http://127.0.0.1:8765/api/ping
```

Deberías ver una respuesta JSON parecida a:

```json
{
  "status": "ok",
  "version": "0.1.0",
  "timestamp": "2026-04-17T11:00:00",
  "python_ok": true,
  "app_state_ok": true,
  "core_modules_checked": ["app_state"]
}
```

Si `app_state_ok` es `true`, **Fase 0 está completa**. Significa que:
- FastAPI arranca
- Tu `app_state.py` se carga sin errores
- El sistema de imports (con `core/` en sys.path) funciona

También podés abrir `http://127.0.0.1:8765/docs` para ver la UI interactiva
de Swagger con todos los endpoints (por ahora solo `/api/ping`).

---

## 6. Primer commit

```bat
cd ..
git add .
git commit -m "Fase 0: backend FastAPI + core/ intacto"
```

---

## Troubleshooting

### "py no se reconoce como comando"
Tenés Python instalado pero no el launcher `py`. Editá `start_backend.bat`
y cambiá `py -3.13` por `python` (o la ruta absoluta a tu Python).

### `app_state_ok: false` en el ping
Algún import de `app_state.py` falla. Abrí la consola donde corre el backend,
vas a ver el error. Lo más probable es que falte alguno de los otros archivos
core (aunque `app_state.py` no debería depender de ninguno).

### Puerto 8765 ocupado
Arrancá con otro puerto: `python run_dev.py 9000` (o editá el .bat).

### Error `ModuleNotFoundError: No module named 'api'`
Estás corriendo `run_dev.py` desde una carpeta incorrecta. Asegurate de
estar dentro de `backend/` cuando ejecutás.

---

## Siguiente paso: Fase 1

Cuando tengas Fase 0 funcionando, me avisás y arrancamos con:
- Setup Electron + React + Vite + TypeScript + Tailwind
- Proceso main de Electron que lanza el backend Python al abrir la ventana
- Primer tab funcional end-to-end: **Diferencias** (la más chica y autocontenida)
