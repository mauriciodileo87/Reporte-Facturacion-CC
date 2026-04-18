# tabs/planilla_area_recaudacion.py
from __future__ import annotations

import copy
import json
import os
from dataclasses import dataclass, field
from typing import Any
from datetime import datetime, date, timedelta

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.font as tkfont
import re

import app_state
from tabs.scroll_utils import bind_smooth_mousewheel
from tabs.planilla_clipboard import (
    bind_active_cell_tracking,
    create_undo_state,
    get_anchor_cell,
    get_clipboard_matrix,
    ordered_selected_rows,
    pop_undo_snapshot,
    push_undo_rows,
    restore_undo_snapshot,
    set_clipboard_matrix,
)


JUEGOS = [
    ("Quiniela", 80),
    ("Quiniela Ya", 79),
    ("Poceada", 82),
    ("Tombolina", 74),
    ("Quini 6", 69),
    ("Brinco", 13),
    ("Loto", 9),
    ("Loto 5", 5),
    ("LT", 41),
]


@dataclass
class PlanillaWidgets:
    juego: str
    codigo_juego: int
    header_canvas: tk.Canvas
    header2_canvas: tk.Canvas
    filter_canvas: tk.Canvas
    data_tree: ttk.Treeview
    cols: list[str]
    semanas: dict[int, list[int]]
    rangos_semana: dict[int, tuple[str, str]] = field(default_factory=dict)
    diff_labels: list[tk.Label] = field(default_factory=list)
    filter_vars: dict[str, tk.StringVar] = field(default_factory=dict)
    all_item_ids: list[str] = field(default_factory=list)


# ======================
# UTILIDADES FORMATO
# ======================

def fmt_pesos(x):
    try:
        x = float(x)
    except Exception:
        x = 0.0
    s = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "$ " + s


def parse_pesos(valor):
    if valor is None:
        return None
    s = str(valor).strip()
    if not s:
        return None
    s = re.sub(r"[^0-9,\.\-]", "", s)
    if not s:
        return None
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _darken(hex_color: str, factor: float = 0.82) -> str:
    """Devuelve un color más oscuro (mismo tono) para separadores."""
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = max(0, min(255, int(r * factor)))
    g = max(0, min(255, int(g * factor)))
    b = max(0, min(255, int(b * factor)))
    return f"#{r:02X}{g:02X}{b:02X}"


def _cell_has_value(v: str) -> bool:
    return bool(str(v).strip())


def _texto_filtro_normalizado(valor) -> str:
    return str(valor or "").strip().lower()


def _fila_planilla_coincide_filtros(w: PlanillaWidgets, vals: list[str]) -> bool:
    filtros = getattr(w, "filter_vars", {}) or {}
    if not filtros:
        return True

    for idx, col in enumerate(w.cols):
        filtro = _texto_filtro_normalizado(filtros.get(col).get() if col in filtros else "")
        if not filtro:
            continue
        valor = _texto_filtro_normalizado(vals[idx] if idx < len(vals) else "")
        if filtro not in valor:
            return False
    return True


def _restaurar_items_planilla(w: PlanillaWidgets):
    for idx, iid in enumerate(getattr(w, "all_item_ids", []) or []):
        try:
            w.data_tree.reattach(iid, "", idx)
        except Exception:
            pass


def _aplicar_filtros_planilla(w: PlanillaWidgets):
    _restaurar_items_planilla(w)

    for iid in list(w.data_tree.get_children()):
        vals = [str(v) for v in w.data_tree.item(iid, "values")]
        sorteo_txt = (vals[0] if vals else "").strip().lower()

        if sorteo_txt == "totales":
            continue
        if not any(str(v).strip() for v in vals):
            continue
        try:
            int((vals[0] if vals else "").strip())
        except Exception:
            continue

        if not _fila_planilla_coincide_filtros(w, vals):
            w.data_tree.detach(iid)

    _actualizar_fila_totales_planilla(w)
    _aplicar_zebra_planilla(w)
    _clear_diff_labels(w)
    _render_diff_labels(w)


def _semana_visible(valor: str) -> str:
    helper = getattr(app_state, "semana_visible_desde_valor", None)
    if callable(helper):
        try:
            return str(helper(valor) or valor)
        except Exception:
            pass
    return str(valor or "")

def _semana_interna(valor: str) -> str:
    txt = str(valor or "").strip()
    if not txt:
        return ""

    helper = getattr(app_state, "semana_interna_desde_visible", None)
    if callable(helper):
        try:
            normalizada = str(helper(valor) or "").strip()
            if normalizada:
                return normalizada
        except Exception:
            pass

    m = re.fullmatch(r"(?i)\s*semana\s*(\d+)\s*", txt)
    if m:
        try:
            return f"Semana {max(1, min(5, int(m.group(1))))}"
        except Exception:
            pass
    return txt

def _combo_values_semanas_desde_numeros(numeros: list[int]) -> list[str]:
    out = []
    for n in numeros:
        try:
            out.append(_semana_visible(f"Semana {int(n)}"))
        except Exception:
            out.append(_semana_visible("Semana 1"))
    return out

def _notificar_recalculo_totales_planilla():
    """
    Fuerza el recálculo inmediato de la sección Totales de Planilla Facturación.
    """
    refresh_totales = getattr(app_state, "_refresh_totales", None)
    if callable(refresh_totales):
        try:
            refresh_totales()
        except Exception:
            pass


# ======================
# PJU: JSON embebido en TXT
# ======================

def leer_json_desde_txt(path: str) -> Any:
    with open(path, "r", encoding="utf-8", errors="replace") as f:
        raw = f.read()
    i_obj = raw.find("{")
    i_arr = raw.find("[")
    if i_obj == -1 and i_arr == -1:
        raise ValueError("No encontré '{' ni '[' en el TXT para detectar JSON.")
    start = min([x for x in (i_obj, i_arr) if x != -1])
    candidate = raw[start:].strip()
    try:
        return json.loads(candidate)
    except Exception:
        last_obj = candidate.rfind("}")
        last_arr = candidate.rfind("]")
        end = max(last_obj, last_arr)
        if end == -1:
            raise ValueError("Encontré inicio de JSON pero no encontré cierre '}' o ']'.")
        return json.loads(candidate[: end + 1].strip())


def extraer_sorteos_por_codigo(obj: Any, codigo_juego_objetivo: int) -> list[int]:
    encontrados: set[int] = set()

    def try_add_from_dict(d: dict):
        pg = d.get("parametros_genericos")
        if isinstance(pg, dict):
            try:
                cj_int = int(pg.get("codigo_juego"))
            except Exception:
                cj_int = None
            if cj_int == codigo_juego_objetivo:
                try:
                    ns_int = int(pg.get("numero_sorteo"))
                    if 1 <= ns_int <= 999999:
                        encontrados.add(ns_int)
                except Exception:
                    pass

    def walk(x: Any):
        if isinstance(x, dict):
            try_add_from_dict(x)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for it in x:
                walk(it)

    walk(obj)
    return sorted(encontrados)


def _parse_fecha_sorteo(valor: Any) -> date | None:
    if valor is None:
        return None
    txt = str(valor).strip()
    if not txt:
        return None

    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
        try:
            return datetime.strptime(txt, fmt).date()
        except Exception:
            pass

    txt_base = txt.split("T", 1)[0].split(" ", 1)[0].strip()
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(txt_base, fmt).date()
        except Exception:
            pass

    # Compatibilidad con ISO 8601 con milisegundos/zona horaria
    # (ej.: 2026-03-21T21:00:00.000-03:00 o con sufijo Z).
    try:
        iso = txt.replace("Z", "+00:00")
        return datetime.fromisoformat(iso).date()
    except Exception:
        pass
    return None


def extraer_sorteos_por_semanas(obj: Any, codigo_juego_objetivo: int) -> tuple[dict[int, list[int]], int | None, date | None]:
    sorteos_por_fecha: list[tuple[int, date]] = []
    sorteos_por_semana_explicita: dict[int, set[int]] = {}

    def _parse_semana(valor: Any) -> int | None:
        if valor is None:
            return None

        # Fast-path para enteros ya válidos.
        if isinstance(valor, int) and valor > 0:
            return valor

        texto = str(valor or "").strip()
        if not texto:
            return None

        # Casos típicos: "4", "4.0", "Semana 4", "4,5", etc.
        m = re.search(r"\d+", texto)
        if not m:
            return None
        try:
            sem = int(m.group(0))
        except Exception:
            return None
        return sem if sem > 0 else None

    def _extraer_semana_explicita(pg: dict, origen: dict) -> int | None:
        candidatos = (
            "semana",
            "numero_semana",
            "nro_semana",
            "semana_sorteo",
            "week",
        )
        for key in candidatos:
            for fuente in (pg, origen):
                if not isinstance(fuente, dict) or key not in fuente:
                    continue
                sem = _parse_semana(fuente.get(key, ""))
                if sem is not None and sem > 0:
                    return sem
        return None

    def try_add_from_dict(d: dict):
        pg = d.get("parametros_genericos")
        if not isinstance(pg, dict):
            return
        try:
            cj_int = int(pg.get("codigo_juego"))
        except Exception:
            cj_int = None
        if cj_int != codigo_juego_objetivo:
            return

        try:
            ns_int = int(pg.get("numero_sorteo"))
        except Exception:
            return
        if not (1 <= ns_int <= 999999):
            return

        semana_explicita = _extraer_semana_explicita(pg, d)
        if semana_explicita is not None:
            sorteos_por_semana_explicita.setdefault(semana_explicita, set()).add(ns_int)

        fecha = _parse_fecha_sorteo(pg.get("fecha_sorteo"))
        if fecha is not None:
            sorteos_por_fecha.append((ns_int, fecha))

    def walk(x: Any):
        if isinstance(x, dict):
            try_add_from_dict(x)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for it in x:
                walk(it)

    walk(obj)

    # Prioridad: si el archivo trae número de semana explícito,
    # usarlo para no colapsar todos los sorteos en Semana 1.
    if sorteos_por_semana_explicita:
        semanas_ordenadas = {
            semana: sorted(vals)
            for semana, vals in sorted(sorteos_por_semana_explicita.items())
            if vals
        }
        return semanas_ordenadas, (max(semanas_ordenadas) if semanas_ordenadas else None), None

    if not sorteos_por_fecha:
        return {}, None, None

    fecha_base = min(f for _, f in sorteos_por_fecha)
    lunes_semana_1 = fecha_base - timedelta(days=fecha_base.weekday())

    semanas: dict[int, set[int]] = {}
    for sorteo, fecha in sorteos_por_fecha:
        semana = ((fecha - lunes_semana_1).days // 7) + 1
        semanas.setdefault(semana, set()).add(sorteo)

    semanas_ordenadas = {semana: sorted(vals) for semana, vals in sorted(semanas.items())}
    return semanas_ordenadas, (max(semanas_ordenadas) if semanas_ordenadas else None), lunes_semana_1


# ======================
# GUARDAR / CARGAR (solo area recaudación)
# ======================

def _ruta_guardado_planilla() -> str:
    appdata = os.environ.get("APPDATA")
    if appdata:
        base = os.path.join(appdata, "ReporteFacturacion")
    else:
        base = os.path.join(os.path.expanduser("~"), ".reporte_facturacion")

    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "planilla_facturacion_guardada.json")


def filas_con_datos(w: PlanillaWidgets) -> list[list[str]]:
    rows = []
    for iid in w.data_tree.get_children():
        vals = [str(v) for v in w.data_tree.item(iid, "values")]
        if not any(str(v).strip() for v in vals):
            continue
        if str(vals[0]).strip().lower() == "totales":
            continue
        rows.append(vals)
    return rows


def _sorteos_visibles_de_la_grilla(w: PlanillaWidgets) -> list[int]:
    """
    Devuelve los sorteos visibles actualmente en la grilla del juego activo.
    Se usa para persistir correctamente la semana donde el usuario agregó
    sorteos manuales, incluso aunque todavía no hayan venido desde PJU.
    """
    out: list[int] = []
    vistos: set[int] = set()
    for vals in filas_con_datos(w):
        if not vals:
            continue
        try:
            sorteo = int(str(vals[0]).strip())
        except Exception:
            continue
        if sorteo in vistos:
            continue
        vistos.add(sorteo)
        out.append(sorteo)
    return sorted(out)


def _inyectar_sorteos_visibles_en_semana_actual(
    semanas_raw: dict | None,
    w: PlanillaWidgets,
) -> dict[int, list[int]]:
    """
    Regla de negocio (ROBUSTEZ):
    - Solo incorporamos al mapa semanal los sorteos *nuevos* creados manualmente.
    - Nunca "movemos" sorteos que ya estaban asignados a otra semana, aunque por un
      bug visual se hayan mostrado en la grilla actual.

    Esto evita que Semana 1 "absorba" sorteos de Semana 2/3/4/5 y que luego el
    filtro colapse mostrando únicamente Semana 1.
    """
    semanas_norm = _normalizar_mapa_semanas(semanas_raw or {})
    sorteos_visibles = _sorteos_visibles_de_la_grilla(w)
    if not sorteos_visibles:
        return semanas_norm

    asignados: set[int] = set()
    for vals in (semanas_norm or {}).values():
        if not isinstance(vals, list):
            continue
        for v in vals:
            try:
                asignados.add(int(v))
            except Exception:
                pass

    nuevos = []
    for s in sorteos_visibles:
        try:
            n = int(s)
        except Exception:
            continue
        if n not in asignados:
            nuevos.append(n)

    if not nuevos:
        return semanas_norm

    try:
        semana_actual = int(getattr(w, "semana_actual", 0) or 0)
    except Exception:
        semana_actual = 0

    if semana_actual <= 0:
        semanas_existentes = sorted(int(k) for k in semanas_norm.keys()) if semanas_norm else []
        if len(semanas_existentes) == 1:
            semana_actual = semanas_existentes[0]
        else:
            semana_actual = 1

    bucket = [int(s) for s in (semanas_norm.get(semana_actual, []) or []) if str(s).strip()]
    bucket_set = {int(v) for v in bucket if str(v).strip()}
    for n in nuevos:
        if n not in bucket_set:
            bucket_set.add(n)

    semanas_norm[semana_actual] = sorted(bucket_set)
    return dict(sorted(semanas_norm.items()))


def _mapear_filas_por_sorteo(filas: list[list[str]] | None, cols_len: int) -> tuple[dict[int, list[str]], list[int]]:
    filas_por_sorteo: dict[int, list[str]] = {}
    orden: list[int] = []

    for fila in filas or []:
        if not isinstance(fila, list) or not fila:
            continue
        try:
            sorteo = int(str(fila[0]).strip())
        except Exception:
            continue

        row = [str(v) for v in fila]
        if len(row) < cols_len:
            row.extend([""] * (cols_len - len(row)))
        row = row[:cols_len]
        filas_por_sorteo[sorteo] = row
        if sorteo not in orden:
            orden.append(sorteo)

    return filas_por_sorteo, orden


def _mergear_filas_guardadas_con_visibles(
    filas_guardadas: list[list[str]] | None,
    filas_visibles: list[list[str]] | None,
    cols_len: int,
) -> list[list[str]]:
    guardadas_map, guardadas_orden = _mapear_filas_por_sorteo(filas_guardadas, cols_len)
    visibles_map, visibles_orden = _mapear_filas_por_sorteo(filas_visibles, cols_len)

    for sorteo, fila in visibles_map.items():
        guardadas_map[sorteo] = fila
        if sorteo not in guardadas_orden:
            guardadas_orden.append(sorteo)

    orden_final = sorted({*guardadas_orden, *visibles_orden})
    return [guardadas_map[sorteo] for sorteo in orden_final if sorteo in guardadas_map]


def _leer_json(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _aplicar_filas_a_widget(w: PlanillaWidgets, filas: list, clear_rows: bool = True):
    total_iid = _ensure_total_row_iid(w)
    ids = [iid for iid in w.data_tree.get_children() if iid != total_iid]

    if clear_rows:
        for iid in ids:
            w.data_tree.item(iid, values=[""] * len(w.cols))

    for idx, vals in enumerate(filas or []):
        if idx >= len(ids):
            break
        row_vals = [str(v) for v in (vals or [])]
        if len(row_vals) < len(w.cols):
            row_vals.extend([""] * (len(w.cols) - len(row_vals)))
        w.data_tree.item(ids[idx], values=row_vals[: len(w.cols)])

    try:
        w.data_tree.item(total_iid, values=["Totales"] + [""] * (len(w.cols) - 1))
    except Exception:
        pass


def _snapshot_juego(w: PlanillaWidgets) -> dict:
    semanas_out: dict[str, list[int]] = {}
    semanas_sync = _inyectar_sorteos_visibles_en_semana_actual(getattr(w, "semanas", {}) or {}, w)
    for sem, sorteos in (semanas_sync or {}).items():
        try:
            sem_n = int(sem)
        except Exception:
            continue
        semanas_out[str(sem_n)] = sorted({int(s) for s in (sorteos or []) if str(s).strip()})

    rangos_raw = getattr(w, "rangos_semana", {})
    rangos_semana: dict[str, dict[str, str]] = {}
    for sem, rango in (rangos_raw.items() if isinstance(rangos_raw, dict) else []):
        try:
            sem_n = int(sem)
        except Exception:
            continue
        if not isinstance(rango, tuple) or len(rango) != 2:
            continue
        desde = str(rango[0] or "").strip()
        hasta = str(rango[1] or "").strip()
        if not desde or not hasta:
            continue
        rangos_semana[str(sem_n)] = {"desde": desde, "hasta": hasta}

    data_guardada = _leer_planilla_juego_desde_storage(w.juego)
    filas_guardadas = data_guardada.get("filas", []) if isinstance(data_guardada, dict) else []

    return {
        "codigo_juego": w.codigo_juego,
        "columnas": w.cols,
        "filas": _mergear_filas_guardadas_con_visibles(filas_guardadas, filas_con_datos(w), len(w.cols)),
        "semanas": semanas_out,
        "rangos_semana": rangos_semana,
    }


def _filtrar_filas_por_sorteos(filas: list[list[str]], sorteos_permitidos: set[int]) -> list[list[str]]:
    filtradas: list[list[str]] = []
    for vals in filas or []:
        if not vals:
            continue
        try:
            sorteo = int(str(vals[0]).strip())
        except Exception:
            continue
        if sorteo in sorteos_permitidos:
            filtradas.append([str(v) for v in vals])
    return filtradas


def _crear_template_semana_1(filas: list[list[str]], cols_len: int) -> list[list[str]]:
    template: list[list[str]] = []
    for vals in filas or []:
        row = [str(v) for v in (vals or [])]
        if len(row) < cols_len:
            row.extend([""] * (cols_len - len(row)))
        row = row[:cols_len]
        if row:
            row[0] = ""
        template.append(row)
    return template


def _leer_planilla_juego_desde_storage(juego: str) -> dict:
    payload = _leer_json(_ruta_guardado_planilla())
    planillas = payload.get("planillas", {}) if isinstance(payload, dict) else {}
    data_juego = planillas.get(juego, {}) if isinstance(planillas, dict) else {}
    return data_juego if isinstance(data_juego, dict) else {}


def _extraer_semanas_guardadas(data_juego: dict) -> dict[int, list[int]]:
    raw = data_juego.get("semanas", {}) if isinstance(data_juego, dict) else {}
    return _normalizar_mapa_semanas(raw)


def _normalizar_rangos_semana(raw: dict | None) -> dict[int, tuple[str, str]]:
    rangos: dict[int, tuple[str, str]] = {}
    if not isinstance(raw, dict):
        return rangos

    for k, v in raw.items():
        try:
            sem = int(str(k).strip().lower().replace("semana", "").strip())
        except Exception:
            continue

        if isinstance(v, dict):
            desde = str(v.get("desde", "") or "").strip()
            hasta = str(v.get("hasta", "") or "").strip()
        elif isinstance(v, (list, tuple)) and len(v) >= 2:
            desde = str(v[0] or "").strip()
            hasta = str(v[1] or "").strip()
        else:
            continue

        if desde and hasta:
            rangos[sem] = (desde, hasta)

    return dict(sorted(rangos.items()))


def _extraer_rangos_semana_guardados(data_juego: dict) -> dict[int, tuple[str, str]]:
    raw = data_juego.get("rangos_semana", {}) if isinstance(data_juego, dict) else {}
    return _normalizar_rangos_semana(raw)


def _calcular_rangos_semana_desde_lunes(
    semanas: dict[int, list[int]] | None,
    lunes_semana_1: date | None,
) -> dict[int, tuple[str, str]]:
    if not semanas:
        return {}

    # Si el PJU no trae fecha_sorteo, igual necesitamos setear Del/Al
    # automáticamente por semana (lunes a domingo).
    if lunes_semana_1 is None:
        hoy = date.today()
        lunes_semana_1 = hoy - timedelta(days=hoy.weekday())

    out: dict[int, tuple[str, str]] = {}
    for sem in sorted(semanas.keys()):
        try:
            sem_n = int(sem)
        except Exception:
            continue
        if sem_n <= 0:
            continue

        desde = lunes_semana_1 + timedelta(days=(sem_n - 1) * 7)
        hasta = desde + timedelta(days=6)
        out[sem_n] = (desde.strftime("%d/%m/%Y"), hasta.strftime("%d/%m/%Y"))

    return out


def _mergear_rangos_semana(
    existentes: dict[int, tuple[str, str]] | None,
    nuevos: dict[int, tuple[str, str]] | None,
) -> dict[int, tuple[str, str]]:
    merged = dict(existentes or {})
    for sem, rango in (nuevos or {}).items():
        if not isinstance(rango, tuple) or len(rango) != 2:
            continue
        desde = str(rango[0] or "").strip()
        hasta = str(rango[1] or "").strip()
        if desde and hasta:
            # El último PJU importado manda: si cambia el rango de una semana
            # se actualiza para mantener Del/Al consistente en todas las secciones.
            merged[sem] = (desde, hasta)
    return dict(sorted(merged.items()))


def _alinear_semanas_importadas_con_rangos_existentes(
    semanas_importadas: dict[int, list[int]] | None,
    rangos_importados: dict[int, tuple[str, str]] | None,
    rangos_existentes: dict[int, tuple[str, str]] | None,
) -> tuple[dict[int, list[int]], dict[int, tuple[str, str]]]:
    """
    Evita colapsar semanas cuando se importa un PJU "cortado" (ej.: trae solo
    la semana 2, pero internamente empieza a numerar desde 1 por fecha mínima).

    Estrategia:
    - Solo se alinea cuando el PJU importado trae UNA única semana (caso típico
      de importes parciales semanales).
    - Si el "desde" de la primera semana importada coincide exactamente con el
      "desde" de una semana ya existente, se remapea a esa base.
    - Si no hay coincidencia exacta, pero empieza justo después del último rango
      existente, se considera continuación y se desplaza al siguiente número.
    """
    semanas_in = dict(semanas_importadas or {})
    rangos_in = dict(rangos_importados or {})
    rangos_prev = dict(rangos_existentes or {})

    if not semanas_in:
        return {}, {}
    if not rangos_in or not rangos_prev:
        return semanas_in, rangos_in
    # Evitar normalizaciones agresivas sobre PJU completos (2+ semanas), que
    # deben conservar su propia segmentación semanal.
    if len(semanas_in) != 1:
        return semanas_in, rangos_in

    try:
        sem_inicial = min(int(k) for k in semanas_in.keys())
    except Exception:
        return semanas_in, rangos_in

    rango_inicial = rangos_in.get(sem_inicial)
    if not (isinstance(rango_inicial, tuple) and len(rango_inicial) == 2):
        return semanas_in, rangos_in

    fecha_desde_inicial = _parse_fecha_sorteo(rango_inicial[0])
    if fecha_desde_inicial is None:
        return semanas_in, rangos_in

    sem_destino: int | None = None
    for sem_prev, rango_prev in sorted(rangos_prev.items()):
        if not (isinstance(rango_prev, tuple) and len(rango_prev) == 2):
            continue
        desde_prev = _parse_fecha_sorteo(rango_prev[0])
        hasta_prev = _parse_fecha_sorteo(rango_prev[1])
        if desde_prev is None or hasta_prev is None:
            continue

        if fecha_desde_inicial == desde_prev:
            sem_destino = int(sem_prev)
            break

    if sem_destino is None:
        sems_prev = sorted(int(k) for k in rangos_prev.keys() if int(k) > 0)
        if sems_prev:
            sem_ultima = sems_prev[-1]
            rango_ultima = rangos_prev.get(sem_ultima)
            if isinstance(rango_ultima, tuple) and len(rango_ultima) == 2:
                hasta_ultima = _parse_fecha_sorteo(rango_ultima[1])
                if hasta_ultima is not None and fecha_desde_inicial == (hasta_ultima + timedelta(days=1)):
                    sem_destino = sem_ultima + 1

    if sem_destino is None:
        return semanas_in, rangos_in

    offset = int(sem_destino) - int(sem_inicial)
    if offset == 0:
        return semanas_in, rangos_in

    semanas_out: dict[int, list[int]] = {}
    for sem, sorteos in semanas_in.items():
        try:
            sem_n = int(sem) + offset
        except Exception:
            continue
        if sem_n < 1:
            continue
        semanas_out[sem_n] = list(sorteos or [])

    rangos_out: dict[int, tuple[str, str]] = {}
    for sem, rango in rangos_in.items():
        try:
            sem_n = int(sem) + offset
        except Exception:
            continue
        if sem_n < 1:
            continue
        if isinstance(rango, tuple) and len(rango) == 2:
            rangos_out[sem_n] = (str(rango[0] or "").strip(), str(rango[1] or "").strip())

    if not semanas_out:
        return semanas_in, rangos_in
    return dict(sorted(semanas_out.items())), dict(sorted(rangos_out.items()))


def _normalizar_mapa_semanas(raw: dict | None) -> dict[int, list[int]]:
    semanas: dict[int, set[int]] = {}
    if not isinstance(raw, dict):
        return {}

    for k, vals in raw.items():
        try:
            sem = int(str(k).strip().lower().replace("semana", "").strip())
        except Exception:
            continue

        if not isinstance(vals, list):
            continue

        bucket = semanas.setdefault(sem, set())
        for v in vals:
            try:
                bucket.add(int(v))
            except Exception:
                continue

    return {sem: sorted(sorteos) for sem, sorteos in sorted(semanas.items())}


def _ensure_total_row_iid(w: PlanillaWidgets) -> str:
    total_iid = getattr(w, "total_iid", None)
    try:
        children = list(w.data_tree.get_children())
    except Exception:
        children = []

    if total_iid and total_iid in children:
        return str(total_iid)

    for iid in children:
        try:
            vals = [str(v) for v in w.data_tree.item(iid, "values")]
        except Exception:
            continue
        if vals and str(vals[0]).strip().lower() == "totales":
            w.total_iid = iid
            if iid not in getattr(w, "all_item_ids", []):
                w.all_item_ids = list(getattr(w, "all_item_ids", []) or []) + [iid]
            return str(iid)

    total_iid = w.data_tree.insert("", "end", values=[""] * len(w.cols))
    try:
        w.data_tree.item(total_iid, values=["Totales"] + [""] * (len(w.cols) - 1))
    except Exception:
        pass
    w.total_iid = total_iid
    if total_iid not in getattr(w, "all_item_ids", []):
        w.all_item_ids = list(getattr(w, "all_item_ids", []) or []) + [total_iid]
    return str(total_iid)

def _aplicar_template_semana_1_a_sorteos(
    w: PlanillaWidgets,
    filas_template: list[list[str]],
    sorteos_semana_1: list[int],
    estado_var=None,
):
    _aplicar_filas_a_widget(w, [], clear_rows=True)

    ids = w.data_tree.get_children()
    cantidad = min(len(ids), len(filas_template), len(sorteos_semana_1))
    for idx in range(cantidad):
        row = [str(v) for v in (filas_template[idx] or [])]
        if len(row) < len(w.cols):
            row.extend([""] * (len(w.cols) - len(row)))
        row = row[: len(w.cols)]
        row[0] = str(int(sorteos_semana_1[idx]))
        w.data_tree.item(ids[idx], values=row)

    refresh_planilla(w, estado_var)




def _normalizar_seccion_semanal_para_semana_nueva(payload: dict, columnas_datos: tuple[str, ...]) -> dict:
    if not isinstance(payload, dict):
        return {}

    semanas_in = payload.get("semanas", {})
    if not isinstance(semanas_in, dict):
        return copy.deepcopy(payload)

    semanas_orden = ["Semana 1", "Semana 2", "Semana 3", "Semana 4", "Semana 5"]

    def _semana_tiene_datos(rows: list[dict]) -> bool:
        if not isinstance(rows, list):
            return False
        for row in rows:
            if not isinstance(row, dict):
                continue
            for col in columnas_datos:
                if str(row.get(col, "") or "").strip():
                    return True
        return False

    ultima = ""
    for sem in reversed(semanas_orden):
        if _semana_tiene_datos(semanas_in.get(sem, [])):
            ultima = sem
            break

    if not ultima:
        actual = str(payload.get("current_semana", "") or "")
        if actual in semanas_orden:
            ultima = actual

    if not ultima:
        return copy.deepcopy(payload)

    semana1_rows = copy.deepcopy(semanas_in.get(ultima, [])) if isinstance(semanas_in.get(ultima, []), list) else []
    semanas_out: dict[str, list[dict]] = {"Semana 1": semana1_rows}
    for sem in semanas_orden[1:]:
        semanas_out[sem] = []

    return {
        "version": int(payload.get("version", 1) or 1),
        "current_semana": "Semana 1",
        "semanas": semanas_out,
    }


def _normalizar_agencia_amiga_para_semana_nueva(payload: dict) -> dict:
    if not isinstance(payload, dict):
        return {"juegos": {}}

    juegos_in = payload.get("juegos", {})
    if not isinstance(juegos_in, dict):
        return copy.deepcopy(payload)

    juegos_out: dict[str, dict[str, dict]] = {}
    semanas_orden = ["1", "2", "3", "4", "5"]

    for juego, semanas_raw in juegos_in.items():
        if not isinstance(semanas_raw, dict):
            continue

        ultima_semana = ""
        for sem in reversed(semanas_orden):
            bucket = semanas_raw.get(sem, {})
            if isinstance(bucket, dict) and bucket:
                ultima_semana = sem
                break

        if not ultima_semana:
            continue

        juegos_out[juego] = {"1": copy.deepcopy(semanas_raw.get(ultima_semana, {}))}

    out = {"juegos": juegos_out}
    ui_in = payload.get("_ui", {})
    if isinstance(ui_in, dict):
        ui_out = copy.deepcopy(ui_in)
        ui_out["semana"] = 1
        out["_ui"] = ui_out
    return out


def exportar_ultima_semana_a_planilla_nueva(
    widgets_por_juego: dict[str, PlanillaWidgets],
    semanas_por_juego: dict[str, dict[int, list[int]]],
    estado_var=None,
):
    if not widgets_por_juego:
        messagebox.showwarning("Planilla nueva", "No hay juegos construidos para exportar.")
        return

    planillas_exportadas: dict[str, dict] = {}
    juegos_sin_semana: list[str] = []

    for juego, w in widgets_por_juego.items():
        semanas = semanas_por_juego.get(juego, {}) or {}
        numeros_semana = sorted(int(n) for n in semanas.keys())
        if not numeros_semana:
            juegos_sin_semana.append(juego)
            continue

        ultima_semana = numeros_semana[-1]
        sorteos_ultima_semana = set(int(s) for s in semanas.get(ultima_semana, []) if str(s).strip())
        if not sorteos_ultima_semana:
            juegos_sin_semana.append(juego)
            continue

        filas_ultima_semana = _filtrar_filas_por_sorteos(filas_con_datos(w), sorteos_ultima_semana)
        if not filas_ultima_semana:
            juegos_sin_semana.append(juego)
            continue

        planillas_exportadas[juego] = {
            "codigo_juego": w.codigo_juego,
            "columnas": w.cols,
            "filas": filas_ultima_semana,
            "semanas": {"1": sorted(sorteos_ultima_semana)},
            "modo_semana_nueva": True,
            "semana_origen": ultima_semana,
            "semana_exportada": 1,
            "filas_template_semana_1": _crear_template_semana_1(filas_ultima_semana, len(w.cols)),
        }

    if not planillas_exportadas:
        detalle = "\n".join(f"- {j}" for j in juegos_sin_semana) if juegos_sin_semana else ""
        msg = "No hay datos de última semana para exportar."
        if detalle:
            msg += f"\n\nJuegos omitidos:\n{detalle}"
        messagebox.showwarning("Planilla nueva", msg)
        return

    fecha_nombre = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = filedialog.asksaveasfilename(
        title="Guardar planilla nueva (última semana como Semana 1)",
        defaultextension=".json",
        initialfile=f"planilla_facturacion_semana_nueva_{fecha_nombre}.json",
        filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
    )
    if not path:
        return

    presc_semanas_src = getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", {}) or {}
    presc_reporte_src = getattr(app_state, "reporte_prescripciones_por_juego", {}) or {}
    presc_tickets_src = getattr(app_state, "tickets_prescripciones_por_juego", {}) or {}
    presc_sfa_src = getattr(app_state, "sfa_prescripciones_por_juego", {}) or {}

    presc_sorteos_sem1: dict[str, dict[int, list[int]]] = {}
    presc_reporte_out: dict[str, dict[str, float]] = {}
    presc_tickets_out: dict[str, dict[str, float]] = {}
    presc_sfa_out: dict[str, dict[str, float]] = {}

    for juego, semanas in (presc_semanas_src.items() if isinstance(presc_semanas_src, dict) else []):
        if not isinstance(semanas, dict):
            continue
        semanas_num = sorted(int(k) for k in semanas.keys() if str(k).strip().isdigit())
        if not semanas_num:
            continue
        ultima = semanas_num[-1]
        sorteos = []
        for s in semanas.get(ultima, []) if isinstance(semanas.get(ultima, []), list) else []:
            try:
                sorteos.append(int(s))
            except Exception:
                continue
        if not sorteos:
            continue

        presc_sorteos_sem1.setdefault(juego, {})
        presc_sorteos_sem1[juego][1] = sorted(set(sorteos))
        permitidos = {str(int(s)) for s in sorteos}

        def _filtrar(src):
            out = {}
            for k, v in (src.get(juego, {}) if isinstance(src, dict) else {}).items():
                kk = str(k).strip().lstrip("0") or "0"
                if kk in permitidos:
                    try:
                        out[kk] = float(v or 0.0)
                    except Exception:
                        out[kk] = 0.0
            return out

        r = _filtrar(presc_reporte_src)
        t = _filtrar(presc_tickets_src)
        s = _filtrar(presc_sfa_src)
        if r:
            presc_reporte_out[juego] = r
        if t:
            presc_tickets_out[juego] = t
        if s:
            presc_sfa_out[juego] = s

    anticipos_payload = _normalizar_seccion_semanal_para_semana_nueva(
        copy.deepcopy(getattr(app_state, "planilla_anticipos_topes_data", {}) or {}),
        ("lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "reporte_prescripto", "sfa_prescripto"),
    )
    control_cio_payload = _normalizar_seccion_semanal_para_semana_nueva(
        copy.deepcopy(getattr(app_state, "planilla_control_cio_data", {}) or {}),
        ("lunes", "martes", "miercoles", "jueves", "viernes", "sabado", "reporte_prescripto", "sfa_prescripto"),
    )
    agencia_amiga_payload = _normalizar_agencia_amiga_para_semana_nueva(
        copy.deepcopy(getattr(app_state, "planilla_agencia_amiga_data", {}) or {})
    )

    payload = {
        "version": 1,
        "generado": datetime.now().isoformat(timespec="seconds"),
        "area_recaudacion": {
            "version": 2,
            "generado": datetime.now().isoformat(timespec="seconds"),
            "planillas": planillas_exportadas,
        },
        "prescripciones": {
            "reporte_prescripciones_por_juego": presc_reporte_out,
            "tickets_prescripciones_por_juego": presc_tickets_out,
            "sfa_prescripciones_por_juego": presc_sfa_out,
            "prescripciones_sorteos_por_semana_por_juego": presc_sorteos_sem1,
        },
        "anticipos_topes": anticipos_payload,
        "control_cio": control_cio_payload,
        "agencia_amiga": agencia_amiga_payload,
    }

    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showerror("Planilla nueva", f"No pude guardar el archivo.\n{e}")
        return

    if estado_var is not None:
        estado_var.set(f"Planilla nueva guardada en: {path}")

    if juegos_sin_semana:
        detalle = "\n".join(f"- {j}" for j in juegos_sin_semana)
        messagebox.showinfo(
            "Exportación completada",
            "Se guardó SOLO la última semana como Semana 1.\n\n"
            f"Juegos exportados: {', '.join(planillas_exportadas.keys())}\n\n"
            f"Juegos omitidos:\n{detalle}",
        )
    else:
        messagebox.showinfo(
            "Exportación completada",
            "Se guardó SOLO la última semana como Semana 1.",
        )


def _recargar_juego_desde_storage(juego: str, w: PlanillaWidgets, estado_var=None):
    payload = _leer_json(_ruta_guardado_planilla())
    planillas = payload.get("planillas", {}) if isinstance(payload, dict) else {}
    data_juego = planillas.get(juego, {}) if isinstance(planillas, dict) else {}
    filas = data_juego.get("filas", []) if isinstance(data_juego, dict) else []

    # Blindaje: si el storage viene sin "semanas"/"rangos", NO pisar lo que ya
    # está en memoria, porque eso colapsa el filtro a "Semana 1".
    semanas_new = _extraer_semanas_guardadas(data_juego)
    if semanas_new:
        w.semanas = semanas_new

    rangos_new = _extraer_rangos_semana_guardados(data_juego)
    if rangos_new:
        w.rangos_semana = rangos_new

    try:
        sem_act = int(data_juego.get("semana_actual", 0) or 0)
    except Exception:
        sem_act = 0
    if sem_act > 0:
        w.semana_actual = sem_act

    # Fuerza reaplicar el filtro semanal luego de una recarga desde storage.
    # Si dejamos el cache anterior, puede "saltarse" la carga y quedar
    # visible una mezcla de sorteos de varias semanas hasta cambiar manualmente.
    w._ultimo_juego_cargado = ""
    w._ultima_semana_cargada = 0
    w._ultimos_sorteos_cargados = []
    _aplicar_filas_a_widget(w, filas, clear_rows=True)
    refresh_planilla(w, estado_var)


def guardar_area_recaudacion(widgets_por_juego: dict[str, PlanillaWidgets], estado_var=None):
    # Si hay una celda en edición (Entry superpuesto), primero la confirmamos
    # para evitar perder el último valor tipeado al cambiar de semana/juego.
    for w in (widgets_por_juego or {}).values():
        cerrar_editor = getattr(w, "close_editor", None)
        if callable(cerrar_editor):
            try:
                cerrar_editor(save=True)
            except Exception:
                pass

    payload = {"version": 2, "generado": datetime.now().isoformat(timespec="seconds"), "planillas": {}}
    path = _ruta_guardado_planilla()

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                existing = json.load(f)
            if isinstance(existing, dict) and isinstance(existing.get("planillas"), dict):
                payload["planillas"].update(existing["planillas"])
        except Exception:
            pass

    for juego, w in widgets_por_juego.items():
        existente = payload["planillas"].get(juego, {}) if isinstance(payload["planillas"].get(juego, {}), dict) else {}
        data_juego = dict(existente)
        semanas_previas = _extraer_semanas_guardadas(existente)
        semanas_base = _mergear_semanas_sin_pisar(semanas_previas, getattr(w, "semanas", {}) or {})
        semanas_norm = _inyectar_sorteos_visibles_en_semana_actual(semanas_base, w)
        # Blindaje: nunca achicar / mover sorteos ya asignados.
        semanas_norm = _mergear_semanas_sin_pisar(semanas_previas, semanas_norm)
        # Persistimos también en el widget para que, al cambiar enseguida de
        # juego/sección/semana, la recarga use ya el mapa corregido.
        w.semanas = dict(semanas_norm)
        rangos_previos = _extraer_rangos_semana_guardados(existente)
        rangos_widget = _normalizar_rangos_semana(getattr(w, "rangos_semana", {}) or {})
        rangos_norm = dict(rangos_previos)
        rangos_norm.update(rangos_widget)
        w.rangos_semana = dict(rangos_norm)
        filas_previas = existente.get("filas", []) if isinstance(existente, dict) else []
        filas_merged = _mergear_filas_guardadas_con_visibles(filas_previas, filas_con_datos(w), len(w.cols))
        data_juego.update({
            "codigo_juego": w.codigo_juego,
            "columnas": w.cols,
            "filas": filas_merged,
            "semanas": {str(int(k)): sorted({int(s) for s in (v or []) if str(s).strip()}) for k, v in (semanas_norm or {}).items()},
            "rangos_semana": {str(int(k)): {"desde": v[0], "hasta": v[1]} for k, v in (rangos_norm or {}).items()},
            "semana_actual": (int(getattr(w, "semana_actual", 0) or 0) or int((existente or {}).get("semana_actual", 0) or 0)),
        })
        payload["planillas"][juego] = data_juego

    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    if estado_var is not None:
        estado_var.set("Planilla guardada. Se restaurará al abrir nuevamente.")


def cargar_area_recaudacion_guardada(widgets_por_juego: dict[str, PlanillaWidgets], estado_var=None):
    path = _ruta_guardado_planilla()
    if not os.path.exists(path):
        return
    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)
    except Exception:
        return

    planillas = payload.get("planillas", {}) if isinstance(payload, dict) else {}
    for juego, w in widgets_por_juego.items():
        data_juego = planillas.get(juego, {}) if isinstance(planillas, dict) else {}
        filas = data_juego.get("filas", []) if isinstance(data_juego, dict) else []
        w.semanas = _extraer_semanas_guardadas(data_juego)
        w.rangos_semana = _extraer_rangos_semana_guardados(data_juego)
        _aplicar_filas_a_widget(w, filas, clear_rows=True)
        refresh_planilla(w, estado_var)


# ======================
# LÓGICA DE CÁLCULO
# ======================

def get_tasa_comision(juego: str) -> float:
    if juego in ("Quiniela", "Quiniela Ya", "Poceada", "Tombolina", "LT"):
        return 0.20
    if juego in ("Loto", "Loto 5", "Quini 6", "Brinco"):
        return 0.14
    return 0.0


def set_row_vals(w: PlanillaWidgets, row_index: int, vals: list[str]):
    ids = w.data_tree.get_children()
    if 0 <= row_index < len(ids):
        w.data_tree.item(ids[row_index], values=vals)


def get_row_vals(w: PlanillaWidgets, row_index: int) -> list[str]:
    ids = w.data_tree.get_children()
    if 0 <= row_index < len(ids):
        return [str(v) for v in w.data_tree.item(ids[row_index], "values")]
    return [""] * len(w.cols)


def buscar_fila_por_sorteo(w: PlanillaWidgets, sorteo: int) -> int | None:
    s = str(int(sorteo))
    for idx, iid in enumerate(w.data_tree.get_children()):
        vals = w.data_tree.item(iid, "values")
        if vals and str(vals[0]).strip() == s:
            return idx
    return None


def asegurar_fila_sorteo(w: PlanillaWidgets, sorteo: int) -> int:
    """
    CREA sorteos (filas). SOLO usar desde PJU.
    Tickets/Reporte/SFA no deben crear filas.
    """
    idx = buscar_fila_por_sorteo(w, sorteo)
    if idx is not None:
        return idx

    for idx2, iid in enumerate(w.data_tree.get_children()):
        vals = w.data_tree.item(iid, "values")
        if not vals or not str(vals[0]).strip():
            row = get_row_vals(w, idx2)
            row[0] = str(int(sorteo))
            set_row_vals(w, idx2, row)
            return idx2

    last_idx = len(w.data_tree.get_children()) - 1
    row = get_row_vals(w, last_idx)
    row[0] = str(int(sorteo))
    set_row_vals(w, last_idx, row)
    return last_idx


def obtener_sorteos_visibles_ordenados(w: PlanillaWidgets) -> list[int]:
    sorteos: list[int] = []
    for iid in w.data_tree.get_children():
        vals = w.data_tree.item(iid, "values")
        if not vals:
            continue
        txt = str(vals[0]).strip()
        if not txt:
            continue
        try:
            sorteos.append(int(txt))
        except Exception:
            pass
    return sorted(sorteos)


def _extraer_sorteos_unicos_ordenados_desde_semanas(semanas: dict[int, list[int]]) -> list[int]:
    sorteos: set[int] = set()
    for vals in (semanas or {}).values():
        if not isinstance(vals, list):
            continue
        for v in vals:
            try:
                sorteos.add(int(v))
            except Exception:
                pass
    return sorted(sorteos)


def _reconstruir_semanas_desde_filas(filas: list[list[str]] | None, sorteos_por_semana: int = 7) -> dict[int, list[int]]:
    """Reconstruye un mapa mínimo sin dividir artificialmente la semana."""
    if not isinstance(filas, list):
        return {}

    sorteos: list[int] = []
    vistos: set[int] = set()
    for fila in filas:
        if not isinstance(fila, list) or not fila:
            continue
        try:
            sorteo = int(str(fila[0]).strip())
        except Exception:
            continue
        if sorteo in vistos:
            continue
        vistos.add(sorteo)
        sorteos.append(sorteo)

    if not sorteos:
        return {}

    return {1: sorted(sorteos)}


def _append_pju_si_continua_desde_ultimo_sorteo_visible(
    w: PlanillaWidgets,
    semanas_importadas: dict[int, list[int]],
    estado_var=None,
) -> bool:
    """
    Si el último sorteo visible en la planilla es justo el anterior
    al primer sorteo del PJU importado, agrega sin borrar.
    Devuelve True si anexó; False si no correspondía anexar.
    """
    sorteos_actuales = obtener_sorteos_visibles_ordenados(w)
    sorteos_importados = _extraer_sorteos_unicos_ordenados_desde_semanas(semanas_importadas)

    if not sorteos_actuales or not sorteos_importados:
        return False

    ultimo_editado_manual = getattr(w, "_ultimo_sorteo_editado_manual", None)
    try:
        ultimo_editado_manual = int(ultimo_editado_manual)
    except Exception:
        ultimo_editado_manual = None

    ultimo_visible = (
        ultimo_editado_manual
        if (ultimo_editado_manual is not None and ultimo_editado_manual in set(sorteos_actuales))
        else max(sorteos_actuales)
    )
    primer_importado = min(sorteos_importados)

    if ultimo_visible + 1 != primer_importado:
        return False

    existentes = set(sorteos_actuales)
    agregados = 0

    for s in sorteos_importados:
        if s not in existentes:
            asegurar_fila_sorteo(w, s)
            agregados += 1

    refresh_planilla(w, estado_var)

    if estado_var is not None:
        estado_var.set(
            f"{w.juego}: PJU anexado sin sobreescribir. "
            f"Último visible={ultimo_visible}, primer importado={primer_importado}, agregados={agregados}."
        )

    return True


def _anclar_sorteos_importados_a_semana_1(semanas_importadas: dict[int, list[int]] | None) -> dict[int, list[int]]:
    """
    Regla de negocio:
    cuando se viene encadenando por continuidad manual + PJU, los sorteos
    importados deben consolidarse en Semana 1.
    """
    sorteos = _extraer_sorteos_unicos_ordenados_desde_semanas(semanas_importadas or {})
    if not sorteos:
        return {}
    return {1: sorteos}


# ======================
# APLICAR DATOS (SOLO si el sorteo EXISTE por PJU)
# ======================

def aplicar_tickets(w: PlanillaWidgets):
    juego = w.juego
    resumen = app_state.tickets_resumen_por_juego.get(juego, {}) if hasattr(app_state, "tickets_resumen_por_juego") else {}
    tasa = get_tasa_comision(juego)

    for sorteo, data in (resumen or {}).items():
        try:
            s = int(sorteo)
        except Exception:
            continue

        idx = buscar_fila_por_sorteo(w, s)
        if idx is None:
            continue

        row = get_row_vals(w, idx)
        rec = float(data.get("recaud", 0.0) or 0.0)
        prem = float(data.get("prem", 0.0) or 0.0)
        comi = rec * tasa

        row[1] = fmt_pesos(rec)
        row[2] = fmt_pesos(comi)
        row[3] = fmt_pesos(prem)
        set_row_vals(w, idx, row)


def aplicar_reporte(w: PlanillaWidgets):
    juego = w.juego
    resumen = app_state.reporte_resumen_por_juego.get(juego, {}) if hasattr(app_state, "reporte_resumen_por_juego") else {}

    for sorteo, data in (resumen or {}).items():
        try:
            s = int(sorteo)
        except Exception:
            continue

        idx = buscar_fila_por_sorteo(w, s)
        if idx is None:
            continue

        row = get_row_vals(w, idx)
        rec = float(data.get("recaud", 0.0) or 0.0)
        comi = float(data.get("comi", 0.0) or 0.0)
        prem = float(data.get("prem", 0.0) or 0.0)

        row[4] = fmt_pesos(rec)
        row[5] = fmt_pesos(comi)
        row[6] = fmt_pesos(prem)
        set_row_vals(w, idx, row)


def aplicar_sfa(w: PlanillaWidgets):
    juego = w.juego
    resumen = app_state.sfa_resumen_por_juego.get(juego, {}) if hasattr(app_state, "sfa_resumen_por_juego") else {}

    for sorteo, data in (resumen or {}).items():
        try:
            s = int(sorteo)
        except Exception:
            continue

        idx = buscar_fila_por_sorteo(w, s)
        if idx is None:
            continue

        row = get_row_vals(w, idx)
        rec = float(data.get("recaud", 0.0) or 0.0)
        comi = float(data.get("comi", 0.0) or 0.0)
        prem = float(data.get("prem", 0.0) or 0.0)

        row[10] = fmt_pesos(rec)
        row[11] = fmt_pesos(comi)
        row[12] = fmt_pesos(prem)
        set_row_vals(w, idx, row)


# ======================
# DIFERENCIAS
# ======================

def _fmt_diff_value(valor: float) -> str:
    return fmt_pesos(valor)


def _calc_diff_cell(a_str: str, b_str: str) -> str:
    if not _cell_has_value(a_str) or not _cell_has_value(b_str):
        return ""

    a = parse_pesos(a_str)
    b = parse_pesos(b_str)
    if a is None or b is None:
        return ""

    return _fmt_diff_value(b - a)


_DIFF_ROW_WARN_MIN = 20.0
_DIFF_ROW_DANGER_MIN = 40.0
_DIFF_ROW_WARN_BG = "#FFF2CC"
_DIFF_ROW_DANGER_BG = "#FDE2E1"
_DIFF_COLUMNS = (7, 8, 9, 13, 14, 15)
_TOTAL_ROW_BG = "#FFF9CC"
_DIFF_RENDER_DEBOUNCE_MS = 90
_DIFF_RENDER_DEBOUNCE_SCROLL_MS = 90
_USE_DIFF_OVERLAY_LABELS = False


def _diff_foreground(valor: str) -> str:
    return "#0F172A"


def _clear_diff_labels(w: PlanillaWidgets, *, destroy: bool = False):
    labels = getattr(w, "diff_labels", [])
    for label in labels:
        try:
            if destroy:
                label.destroy()
            else:
                label.place_forget()
        except Exception:
            pass
    if destroy:
        w.diff_labels.clear()




def _ensure_diff_label_pool(w: PlanillaWidgets, count: int):
    labels = getattr(w, "diff_labels", [])
    while len(labels) < count:
        labels.append(tk.Label(
            w.data_tree,
            anchor="e",
            padx=6,
            pady=0,
            borderwidth=0,
            highlightthickness=0,
            font=("Segoe UI", 9),
        ))
def _row_background(tree: ttk.Treeview, iid: str) -> str:
    if iid in tree.selection():
        return "#DCEBFF"

    tags = tree.item(iid, "tags")
    if "total" in tags:
        return _TOTAL_ROW_BG
    if "diff_danger" in tags:
        return _DIFF_ROW_DANGER_BG
    if "diff_warn" in tags:
        return _DIFF_ROW_WARN_BG
    if "odd" in tags:
        return "#F8FAFC"
    return "#FFFFFF"


def _visible_item_ids(tree: ttk.Treeview) -> list[str]:
    visibles: list[str] = []
    alto = tree.winfo_height()
    if alto <= 0:
        return visibles

    primero = tree.identify_row(1)
    if not primero:
        return visibles

    ultimo = tree.identify_row(max(1, alto - 2))
    iid = primero
    while iid:
        visibles.append(iid)
        if iid == ultimo:
            break
        iid = tree.next(iid)
    return visibles


def _render_diff_labels(w: PlanillaWidgets):
    if not _USE_DIFF_OVERLAY_LABELS:
        _clear_diff_labels(w)
        return

    _clear_diff_labels(w)
    tree = w.data_tree
    if not tree.winfo_ismapped():
        return
    visibles = tree.winfo_height()
    if visibles <= 0:
        return

    diff_cells: list[tuple[int, int, int, int, str, str, str]] = []
    for iid in _visible_item_ids(tree):
        fondo_fila = _row_background(tree, iid)
        valores = [str(v) for v in tree.item(iid, "values")]
        for col_idx in _DIFF_COLUMNS:
            if col_idx >= min(len(valores), len(w.cols)):
                continue
            valor = valores[col_idx].strip()
            if not valor:
                continue

            color_texto = _diff_foreground(valor)

            bbox = tree.bbox(iid, f"#{col_idx + 1}")
            if not bbox:
                continue

            x, y, width, height = bbox
            if y >= visibles or (y + height) <= 0 or width <= 2 or height <= 2:
                continue

            diff_cells.append((x, y, width, height, valor, fondo_fila, color_texto))

    _ensure_diff_label_pool(w, len(diff_cells))
    for idx, (x, y, width, height, valor, fondo_fila, foreground) in enumerate(diff_cells):
        label = w.diff_labels[idx]
        label.configure(text=valor, bg=fondo_fila, fg=foreground)
        label.place(x=x + 1, y=y + 1, width=width - 2, height=height - 2)


def _clasificar_diferencia(valor: str) -> str | None:
    numero = parse_pesos(valor)
    if numero is None:
        return None
    numero = abs(float(numero))
    if numero > _DIFF_ROW_DANGER_MIN:
        return "danger"
    if numero > _DIFF_ROW_WARN_MIN:
        return "warn"
    return None


def calcular_dif_tickets_vs_reporte(w: PlanillaWidgets):
    for idx, _iid in enumerate(w.data_tree.get_children()):
        row = get_row_vals(w, idx)
        t_rec, t_com, t_pre = row[1], row[2], row[3]
        r_rec, r_com, r_pre = row[4], row[5], row[6]

        row[7] = _calc_diff_cell(t_rec, r_rec)
        row[8] = _calc_diff_cell(t_com, r_com)
        row[9] = _calc_diff_cell(t_pre, r_pre)
        set_row_vals(w, idx, row)


def calcular_dif_tickets_vs_sfa(w: PlanillaWidgets):
    for idx, _iid in enumerate(w.data_tree.get_children()):
        row = get_row_vals(w, idx)
        t_rec, t_com, t_pre = row[1], row[2], row[3]
        s_rec, s_com, s_pre = row[10], row[11], row[12]

        row[13] = _calc_diff_cell(t_rec, s_rec)
        row[14] = _calc_diff_cell(t_com, s_com)
        row[15] = _calc_diff_cell(t_pre, s_pre)
        set_row_vals(w, idx, row)



def _actualizar_fila_totales_planilla(w: PlanillaWidgets):
    total_iid = _ensure_total_row_iid(w)
    ids = list(w.data_tree.get_children())

    last_data_idx = -1
    sumas = [0.0] * 15

    for idx, iid in enumerate(ids):
        if iid == total_iid:
            continue
        vals = [str(v) for v in w.data_tree.item(iid, "values")]
        sorteo_txt = (vals[0] if vals else "").strip().lower()
        if sorteo_txt == "totales":
            continue
        try:
            int((vals[0] if vals else "").strip())
            es_sorteo = True
        except Exception:
            es_sorteo = False
        if not es_sorteo:
            continue

        last_data_idx = idx
        for col in range(1, min(16, len(vals))):
            n = parse_pesos(vals[col])
            if n is not None:
                sumas[col - 1] += float(n)

    row = [""] * len(w.cols)
    row[0] = "Totales"
    if last_data_idx >= 0:
        for col in range(1, min(16, len(w.cols))):
            row[col] = fmt_pesos(sumas[col - 1])

    idx_total = 0 if last_data_idx < 0 else min(last_data_idx + 1, max(len(ids) - 1, 0))
    try:
        w.data_tree.move(total_iid, "", idx_total)
    except Exception:
        pass
    w.data_tree.item(total_iid, values=row)

    if total_iid not in getattr(w, "all_item_ids", []):
        w.all_item_ids = list(getattr(w, "all_item_ids", []) or []) + [total_iid]

def _aplicar_zebra_planilla(w: PlanillaWidgets):
    tree = w.data_tree
    tree.tag_configure("even", background="#FFFFFF", foreground="#0F172A")
    tree.tag_configure("odd", background="#F7FAFD", foreground="#0F172A")
    tree.tag_configure("empty", foreground="#9CA3AF")
    tree.tag_configure("total", background="#EAF6ED", foreground="#1F2937", font=("Segoe UI Semibold", 9))
    for idx, iid in enumerate(tree.get_children()):
        vals = tree.item(iid, "values")
        has_data = any(str(v).strip() for v in vals)
        es_total = bool(vals and str(vals[0]).strip().lower() == "totales")
        base_tag = "even" if idx % 2 == 0 else "odd"
        if es_total:
            tree.item(iid, tags=("total",))
            continue
        if not has_data:
            tree.item(iid, tags=(base_tag, "empty"))
            continue

        tree.item(iid, tags=(base_tag,))


def refresh_planilla(w: PlanillaWidgets, estado_var=None):
    _restaurar_items_planilla(w)
    aplicar_tickets(w)
    aplicar_reporte(w)
    aplicar_sfa(w)
    calcular_dif_tickets_vs_reporte(w)
    calcular_dif_tickets_vs_sfa(w)
    _actualizar_fila_totales_planilla(w)
    _aplicar_zebra_planilla(w)
    _render_diff_labels(w)
    _aplicar_filtros_planilla(w)
    _notificar_recalculo_totales_planilla()


def cargar_sorteos_en_planilla(
    w: PlanillaWidgets,
    sorteos: list[int],
    estado_var=None,
    append_if_contiguous: bool = False,
    sorteo_preferido: int | None = None,
):
    sorteos_nuevos = []
    for s in sorteos:
        try:
            sorteos_nuevos.append(int(s))
        except Exception:
            pass

    if append_if_contiguous and sorteos_nuevos:
        sorteos_actuales = obtener_sorteos_visibles_ordenados(w)
        if sorteos_actuales:
            ultimo_existente = max(sorteos_actuales)
            primer_nuevo = min(sorteos_nuevos)

            if ultimo_existente == (primer_nuevo - 1):
                existentes_set = set(sorteos_actuales)
                agregados = 0
                for s in sorteos_nuevos:
                    if s not in existentes_set:
                        asegurar_fila_sorteo(w, s)
                        agregados += 1

                refresh_planilla(w, estado_var)
                if estado_var is not None:
                    estado_var.set(
                        f"{w.juego}: PJU agregado sin sobreescribir. "
                        f"Último existente={ultimo_existente}, primer nuevo={primer_nuevo}, agregados={agregados}."
                    )
                return

    _aplicar_filas_a_widget(w, [], clear_rows=True)

    for s in sorteos_nuevos:
        asegurar_fila_sorteo(w, s)

    # Rehidratar valores previamente guardados para estos sorteos antes de
    # recalcular/importar. Así, al cambiar de semana y volver, no se pierden
    # ediciones manuales (por ejemplo columnas de Reporte cargadas a mano).
    try:
        data_juego_guardado = _leer_planilla_juego_desde_storage(w.juego)
        filas_guardadas = data_juego_guardado.get("filas", []) if isinstance(data_juego_guardado, dict) else []
        guardadas_por_sorteo, _ = _mapear_filas_por_sorteo(filas_guardadas, len(w.cols))
        if guardadas_por_sorteo:
            for idx, s in enumerate(sorteos_nuevos):
                fila_guardada = guardadas_por_sorteo.get(int(s))
                if not fila_guardada:
                    continue
                set_row_vals(w, idx, list(fila_guardada))
    except Exception:
        pass

    refresh_planilla(w, estado_var)

    target_sorteo = None
    if sorteo_preferido is not None:
        try:
            target_sorteo = int(sorteo_preferido)
        except Exception:
            target_sorteo = None
    if target_sorteo is None and sorteos_nuevos:
        target_sorteo = sorteos_nuevos[0]

    if target_sorteo is not None:
        target_iid = None
        for iid in w.data_tree.get_children():
            vals = [str(v) for v in w.data_tree.item(iid, "values")]
            if not vals:
                continue
            try:
                if int(str(vals[0]).strip()) == target_sorteo:
                    target_iid = iid
                    break
            except Exception:
                continue
        if target_iid is not None:
            try:
                w.data_tree.selection_set(target_iid)
                w.data_tree.focus(target_iid)
                w.data_tree.see(target_iid)
            except Exception:
                pass

    if estado_var is not None:
        estado_var.set(f"{w.juego}: cargados {len(sorteos_nuevos)} sorteos desde PJU.")


# ======================
# UI: PLANILLA EXCEL-LIKE
# ======================

def crear_planilla_excel_like(parent: ttk.Frame, juego: str, filas: int, estado_var, on_import_pju=None) -> PlanillaWidgets:
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(0, weight=1)

    style = ttk.Style(parent)
    tree_style = f"Planilla.{juego}.Treeview"
    heading_style = f"Planilla.{juego}.Treeview.Heading"
    style.configure(
        tree_style,
        rowheight=30,
        font=("Segoe UI", 9),
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        bordercolor="#D7E1EB",
        lightcolor="#D7E1EB",
        darkcolor="#D7E1EB",
        foreground="#0F172A",
        relief="flat",
    )
    style.configure(
        heading_style,
        font=("Segoe UI Semibold", 10),
        background="#F4F7FB",
        foreground="#1F2937",
        relief="flat",
    )
    style.map(
        tree_style,
        background=[("selected", "#DCEBFF")],
        foreground=[("selected", "#0F172A")],
    )

    cols = [
        "sorteo",
        "t_recaud", "t_comi", "t_prem",
        "r_recaud", "r_comi", "r_prem",
        "d1_recaud", "d1_comi", "d1_prem",
        "s_recaud", "s_comi", "s_prem",
        "d2_recaud", "d2_comi", "d2_prem",
    ]

    card = tk.Frame(
        parent,
        bg="#E7EEF6",
        highlightthickness=1,
        highlightbackground="#C7D5E4",
        bd=0,
    )
    card.grid(row=0, column=0, sticky="nsew", padx=8, pady=(2, 8))
    card.columnconfigure(0, weight=1)
    card.rowconfigure(0, weight=1)

    cont = ttk.Frame(card, style="Panel.TFrame")
    cont.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
    cont.columnconfigure(0, weight=1)
    cont.rowconfigure(0, weight=0)
    cont.rowconfigure(1, weight=0)
    cont.rowconfigure(2, weight=0)
    cont.rowconfigure(3, weight=1)
    cont.rowconfigure(4, weight=0)

    GRIS_BASE = "#E9EEF5"
    COLOR_TICKETS = "#BFD7F6"
    COLOR_REPORTE = "#BFEBD3"
    COLOR_DIF = "#FCE9A7"
    COLOR_SFA = "#FFD7AA"

    BORDER_TICKETS = _darken(COLOR_TICKETS, 0.74)
    BORDER_REPORTE = _darken(COLOR_REPORTE, 0.74)
    BORDER_DIF = _darken(COLOR_DIF, 0.74)
    BORDER_SFA = _darken(COLOR_SFA, 0.74)
    BORDER_GRIS = _darken(GRIS_BASE, 0.74)

    W_SORTEO = 118
    W_COL = 146
    MIN_W_SORTEO = 102
    MIN_W_COL = 124
    widths_base = [W_SORTEO] + [W_COL] * 15

    H1 = 34
    H2 = 30
    header_canvas = tk.Canvas(cont, height=H1, highlightthickness=0, bg=GRIS_BASE, bd=0)
    header_canvas.grid(row=0, column=0, sticky="ew")

    btn_importar_pju = ttk.Button(
        header_canvas,
        text="Importar PJU",
        style="Marino.TButton",
        command=on_import_pju,
    )
    btn_pju_window_id = None
    if on_import_pju is None:
        btn_importar_pju.state(["disabled"])

    header2_canvas = tk.Canvas(cont, height=H2, highlightthickness=0, bg=GRIS_BASE, bd=0)
    header2_canvas.grid(row=1, column=0, sticky="ew")

    H3 = 0
    filter_canvas = tk.Canvas(cont, height=H3, highlightthickness=0, bg="#F8FBFD", bd=0)

    tree = ttk.Treeview(cont, columns=cols, show="tree", height=18, style=tree_style)
    tree.grid(row=3, column=0, sticky="nsew")
    tree.column("#0", width=0, stretch=False)
    tree.heading("#0", text="")

    vs = ttk.Scrollbar(cont, orient="vertical")
    vs.grid(row=3, column=1, sticky="ns")

    hs = ttk.Scrollbar(cont, orient="horizontal")
    hs.grid(row=4, column=0, sticky="ew")

    filter_vars = {col: tk.StringVar() for col in cols}
    header_labels = {
        "sorteo": "Sorteo",
        "t_recaud": "Tickets / Recaudación",
        "t_comi": "Tickets / Comisión",
        "t_prem": "Tickets / Premios",
        "r_recaud": "Reporte / Recaudación",
        "r_comi": "Reporte / Comisión",
        "r_prem": "Reporte / Premios",
        "d1_recaud": "Diferencias / Recaudación",
        "d1_comi": "Diferencias / Comisión",
        "d1_prem": "Diferencias / Premios",
        "s_recaud": "SFA / Recaudación",
        "s_comi": "SFA / Comisión",
        "s_prem": "SFA / Premios",
        "d2_recaud": "Dif SFA / Recaudación",
        "d2_comi": "Dif SFA / Comisión",
        "d2_prem": "Dif SFA / Premios",
    }

    for _ in range(filas):
        tree.insert("", "end", text="", values=[""] * len(cols))

    def _aplicar_estilo_filas():
        _actualizar_fila_totales_planilla(w_tmp)
        _aplicar_zebra_planilla(w_tmp)
        _request_diff_render()

    resize_job = None
    last_x0 = None
    diff_render_job = None
    diff_render_after_job = None
    diff_render_debounce_job = None
    diff_labels_hidden = False

    def _draw_headers(widths: list[int]):
        nonlocal btn_pju_window_id
        header_canvas.delete("all")
        header2_canvas.delete("all")
        filter_canvas.delete("all")
        btn_pju_window_id = None

        x = 0

        w0 = widths[0]
        header_canvas.create_rectangle(x, 0, x + w0, H1, fill=GRIS_BASE, outline=BORDER_GRIS, width=1)
        header_canvas.create_line(x, H1 - 1, x + w0, H1 - 1, fill="#B7C8DA", width=1)
        header_canvas.create_text(x + w0 / 2, H1 / 2, font=("Segoe UI Semibold", 10))
        x += w0

        tickets_x_inicio = x
        w = sum(widths[1:4])
        tickets_x_fin = x + w
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_TICKETS, outline=BORDER_TICKETS, width=1)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Tickets", font=("Segoe UI Semibold", 10))
        x += w

        w = sum(widths[4:7])
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_REPORTE, outline=BORDER_REPORTE, width=1)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Reporte", font=("Segoe UI Semibold", 10))
        x += w

        w = sum(widths[7:10])
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_DIF, outline=BORDER_DIF, width=1)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Diferencias", font=("Segoe UI Semibold", 10))
        x += w

        w = sum(widths[10:13])
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_SFA, outline=BORDER_SFA, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="SFA", font=("Segoe UI Semibold", 10))
        x += w

        w = sum(widths[13:16])
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_DIF, outline=BORDER_DIF, width=1)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Dif SFA", font=("Segoe UI Semibold", 10))

        linea_suave = "#8CA3B8"
        for xpos in (tickets_x_inicio, tickets_x_fin):
            header_canvas.create_line(xpos, 5, xpos, H1 - 5, fill=linea_suave, width=1)
            header_canvas.create_line(xpos + 1, 6, xpos + 1, H1 - 6, fill="#DDE7F0", width=1)

        if btn_importar_pju is not None:
            if btn_pju_window_id is None:
                btn_pju_window_id = header_canvas.create_window(
                    1,
                    1,
                    anchor="nw",
                    width=max(1, widths[0] - 2),
                    height=max(1, H1 - 2),
                    window=btn_importar_pju,
                    tags=("btn_pju",),
                )
            else:
                header_canvas.coords(btn_pju_window_id, 1, 1)
                header_canvas.itemconfigure(
                    btn_pju_window_id,
                    width=max(1, widths[0] - 2),
                    height=max(1, H1 - 2),
                )

        texts = (
            ["Sorteo"]
            + ["Recaudación", "Comisión", "Premios"]
            + ["Recaudación", "Comisión", "Premios"]
            + ["Recaudación", "Comisión", "Premios"]
            + ["Recaudación", "Comisión", "Premios"]
            + ["Recaudación", "Comisión", "Premios"]
        )
        bgs = (
            [GRIS_BASE]
            + [COLOR_TICKETS] * 3
            + [COLOR_REPORTE] * 3
            + [COLOR_DIF] * 3
            + [COLOR_SFA] * 3
            + [COLOR_DIF] * 3
        )
        borders = (
            [BORDER_GRIS]
            + [BORDER_TICKETS] * 3
            + [BORDER_REPORTE] * 3
            + [BORDER_DIF] * 3
            + [BORDER_SFA] * 3
            + [BORDER_DIF] * 3
        )

        x = 0
        for tx, wpx, bg, br in zip(texts, widths, bgs, borders):
            header2_canvas.create_rectangle(x, 0, x + wpx, H2, fill=bg, outline=br, width=1)
            header2_canvas.create_text(x + wpx / 2, H2 / 2, text=tx, font=("Segoe UI Semibold", 9))
            x += wpx

        total = sum(widths)
        header_canvas.configure(scrollregion=(0, 0, total, H1))
        header2_canvas.configure(scrollregion=(0, 0, total, H2))
        filter_canvas.configure(scrollregion=(0, 0, total, H3))

    def _set_column_widths() -> tuple[list[int], bool]:
        widths = list(widths_base)
        min_widths = [MIN_W_SORTEO] + [MIN_W_COL] * (len(widths) - 1)

        # Usar el ancho visible real del Treeview para evitar "huecos"
        # a la derecha (por ejemplo, al final del bloque Diferencias).
        tree_viewport = tree.winfo_width()
        if tree_viewport <= 1:
            scrollbar_w = vs.winfo_width() if vs.winfo_ismapped() else 0
            tree_viewport = max(0, cont.winfo_width() - scrollbar_w)
        available = max(0, tree_viewport)
        current = sum(widths)
        if available > current:
            extra = available - current
            grow_cols = len(widths) - 1
            add_each = extra // grow_cols if grow_cols else 0
            remainder = extra % grow_cols if grow_cols else 0
            for i in range(1, len(widths)):
                widths[i] += add_each
                if (i - 1) < remainder:
                    widths[i] += 1
        elif 0 < available < current:
            ratio = available / current
            widths = [max(min_widths[i], int(widths[i] * ratio)) for i in range(len(widths))]
            ajuste = available - sum(widths)
            if ajuste > 0:
                for i in range(1, len(widths)):
                    if ajuste <= 0:
                        break
                    widths[i] += 1
                    ajuste -= 1

        min_total = sum(min_widths)
        needs_horizontal_scroll = available < min_total

        tree.column("sorteo", width=widths[0], anchor="center", stretch=False)
        for i, c in enumerate(cols[1:], start=1):
            tree.column(c, width=widths[i], anchor="center", stretch=False)

        return widths, needs_horizontal_scroll

    def _actualizar_textos_header_filtros():
        header2_canvas.delete("filter_icon")
        x = 0
        for col, wpx in zip(cols, [tree.column(c, option="width") for c in cols]):
            if str(filter_vars.get(col).get() or "").strip():
                header2_canvas.create_text(
                    x + wpx - 12,
                    H2 / 2,
                    text="🔍",
                    font=("Segoe UI Emoji", 9),
                    fill="#1F2937",
                    tags=("filter_icon",),
                )
            x += wpx

    def _show_filter_popup(col: str):
        win = tk.Toplevel(parent)
        win.title(f"Filtro: {header_labels.get(col, col)}")
        win.transient(parent.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(
            win,
            text=f"Filtrar '{header_labels.get(col, col)}' (contiene):",
        ).grid(row=0, column=0, columnspan=3, padx=10, pady=(10, 4), sticky="w")

        justify = "center" if col == "sorteo" else "right"
        entry = ttk.Entry(win, width=34, textvariable=tk.StringVar(value=filter_vars[col].get()), justify=justify)
        entry.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="ew")
        entry.focus_set()
        entry.icursor("end")

        def _aplicar():
            filter_vars[col].set(str(entry.get() or "").strip())
            if estado_var is not None:
                if filter_vars[col].get():
                    estado_var.set(f"Filtro aplicado en {header_labels.get(col, col)}.")
                else:
                    estado_var.set(f"Filtro limpiado en {header_labels.get(col, col)}.")
            win.destroy()

        def _limpiar_columna():
            filter_vars[col].set("")
            if estado_var is not None:
                estado_var.set(f"Filtro limpiado en {header_labels.get(col, col)}.")
            win.destroy()

        def _limpiar_todo():
            for var in filter_vars.values():
                var.set("")
            if estado_var is not None:
                estado_var.set("Filtros limpiados.")
            win.destroy()

        ttk.Button(win, text="Aplicar", command=_aplicar, style="Marino.TButton").grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")
        ttk.Button(win, text="Limpiar columna", command=_limpiar_columna).grid(row=2, column=1, padx=5, pady=(0, 10))
        ttk.Button(win, text="Limpiar todo", command=_limpiar_todo).grid(row=2, column=2, padx=(5, 10), pady=(0, 10), sticky="e")

        entry.bind("<Return>", lambda _e: _aplicar())

    def _columna_desde_x(x_click: int) -> str | None:
        x_real = header2_canvas.canvasx(x_click)
        acum = 0
        for col, wpx in zip(cols, [tree.column(c, option="width") for c in cols]):
            if acum <= x_real < (acum + wpx):
                return col
            acum += wpx
        return None

    def _on_header2_click(evt):
        col = _columna_desde_x(evt.x)
        if not col:
            return
        _show_filter_popup(col)

    def _sync_headers(force=False):
        nonlocal last_x0
        x0, _ = tree.xview()
        if (not force) and (last_x0 is not None) and abs(x0 - last_x0) < 1e-9:
            return
        last_x0 = x0
        header_canvas.xview_moveto(x0)
        header2_canvas.xview_moveto(x0)
        filter_canvas.xview_moveto(x0)

    def _apply_layout_sync():
        nonlocal resize_job
        resize_job = None
        widths, _needs_horizontal_scroll = _set_column_widths()
        _draw_headers(widths)
        hs.grid(row=4, column=0, sticky="ew")
        _sync_headers(force=True)

    def _request_layout_sync(_evt=None):
        nonlocal resize_job
        if resize_job is not None:
            return
        resize_job = cont.after_idle(_apply_layout_sync)

    def _render_diff_labels_idle():
        nonlocal diff_render_job, diff_render_after_job, diff_labels_hidden
        diff_render_after_job = None
        diff_render_job = None
        diff_labels_hidden = False
        if not cont.winfo_exists() or not tree.winfo_exists():
            return
        _render_diff_labels(w_tmp)

    def _hide_diff_labels():
        nonlocal diff_labels_hidden
        if diff_labels_hidden:
            return
        diff_labels_hidden = True
        _clear_diff_labels(w_tmp)

    def _request_diff_render(_evt=None, *, quick: bool = False):
        nonlocal diff_render_job, diff_render_after_job, diff_render_debounce_job
        if not cont.winfo_exists():
            return
        if diff_render_debounce_job is not None:
            try:
                cont.after_cancel(diff_render_debounce_job)
            except Exception:
                pass
        delay = _DIFF_RENDER_DEBOUNCE_SCROLL_MS if quick else _DIFF_RENDER_DEBOUNCE_MS
        diff_render_debounce_job = cont.after(delay, _request_diff_render_idle)

    def _request_diff_render_quick(_evt=None):
        _request_diff_render(_evt, quick=True)

    def _request_diff_render_idle():
        nonlocal diff_render_job, diff_render_after_job, diff_render_debounce_job
        diff_render_debounce_job = None
        if not cont.winfo_exists():
            return
        if diff_render_job is not None or diff_render_after_job is not None:
            return
        diff_render_after_job = cont.after_idle(_render_diff_labels_idle)

    def _xscroll(f, l):
        hs.set(f, l)
        _sync_headers()
        _hide_diff_labels()
        _request_diff_render_quick()

    def _tree_xview(*args):
        tree.xview(*args)
        _sync_headers(force=True)
        _hide_diff_labels()
        _request_diff_render_quick()

    def _tree_yview(*args):
        tree.yview(*args)
        _hide_diff_labels()
        _request_diff_render_quick()

    def _on_tree_yscroll(f, l):
        vs.set(f, l)
        _hide_diff_labels()
        _request_diff_render_quick()

    tree.configure(xscrollcommand=_xscroll, yscrollcommand=_on_tree_yscroll)
    hs.configure(command=_tree_xview)
    vs.configure(command=_tree_yview)
    bind_smooth_mousewheel(
        tree=tree,
        targets=(tree, cont, header_canvas, header2_canvas, filter_canvas),
        on_scroll=_request_diff_render_quick,
    )
    tree.bind("<Configure>", _request_layout_sync)
    tree.bind("<Configure>", _request_diff_render, add="+")
    tree.bind("<<TreeviewSelect>>", _request_diff_render, add="+")
    for event_name in (
        "<KeyRelease-Up>",
        "<KeyRelease-Down>",
        "<KeyRelease-Left>",
        "<KeyRelease-Right>",
        "<KeyRelease-Prior>",
        "<KeyRelease-Next>",
        "<KeyRelease-Home>",
        "<KeyRelease-End>",
    ):
        tree.bind(event_name, _request_diff_render_quick, add="+")
    cont.bind("<Configure>", _request_layout_sync)

    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)
    edit_entry = None

    def _normalizar_valor_edicion(col_idx: int, valor_raw: str) -> str:
        valor = str(valor_raw or "").strip()
        if col_idx == 0:
            if not valor:
                return ""
            digitos = "".join(ch for ch in valor if ch.isdigit())
            return str(int(digitos)) if digitos else ""
        if not valor:
            return ""
        n = parse_pesos(valor)
        return fmt_pesos(n) if n is not None else valor

    def _recalcular_grilla_editable():
        calcular_dif_tickets_vs_reporte(w_tmp)
        calcular_dif_tickets_vs_sfa(w_tmp)
        _actualizar_fila_totales_planilla(w_tmp)
        _aplicar_estilo_filas()
        _notificar_recalculo_totales_planilla()
        _request_diff_render()
        on_grid_changed = getattr(w_tmp, "on_grid_changed", None)
        if callable(on_grid_changed):
            try:
                on_grid_changed()
            except Exception:
                pass

    def _cerrar_editor(save=True):
        nonlocal edit_entry
        if edit_entry is None:
            return
        if save:
            try:
                iid, col = edit_entry._cell
                col_idx = int(col.replace("#", "")) - 1
                vals = list(tree.item(iid, "values"))
                while len(vals) < len(cols):
                    vals.append("")
                nuevo_valor = _normalizar_valor_edicion(col_idx, edit_entry.get())
                if vals[col_idx] != nuevo_valor:
                    push_undo_rows(undo_state, tree, [(iid, list(vals))], meta={"seccion": "area_recaudacion"})
                    vals[col_idx] = nuevo_valor
                    tree.item(iid, values=vals)
                    if col_idx == 0:
                        try:
                            w_tmp._ultimo_sorteo_editado_manual = int(str(nuevo_valor).strip())
                        except Exception:
                            w_tmp._ultimo_sorteo_editado_manual = None
                    _recalcular_grilla_editable()
            except Exception:
                pass

        try:
            edit_entry.destroy()
        except Exception:
            pass
        edit_entry = None

    def _copiar_celdas(_evt=None):
        _cerrar_editor(save=True)
        iid, col_idx = get_anchor_cell(tree, clipboard_state, default_col=0)
        if not iid or col_idx < 0 or col_idx >= len(cols):
            return "break"

        seleccion = ordered_selected_rows(tree)
        row_ids = seleccion if len(seleccion) > 1 and iid in seleccion else [iid]
        matrix = []
        for row_iid in row_ids:
            vals = [str(v) for v in tree.item(row_iid, "values")]
            if vals and str(vals[0]).strip().lower() == "totales":
                continue
            while len(vals) < len(cols):
                vals.append("")
            matrix.append([vals[col_idx]])
        set_clipboard_matrix(tree, matrix)
        return "break"

    def _pegar_celdas(_evt=None):
        _cerrar_editor(save=True)
        matrix = get_clipboard_matrix(tree)
        if not matrix:
            return "break"

        anchor_iid, anchor_col = get_anchor_cell(tree, clipboard_state, default_col=0)
        row_ids = []
        for row_iid in tree.get_children():
            vals = [str(v) for v in tree.item(row_iid, "values")]
            if vals and str(vals[0]).strip().lower() == "totales":
                continue
            row_ids.append(row_iid)
        if not anchor_iid or anchor_iid not in row_ids:
            return "break"

        start = row_ids.index(anchor_iid)
        hubo_cambios = False
        undo_rows = []
        for r_off, row_data in enumerate(matrix):
            row_pos = start + r_off
            if row_pos >= len(row_ids):
                break
            target_iid = row_ids[row_pos]
            vals = list(tree.item(target_iid, "values"))
            before_vals = list(vals)
            while len(vals) < len(cols):
                vals.append("")
            while len(before_vals) < len(cols):
                before_vals.append("")
            row_changed = False
            for c_off, cell_raw in enumerate(row_data):
                col_idx = anchor_col + c_off
                if col_idx < 0 or col_idx >= len(cols):
                    continue
                vals[col_idx] = _normalizar_valor_edicion(col_idx, cell_raw)
                row_changed = True
            if row_changed and vals != before_vals:
                undo_rows.append((target_iid, before_vals))
                tree.item(target_iid, values=vals)
                hubo_cambios = True

        if hubo_cambios:
            push_undo_rows(undo_state, tree, undo_rows, meta={"seccion": "area_recaudacion"})
            _recalcular_grilla_editable()
        return "break"

    def _deshacer_celdas(_evt=None):
        nonlocal edit_entry
        if edit_entry is not None:
            _cerrar_editor(save=False)
        snapshot = pop_undo_snapshot(undo_state)
        if not snapshot:
            return "break"
        restore_undo_snapshot(tree, snapshot)
        _recalcular_grilla_editable()
        return "break"

    def _on_double_click(evt):
        nonlocal edit_entry
        _cerrar_editor(save=True)

        if tree.identify("region", evt.x, evt.y) != "cell":
            return

        col = tree.identify_column(evt.x)
        iid = tree.identify_row(evt.y)
        if not iid:
            return

        col_idx = int(col.replace("#", "")) - 1
        if col_idx < 0 or col_idx > 15:
            return

        clipboard_state["cell"] = (iid, col_idx)

        row_vals = list(tree.item(iid, "values"))
        if row_vals and str(row_vals[0]).strip().lower() == "totales":
            return

        x, y, w, h = tree.bbox(iid, col)
        if w <= 0 or h <= 0:
            return

        _clear_diff_labels(w_tmp)
        val = list(tree.item(iid, "values"))[col_idx]
        justify = "center" if col_idx == 0 else "right"
        edit_entry = ttk.Entry(tree, justify=justify)
        edit_entry.place(x=x, y=y, width=w, height=h)
        edit_entry.insert(0, val)
        edit_entry.focus_set()
        edit_entry._cell = (iid, col)
        edit_entry.bind("<Return>", lambda _e: _cerrar_editor(save=True))
        edit_entry.bind("<Escape>", lambda _e: _cerrar_editor(save=False))
        edit_entry.bind("<FocusOut>", lambda _e: _cerrar_editor(save=True))

    w_tmp = PlanillaWidgets(
        juego=juego,
        codigo_juego=-1,
        header_canvas=header_canvas,
        header2_canvas=header2_canvas,
        filter_canvas=filter_canvas,
        data_tree=tree,
        cols=cols,
        semanas={},
        diff_labels=[],
    )
    w_tmp.filter_vars = filter_vars
    w_tmp.total_iid = _ensure_total_row_iid(w_tmp)
    w_tmp.all_item_ids = list(tree.get_children())
    w_tmp.request_diff_render = _request_diff_render
    w_tmp.hide_diff_labels = _hide_diff_labels
    w_tmp.close_editor = _cerrar_editor

    def _on_filter_change(*_args):
        _aplicar_filtros_planilla(w_tmp)
        _actualizar_textos_header_filtros()

    for var in filter_vars.values():
        var.trace_add("write", _on_filter_change)
    _aplicar_estilo_filas()
    tree.bind("<Double-1>", _on_double_click)
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")

    header2_canvas.bind("<Button-1>", _on_header2_click)

    _apply_layout_sync()
    _actualizar_textos_header_filtros()
    _request_diff_render()

    return w_tmp


PANEL_BG = "#D3E6F7"
BADGE_BG = "#F7FAFE"
BADGE_BORDER = "#C7D2E0"
BADGE_TEXT = "#425466"


def _ensure_area_toolbar_style(parent):
    style = ttk.Style(parent)
    filtro_style = "Planilla.Toolbar.TCombobox"
    style.configure(
        filtro_style,
        padding=(9, 5),
        fieldbackground="#FBFDFF",
        background="#FBFDFF",
        foreground="#0F172A",
        bordercolor="#9FB6CD",
        lightcolor="#B9CCE0",
        darkcolor="#9FB6CD",
        arrowcolor="#1A3A6B",
    )
    style.map(
        filtro_style,
        background=[("active", "#EEF4FC"), ("!active", "#FBFDFF")],
        fieldbackground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF"), ("disabled", "#E9EFF6")],
        foreground=[("readonly", "#0F172A"), ("disabled", "#94A3B8")],
        bordercolor=[("focus", "#5A7EA5"), ("!focus", "#9FB6CD")],
        lightcolor=[("focus", "#5A7EA5"), ("!focus", "#B9CCE0")],
        darkcolor=[("focus", "#5A7EA5"), ("!focus", "#9FB6CD")],
        arrowcolor=[("active", "#235291"), ("!active", "#1A3A6B")],
    )
    return filtro_style


def _rounded_rect(canvas: tk.Canvas, x1, y1, x2, y2, r=10, **kwargs):
    points = [
        x1 + r, y1,
        x1 + r, y1,
        x2 - r, y1,
        x2 - r, y1,
        x2, y1,
        x2, y1 + r,
        x2, y1 + r,
        x2, y2 - r,
        x2, y2 - r,
        x2, y2,
        x2 - r, y2,
        x2 - r, y2,
        x1 + r, y2,
        x1 + r, y2,
        x1, y2,
        x1, y2 - r,
        x1, y2 - r,
        x1, y1 + r,
        x1, y1 + r,
        x1, y1,
    ]
    return canvas.create_polygon(points, smooth=True, **kwargs)


def _crear_badge_redondeado(parent, textvariable: tk.StringVar, *, width=250, height=28):
    canvas = tk.Canvas(
        parent,
        width=width,
        height=height,
        bg=PANEL_BG,
        highlightthickness=0,
        bd=0,
        relief="flat",
    )

    font = tkfont.Font(family="Segoe UI", size=9)

    def _redraw(*_args):
        texto = str(textvariable.get() or "").strip()
        canvas.delete("all")

        if not texto:
            return

        pad_x = 14
        text_w = font.measure(texto)
        badge_w = max(175, min(width, text_w + pad_x * 2))
        badge_h = height - 4

        x1 = 2
        y1 = 2
        x2 = x1 + badge_w
        y2 = y1 + badge_h

        _rounded_rect(
            canvas,
            x1,
            y1,
            x2,
            y2,
            r=10,
            fill=BADGE_BG,
            outline=BADGE_BORDER,
            width=1,
        )

        canvas.create_text(
            x1 + pad_x,
            height // 2,
            text=texto,
            anchor="w",
            font=font,
            fill=BADGE_TEXT,
        )

    textvariable.trace_add("write", _redraw)
    _redraw()
    return canvas


# ======================
# UI BUILDER SECCIÓN AREA RECAUDACIÓN
# ======================

def build_area_recaudacion(fr_seccion: ttk.Frame, estado_var):
    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(1, weight=1)

    filtro_style = _ensure_area_toolbar_style(fr_seccion)
    juegos_nombres = [j for j, _ in JUEGOS]
    widgets_por_juego: dict[str, PlanillaWidgets] = {}
    built: set[str] = set()
    semanas_por_juego: dict[str, dict[int, list[int]]] = {}
    rangos_semana_por_juego: dict[str, dict[int, tuple[str, str]]] = {}
    semana_actual_por_juego: dict[str, int] = {}
    guardado_job = None

    top = ttk.Frame(fr_seccion, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 10))
    top.columnconfigure(5, weight=1)

    ttk.Label(top, text="Juego:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")

    combo = ttk.Combobox(top, state="readonly", values=juegos_nombres, width=30, style=filtro_style)
    combo.grid(row=0, column=1, sticky="w", padx=(8, 10))

    ttk.Label(top, text="Semana:", style="PanelLabel.TLabel").grid(row=0, column=2, sticky="e", padx=(12, 6))
    semana_var = tk.StringVar()
    combo_semana = ttk.Combobox(top, state="readonly", textvariable=semana_var, values=[], width=28, style=filtro_style)
    combo_semana.grid(row=0, column=3, sticky="w")

    rango_semana_var = tk.StringVar(value="")
    # Del/Al se muestra directamente en el filtro de semana.
    badge_rango_semana = ttk.Frame(top, width=1, height=1)
    desde_var = tk.StringVar(value="")
    hasta_var = tk.StringVar(value="")


    # Compatibilidad defensiva: elimina botones legados que no deben mostrarse
    # en esta UI (acciones migradas al flujo inline o al menú "Voy a...").
    textos_boton_ocultos = (
        "ultima semana",
        "última semana",
        "guardar del/al",
        "editar del/al",
    )
    for _child in top.winfo_children():
        try:
            if isinstance(_child, ttk.Button):
                _txt = str(_child.cget("text") or "").strip().lower()
                if any(txt in _txt for txt in textos_boton_ocultos):
                    _child.destroy()
        except Exception:
            pass

    def _programar_guardado_area():
        nonlocal guardado_job

        def _ejecutar_guardado():
            nonlocal guardado_job
            guardado_job = None
            guardar_area_recaudacion(widgets_por_juego, estado_var=None)

        if guardado_job is not None:
            try:
                top.after_cancel(guardado_job)
            except Exception:
                pass
        guardado_job = top.after(120, _ejecutar_guardado)

    def _actualizar_label_rango_semana(juego: str, semana_txt: str):
        semana_txt = _semana_interna(semana_txt)
        try:
            semana_n = int(str(semana_txt).strip().lower().replace("semana", "").strip())
        except Exception:
            rango_semana_var.set("")
            desde_var.set("")
            hasta_var.set("")
            return

        def _obtener_rango_semana(juego_nom: str, semana_num: int) -> tuple[str, str] | None:
            rangos_locales = rangos_semana_por_juego.get(juego_nom, {})
            desde_hasta_local = rangos_locales.get(semana_num)
            if isinstance(desde_hasta_local, tuple) and len(desde_hasta_local) == 2:
                desde_l = str(desde_hasta_local[0] or "").strip()
                hasta_l = str(desde_hasta_local[1] or "").strip()
                if desde_l and hasta_l:
                    return desde_l, hasta_l

            if hasattr(app_state, "planilla_rangos_semana_global"):
                try:
                    global_raw = app_state.planilla_rangos_semana_global.get(semana_num)
                except Exception:
                    global_raw = None

                if isinstance(global_raw, dict):
                    desde_g = str(global_raw.get("desde", "") or "").strip()
                    hasta_g = str(global_raw.get("hasta", "") or "").strip()
                elif isinstance(global_raw, (tuple, list)) and len(global_raw) >= 2:
                    desde_g = str(global_raw[0] or "").strip()
                    hasta_g = str(global_raw[1] or "").strip()
                else:
                    desde_g = ""
                    hasta_g = ""

                if desde_g and hasta_g:
                    return desde_g, hasta_g

            return None

        desde_hasta = _obtener_rango_semana(juego, semana_n)
        if not desde_hasta:
            rango_semana_var.set("Del: --/--/---- al: --/--/----")
            desde_var.set("")
            hasta_var.set("")
            return

        desde, hasta = desde_hasta
        desde_var.set(desde)
        hasta_var.set(hasta)
        rango_semana_var.set(f"Del: {desde} al: {hasta}")

    def _mostrar_edicion_del_al():
        juego_actual = combo.get().strip()
        sem_txt = semana_var.get().strip()
        if not juego_actual or not sem_txt:
            messagebox.showwarning("Semana", "Seleccioná juego y semana para editar el rango.")
            return

        win = tk.Toplevel(top)
        win.title("Editar rango semanal")
        win.transient(top.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text="Del:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w", padx=(10, 6), pady=(10, 6))
        ent_desde = ttk.Entry(win, width=14)
        ent_desde.grid(row=0, column=1, sticky="w", padx=(0, 10), pady=(10, 6))
        ent_desde.insert(0, str(desde_var.get() or "").strip())

        ttk.Label(win, text="Al:", style="PanelLabel.TLabel").grid(row=1, column=0, sticky="w", padx=(10, 6), pady=(0, 10))
        ent_hasta = ttk.Entry(win, width=14)
        ent_hasta.grid(row=1, column=1, sticky="w", padx=(0, 10), pady=(0, 10))
        ent_hasta.insert(0, str(hasta_var.get() or "").strip())

        def _guardar_y_cerrar():
            desde_var.set(str(ent_desde.get() or "").strip())
            hasta_var.set(str(ent_hasta.get() or "").strip())
            win.destroy()
            _guardar_rango_semana_manual()

        ttk.Button(win, text="Guardar", command=_guardar_y_cerrar, style="Marino.TButton").grid(row=2, column=0, padx=(10, 6), pady=(0, 10), sticky="w")
        ttk.Button(win, text="Cancelar", command=win.destroy).grid(row=2, column=1, padx=(0, 10), pady=(0, 10), sticky="e")

        ent_desde.bind("<Return>", lambda _e: (ent_hasta.focus_set(), "break")[1])
        ent_hasta.bind("<Return>", lambda _e: (_guardar_y_cerrar(), "break")[1])
        ent_desde.focus_set()
        ent_desde.icursor("end")

    # Caja Del/Al removida visualmente.

    def _guardar_rango_semana_manual():
        juego_actual = combo.get().strip()
        sem_txt = semana_var.get().strip()
        if not juego_actual or not sem_txt:
            messagebox.showwarning("Semana", "Seleccioná juego y semana para guardar el rango.")
            return

        try:
            semana_n = int(str(sem_txt).strip().lower().replace("semana", "").strip())
        except Exception:
            messagebox.showwarning("Semana", f"No pude interpretar la semana seleccionada: {sem_txt}")
            return

        desde = str(desde_var.get() or "").strip()
        hasta = str(hasta_var.get() or "").strip()
        if not desde or not hasta:
            messagebox.showwarning("Semana", "Completá manualmente 'Del' y 'Al'.")
            return

        fecha_desde = _parse_fecha_sorteo(desde)
        fecha_hasta = _parse_fecha_sorteo(hasta)
        if fecha_desde is None or fecha_hasta is None:
            messagebox.showwarning("Semana", "Las fechas 'Del' y 'Al' deben tener formato válido (dd/mm/aaaa).")
            return
        if fecha_desde > fecha_hasta:
            messagebox.showwarning("Semana", "La fecha 'Del' no puede ser mayor que la fecha 'Al'.")
            return

        desde = fecha_desde.strftime("%d/%m/%Y")
        hasta = fecha_hasta.strftime("%d/%m/%Y")

        rangos = rangos_semana_por_juego.setdefault(juego_actual, {})
        rangos[semana_n] = (desde, hasta)
        if hasattr(app_state, "planilla_rangos_semana_global"):
            app_state.planilla_rangos_semana_global[semana_n] = (desde, hasta)

        w_local = widgets_por_juego.get(juego_actual)
        if w_local is not None:
            w_local.rangos_semana = dict(rangos)

        _actualizar_label_rango_semana(juego_actual, sem_txt)
        if hasattr(app_state, "publicar_filtro_area_recaudacion"):
            app_state.publicar_filtro_area_recaudacion(juego_actual, semana_n, desde, hasta)
        guardar_area_recaudacion(widgets_por_juego, estado_var=None)
        if estado_var is not None:
            estado_var.set(f"{juego_actual}: semana {semana_n} actualizada manualmente (Del {desde} al {hasta}).")

    stack = ttk.Frame(fr_seccion)
    stack.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
    stack.columnconfigure(0, weight=1)
    stack.rowconfigure(0, weight=1)

    frames = {}

    def _sincronizar_semanas_desde_storage(juego: str):
        data_juego = _leer_planilla_juego_desde_storage(juego)

        semanas_guardadas = _extraer_semanas_guardadas(data_juego)

        # Blindaje: si el storage viene sin mapa (o vacío) NO colapsar si ya
        # tenemos semanas en memoria para ese juego.
        semanas_mem = semanas_por_juego.get(juego, {})
        if (not semanas_guardadas) and isinstance(semanas_mem, dict) and semanas_mem:
            semanas_guardadas = dict(semanas_mem)

        if not semanas_guardadas:
            # Compatibilidad con guardados viejos: si no vino el mapa de semanas,
            # lo reconstruimos desde los sorteos de filas para no dejar el filtro vacío.
            filas_guardadas = data_juego.get("filas", []) if isinstance(data_juego, dict) else []
            semanas_guardadas = _reconstruir_semanas_desde_filas(filas_guardadas)

        if semanas_guardadas:
            semanas_por_juego[juego] = semanas_guardadas

        rangos_guardados = _extraer_rangos_semana_guardados(data_juego)

        rangos_mem = rangos_semana_por_juego.get(juego, {})
        if (not rangos_guardados) and isinstance(rangos_mem, dict) and rangos_mem:
            rangos_guardados = dict(rangos_mem)

        if rangos_guardados:
            rangos_semana_por_juego[juego] = rangos_guardados
            if hasattr(app_state, "planilla_rangos_semana_global"):
                for sem, rango in rangos_guardados.items():
                    try:
                        sem_n = int(sem)
                    except Exception:
                        continue
                    if sem_n < 1:
                        continue
                    if isinstance(rango, tuple) and len(rango) == 2:
                        desde = str(rango[0] or "").strip()
                        hasta = str(rango[1] or "").strip()
                        if desde and hasta:
                            app_state.planilla_rangos_semana_global[sem_n] = (desde, hasta)

        try:
            sem_act = int(data_juego.get("semana_actual", 0) or 0)
        except Exception:
            sem_act = 0

        sem_mem = semana_actual_por_juego.get(juego, 0)
        if sem_act <= 0 and isinstance(sem_mem, int) and sem_mem > 0:
            sem_act = sem_mem

        if sem_act > 0:
            semana_actual_por_juego[juego] = sem_act

        w_local = widgets_por_juego.get(juego)
        if w_local is not None:
            w_local.semanas = dict(semanas_por_juego.get(juego, {}))
            w_local.rangos_semana = dict(rangos_semana_por_juego.get(juego, {}))
            w_local.semana_actual = int(semana_actual_por_juego.get(juego, 0) or 0)
        return data_juego, semanas_guardadas

    def _aplicar_guardado_a_juego(juego: str, w: PlanillaWidgets):
        _recargar_juego_desde_storage(juego, w, estado_var)
        _sincronizar_semanas_desde_storage(juego)


    def _ajustar_ancho_combo_semana(valores: list[str] | None = None, texto_actual: str = ""):
        candidatos = [str(v or "").strip() for v in (valores or [])]
        if texto_actual:
            candidatos.append(str(texto_actual).strip())
        largo = max([len(x) for x in candidatos if x] or [34])
        ancho = max(34, min(42, largo + 3))
        try:
            combo_semana.configure(width=ancho)
        except Exception:
            pass

    def _actualizar_combo_semana(juego: str, semana_preferida: int | None = None):
        semanas = semanas_por_juego.get(juego, {})
        numeros = sorted(int(n) for n in semanas.keys())

        if not numeros:
            combo_semana.configure(values=[])
            _ajustar_ancho_combo_semana([], "")
            semana_var.set("")
            rango_semana_var.set("")
            try:
                combo_semana.set("")
            except Exception:
                pass
            combo_semana.state(["disabled"])
            return

        etiquetas = _combo_values_semanas_desde_numeros(numeros)
        combo_semana.configure(values=etiquetas)
        combo_semana.state(["!disabled"])

        if semana_preferida is None:
            # Regla de negocio:
            # al cambiar de juego, se debe respetar primero la semana global
            # actualmente elegida para toda la Planilla. Solo si ese juego no
            # tiene esa semana, caer a la memoria propia del juego.
            semana_global = 0
            try:
                payload_global = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
                semana_global = int(payload_global.get("semana", 0) or 0)
            except Exception:
                semana_global = 0

            if semana_global in numeros:
                semana_preferida = semana_global
            else:
                semana_preferida = int(semana_actual_por_juego.get(juego, 0) or 0)

        if semana_preferida not in numeros:
            semana_preferida = numeros[0]

        # Sincronizar también la memoria por juego con la semana elegida,
        # para que todas las secciones y juegos queden alineados al navegar.
        try:
            semana_actual_por_juego[juego] = int(semana_preferida)
        except Exception:
            pass

        visible = _semana_visible(f"Semana {semana_preferida}")
        semana_var.set(visible)
        _ajustar_ancho_combo_semana(etiquetas, visible)
        try:
            combo_semana.set(visible)
        except Exception:
            pass
        _actualizar_label_rango_semana(juego, visible)


    def _resolver_numero_semana_desde_filtro(juego_actual: str, sem_txt: str, w: PlanillaWidgets | None) -> int:
        texto = str(sem_txt or "").strip()
        if not texto:
            raise ValueError("Filtro de semana vacío")

        interna = _semana_interna(texto)
        m = re.fullmatch(r"(?i)\s*semana\s*(\d+)\s*", str(interna or "").strip())
        if m:
            return max(1, min(5, int(m.group(1))))

        # Caso del video:
        # el combo muestra "Del: ... al: ..." pero internamente todavía no hay
        # un helper que lo traduzca a "Semana N". Si el juego tiene una única
        # semana cargada, esa es inequívocamente la semana a mostrar.
        semanas_juego = semanas_por_juego.get(juego_actual, {}) or {}
        numeros = sorted(
            int(k) for k in semanas_juego.keys()
            if str(k).strip()
        )

        if len(numeros) == 1:
            return numeros[0]

        for candidato in (
            int(getattr(w, "semana_actual", 0) or 0) if w is not None else 0,
            int(semana_actual_por_juego.get(juego_actual, 0) or 0),
        ):
            if candidato > 0:
                return candidato

        raise ValueError(f"No pude resolver la semana desde el filtro: {texto}")

    def _aplicar_filtro_semana_actual(show_warning: bool = True, publish_global: bool = True):
        juego_actual = combo.get().strip()
        if not juego_actual or juego_actual not in built:
            return

        w = widgets_por_juego.get(juego_actual)
        if w is None:
            return

        def _limpiar_grilla_visible():
            _aplicar_filas_a_widget(w, [], clear_rows=True)
            w._ultimo_juego_cargado = juego_actual
            w._ultima_semana_cargada = 0
            w._ultimos_sorteos_cargados = []
            refresh_planilla(w, estado_var=None)
            try:
                w.request_diff_render()
            except Exception:
                _render_diff_labels(w)

        sem_txt = semana_var.get().strip()
        if not sem_txt:
            _limpiar_grilla_visible()
            return

        try:
            semana_nro = _resolver_numero_semana_desde_filtro(juego_actual, sem_txt, w)
        except Exception:
            _limpiar_grilla_visible()
            if show_warning:
                messagebox.showwarning("Semana", f"No pude interpretar la semana seleccionada: {sem_txt}")
            return

        # Persistir inmediatamente la grilla visible ANTES de recargar otra semana.
        # Si se deja solo el guardado diferido (after 120ms), al cambiar de semana
        # la planilla se sobreescribe y se pierden ediciones manuales recientes.
        guardar_area_recaudacion(widgets_por_juego, estado_var=None)

        # Mantener la misma semana para TODOS los juegos de Área Recaudación.
        for _juego_sync in list(semanas_por_juego.keys()):
            try:
                semana_actual_por_juego[_juego_sync] = semana_nro
            except Exception:
                pass
        semana_actual_por_juego[juego_actual] = semana_nro
        if w is not None:
            w.semana_actual = semana_nro
            try:
                w.hide_diff_labels()
            except Exception:
                _clear_diff_labels(w)
        _programar_guardado_area()

        _actualizar_label_rango_semana(juego_actual, sem_txt)

        desde_filtro = ""
        hasta_filtro = ""
        try:
            rango_local = rangos_semana_por_juego.get(juego_actual, {}).get(semana_nro, ("", ""))
            if isinstance(rango_local, tuple) and len(rango_local) == 2:
                desde_filtro = str(rango_local[0] or "").strip()
                hasta_filtro = str(rango_local[1] or "").strip()
        except Exception:
            pass

        if (not desde_filtro or not hasta_filtro) and hasattr(app_state, "planilla_rangos_semana_global"):
            try:
                rango_global = app_state.planilla_rangos_semana_global.get(semana_nro)
            except Exception:
                rango_global = None
            if isinstance(rango_global, dict):
                desde_filtro = desde_filtro or str(rango_global.get("desde", "") or "").strip()
                hasta_filtro = hasta_filtro or str(rango_global.get("hasta", "") or "").strip()
            elif isinstance(rango_global, (tuple, list)) and len(rango_global) >= 2:
                desde_filtro = desde_filtro or str(rango_global[0] or "").strip()
                hasta_filtro = hasta_filtro or str(rango_global[1] or "").strip()
        if hasattr(app_state, "publicar_filtro_area_recaudacion"):
            app_state.publicar_filtro_area_recaudacion(juego_actual, semana_nro, desde_filtro, hasta_filtro)

        semanas = semanas_por_juego.get(juego_actual, {})
        sorteos = semanas.get(semana_nro, [])
        if not sorteos:
            _limpiar_grilla_visible()
            if show_warning:
                messagebox.showwarning("Semana", f"No hay sorteos para {juego_actual} en la semana {semana_nro}.")
            if estado_var is not None:
                estado_var.set(f"{juego_actual}: semana {semana_nro} sin sorteos, grilla limpiada.")
            return
        sorteos_normalizados = []
        for s in sorteos:
            try:
                sorteos_normalizados.append(int(s))
            except Exception:
                continue

        if (
            getattr(w, "_ultimo_juego_cargado", "") == juego_actual
            and int(getattr(w, "_ultima_semana_cargada", 0) or 0) == semana_nro
            and getattr(w, "_ultimos_sorteos_cargados", []) == sorteos_normalizados
        ):
            # Recalcular siempre la fila Totales y estilos, incluso si la semana
            # no cambió. Evita que quede desactualizada/oculta hasta forzar cambio.
            refresh_planilla(w, estado_var=None)
            try:
                w.request_diff_render()
            except Exception:
                _render_diff_labels(w)
            return

        cargar_sorteos_en_planilla(
            w,
            sorteos,
            estado_var,
            sorteo_preferido=(sorteos_normalizados[0] if sorteos_normalizados else None),
        )
        w._ultimo_juego_cargado = juego_actual
        w._ultima_semana_cargada = semana_nro
        w._ultimos_sorteos_cargados = sorteos_normalizados
        try:
            w.request_diff_render()
        except Exception:
            _render_diff_labels(w)
        if estado_var is not None:
            estado_var.set(f"{juego_actual}: semana {semana_nro} cargada con {len(sorteos)} sorteos.")

    combo_semana.bind("<<ComboboxSelected>>", lambda _e: _aplicar_filtro_semana_actual(show_warning=False))

    def build_juego(juego: str, fr_juego: ttk.Frame):
        fr_juego.columnconfigure(0, weight=1)
        fr_juego.rowconfigure(0, weight=1)

        codigo_juego = None
        for j, cj in JUEGOS:
            if j == juego:
                codigo_juego = cj
                break

        def importar_pju():
            # Si el usuario estaba editando una celda (Entry superpuesto),
            # confirmar ese valor antes de calcular continuidad manual+PJU.
            cerrar_editor = getattr(w, "close_editor", None)
            if callable(cerrar_editor):
                try:
                    cerrar_editor(save=True)
                except Exception:
                    pass

            path = filedialog.askopenfilename(
                title=f"Seleccionar PJU (TXT) - {juego}",
                filetypes=[("TXT", "*.txt"), ("Todos", "*.*")],
            )
            if not path:
                return

            try:
                data = leer_json_desde_txt(path)
            except Exception as e:
                messagebox.showerror("PJU", f"No pude extraer JSON desde el TXT\n{e}")
                return

            if codigo_juego is None:
                messagebox.showerror("PJU", f"No tengo codigo_juego para {juego}.")
                return

            semanas, ultima_semana, lunes_semana_1 = extraer_sorteos_por_semanas(data, codigo_juego)
            if not semanas:
                sorteos = extraer_sorteos_por_codigo(data, codigo_juego)
                if not sorteos:
                    messagebox.showwarning("PJU", f"Importaste cualquier cosa menos {juego}.")
                    return
                semanas = {1: sorteos}
                ultima_semana = 1
                lunes_semana_1 = None

            rangos_importados = _calcular_rangos_semana_desde_lunes(semanas, lunes_semana_1)
            semanas, rangos_importados = _alinear_semanas_importadas_con_rangos_existentes(
                semanas,
                rangos_importados,
                rangos_semana_por_juego.get(juego, {}),
            )

            # Propaga inmediatamente Del/Al de todas las semanas importadas
            # para que Prescripciones, Anticipos/Topes y Totales ya tengan
            # los rangos listos aunque todavía no se haya cambiado el filtro.
            if isinstance(rangos_importados, dict) and hasattr(app_state, "planilla_rangos_semana_global"):
                for sem, rango in rangos_importados.items():
                    try:
                        sem_n = int(sem)
                    except Exception:
                        continue
                    if sem_n < 1:
                        continue
                    if not isinstance(rango, (tuple, list)) or len(rango) < 2:
                        continue
                    desde = str(rango[0] or "").strip()
                    hasta = str(rango[1] or "").strip()
                    if desde and hasta:
                        app_state.planilla_rangos_semana_global[sem_n] = (desde, hasta)

            # Sincroniza sorteos manuales visibles de la semana activa ANTES de
            # mergear el PJU. Esto preserva los sorteos creados a mano y evita
            # que el append del PJU termine pisando el mapa semanal vigente.
            semanas_base = _inyectar_sorteos_visibles_en_semana_actual(
                semanas_por_juego.get(juego) or getattr(w, "semanas", {}) or {},
                w,
            )

            if _append_pju_si_continua_desde_ultimo_sorteo_visible(w, semanas, estado_var):
                semanas_ancladas = _anclar_sorteos_importados_a_semana_1(semanas)
                semanas_por_juego[juego] = _mergear_semanas_sin_pisar(semanas_base, semanas_ancladas)
                rangos_semana_por_juego[juego] = _mergear_rangos_semana(
                    rangos_semana_por_juego.get(juego),
                    rangos_importados,
                )
                w.semanas = dict(semanas_por_juego.get(juego, {}))
                w.rangos_semana = dict(rangos_semana_por_juego.get(juego, {}))
                semana_actual_por_juego[juego] = 1
                w.semana_actual = 1
                # Persistencia inmediata: al reabrir la app debe conservar Del/Al.
                guardar_area_recaudacion(widgets_por_juego, estado_var=None)
                _actualizar_combo_semana(juego, semana_preferida=1)
                if combo.get().strip() == juego:
                    _aplicar_filtro_semana_actual(show_warning=False, publish_global=False)
                return

            semanas_por_juego[juego] = _mergear_semanas_sin_pisar(semanas_base, semanas)
            rangos_semana_por_juego[juego] = _mergear_rangos_semana(rangos_semana_por_juego.get(juego), rangos_importados)
            w.semanas = dict(semanas_por_juego.get(juego, {}))
            w.rangos_semana = dict(rangos_semana_por_juego.get(juego, {}))
            # Persistencia inmediata: al reabrir la app debe conservar Del/Al.
            guardar_area_recaudacion(widgets_por_juego, estado_var=None)

            data_juego_guardado = _leer_planilla_juego_desde_storage(juego)
            modo_semana_nueva = bool(data_juego_guardado.get("modo_semana_nueva"))
            semana_exportada = int(data_juego_guardado.get("semana_exportada", 0) or 0)
            filas_template_semana_1 = data_juego_guardado.get("filas_template_semana_1", [])

            if modo_semana_nueva and semana_exportada == 1 and isinstance(filas_template_semana_1, list):
                sorteos_semana_1 = semanas.get(1, [])
                if sorteos_semana_1:
                    _actualizar_combo_semana(juego, semana_preferida=1)
                    _aplicar_template_semana_1_a_sorteos(w, filas_template_semana_1, sorteos_semana_1, estado_var)
                    if estado_var is not None:
                        estado_var.set(f"{juego}: se aplicó el archivo cargado como Semana 1.")
                    return

            # Requisito funcional: al importar PJU siempre mantener foco en Semana 1.
            semana_destino = 1 if 1 in semanas else min(semanas.keys())

            if combo.get().strip() == juego:
                _actualizar_combo_semana(juego, semana_preferida=semana_destino)
                sorteos_semana = semanas.get(semana_destino, [])
                cargar_sorteos_en_planilla(
                    w,
                    sorteos_semana,
                    estado_var,
                    append_if_contiguous=False,
                )
                _aplicar_filtro_semana_actual(show_warning=False)
            else:
                _actualizar_combo_semana(combo.get().strip())

        w = crear_planilla_excel_like(
            fr_juego,
            juego=juego,
            filas=120,
            estado_var=estado_var,
            on_import_pju=importar_pju,
        )
        if codigo_juego is not None:
            w.codigo_juego = codigo_juego
        w.on_grid_changed = _programar_guardado_area
        widgets_por_juego[juego] = w

        if not hasattr(app_state, "planilla_refresh_hooks"):
            app_state.planilla_refresh_hooks = {}
        app_state.planilla_refresh_hooks[juego] = (lambda ww=w: lambda: refresh_planilla(ww, estado_var))()

        if not hasattr(app_state, "planilla_area_reload_hooks"):
            app_state.planilla_area_reload_hooks = {}

        def _reload_hook(jj=juego, ww=w):
            try:
                ww.hide_diff_labels()
            except Exception:
                _clear_diff_labels(ww)
            _recargar_juego_desde_storage(jj, ww, estado_var)
            _sincronizar_semanas_desde_storage(jj)
            if combo.get().strip() == jj:
                _actualizar_combo_semana(jj)
                _aplicar_filtro_semana_actual(show_warning=False)
                try:
                    ww.request_diff_render()
                except Exception:
                    _render_diff_labels(ww)

        app_state.planilla_area_reload_hooks[juego] = _reload_hook

        if not hasattr(app_state, "planilla_area_snapshot_hooks"):
            app_state.planilla_area_snapshot_hooks = {}
        app_state.planilla_area_snapshot_hooks[juego] = (lambda ww=w: lambda: _snapshot_juego(ww))()

        _aplicar_guardado_a_juego(juego, w)
        w.semanas = dict(semanas_por_juego.get(juego, {}))

    for j in juegos_nombres:
        fr = ttk.Frame(stack)
        fr.grid(row=0, column=0, sticky="nsew")
        fr.columnconfigure(0, weight=1)
        fr.rowconfigure(0, weight=1)
        frames[j] = fr

    _show_job = {"id": None, "juego": ""}
    _juego_visible = {"nombre": ""}

    def _cancelar_guardado_diferido():
        nonlocal guardado_job
        if guardado_job is not None:
            try:
                top.after_cancel(guardado_job)
            except Exception:
                pass
            guardado_job = None

    def _persistir_juego_si_corresponde(juego_nom: str):
        juego_src = (juego_nom or "").strip()
        if not juego_src:
            return
        w_src = widgets_por_juego.get(juego_src)
        if w_src is None:
            return

        _cancelar_guardado_diferido()

        cerrar_editor = getattr(w_src, "close_editor", None)
        if callable(cerrar_editor):
            try:
                cerrar_editor(save=True)
            except Exception:
                pass

        # Base: lo que ya está guardado + lo que está en memoria (sin mover sorteos).
        data_prev = _leer_planilla_juego_desde_storage(juego_src)
        semanas_prev = _extraer_semanas_guardadas(data_prev)
        semanas_mem = semanas_por_juego.get(juego_src) or getattr(w_src, "semanas", {}) or {}
        semanas_base = _mergear_semanas_sin_pisar(semanas_prev, semanas_mem)

        semanas_sync = _inyectar_sorteos_visibles_en_semana_actual(semanas_base, w_src)
        semanas_sync = _mergear_semanas_sin_pisar(semanas_prev, semanas_sync)

        semanas_por_juego[juego_src] = dict(semanas_sync)
        w_src.semanas = dict(semanas_sync)

        # Rangos: nunca perder Del/Al ya guardados.
        rangos_prev = _extraer_rangos_semana_guardados(data_prev)
        rangos_widget = _normalizar_rangos_semana(getattr(w_src, "rangos_semana", {}) or {})
        rangos_sync = dict(rangos_prev)
        rangos_sync.update(rangos_widget)
        rangos_semana_por_juego[juego_src] = dict(rangos_sync)
        w_src.rangos_semana = dict(rangos_sync)

        try:
            sem_act = int(getattr(w_src, "semana_actual", 0) or 0)
        except Exception:
            sem_act = 0
        if sem_act <= 0:
            try:
                sem_act = int(data_prev.get("semana_actual", 0) or 0)
            except Exception:
                sem_act = 0
        if sem_act > 0:
            semana_actual_por_juego[juego_src] = sem_act
        else:
            try:
                semanas_existentes = sorted(int(k) for k in semanas_sync.keys())
            except Exception:
                semanas_existentes = []
            if len(semanas_existentes) == 1:
                sem_act = semanas_existentes[0]
                try:
                    w_src.semana_actual = sem_act
                except Exception:
                    pass
                semana_actual_por_juego[juego_src] = sem_act

        guardar_area_recaudacion(widgets_por_juego, estado_var=None)

    def _refresh_visual_area_recaudacion(juego_actual: str | None = None):
        juego_target = (juego_actual or combo.get() or "").strip()
        if not juego_target:
            return

        # Refresh visual puro: NO leer storage ni resincronizar desde disco.
        # Acá solo se usa el estado ya cargado en memoria para que el cambio de
        # juego/retorno a la sección sea inmediato para el usuario final.
        w_activo = widgets_por_juego.get(juego_target)
        if w_activo is not None:
            try:
                w_activo.hide_diff_labels()
            except Exception:
                _clear_diff_labels(w_activo)

        _actualizar_combo_semana(juego_target)
        _aplicar_filtro_semana_actual(show_warning=False)
        if w_activo is not None:
            try:
                fr_seccion.after_idle(w_activo.request_diff_render)
            except Exception:
                try:
                    w_activo.request_diff_render()
                except Exception:
                    _render_diff_labels(w_activo)

    def mostrar(juego):
        if juego not in built:
            build_juego(juego, frames[juego])
            built.add(juego)
        frames[juego].tkraise()
        _juego_visible["nombre"] = juego
        _refresh_visual_area_recaudacion(juego)

    def _guardar_planilla_nueva_desde_ultima_semana():
        for nombre_juego in juegos_nombres:
            if nombre_juego not in built:
                build_juego(nombre_juego, frames[nombre_juego])
                built.add(nombre_juego)
        exportar_ultima_semana_a_planilla_nueva(widgets_por_juego, semanas_por_juego, estado_var)

    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    def _refresh_visual_area_recaudacion_hook():
        try:
            _refresh_visual_area_recaudacion()
        except Exception:
            pass

    app_state.planilla_visual_refresh_hooks["Area Recaudación"] = _refresh_visual_area_recaudacion_hook

    app_state.planilla_pasarela_ultima_semana_hook = _guardar_planilla_nueva_desde_ultima_semana

    def _guardar_hook_area_recaudacion():
        juego_visible = (_juego_visible.get("nombre", "") or combo.get() or "").strip()
        if juego_visible:
            _persistir_juego_si_corresponde(juego_visible)
        else:
            _cancelar_guardado_diferido()
            guardar_area_recaudacion(widgets_por_juego, estado_var=None)

    app_state.planilla_area_guardar_hook = _guardar_hook_area_recaudacion

    def _reset_area_recaudacion():
        nonlocal guardado_job

        try:
            if guardado_job is not None:
                top.after_cancel(guardado_job)
        except Exception:
            pass
        guardado_job = None

        semanas_por_juego.clear()
        rangos_semana_por_juego.clear()
        semana_actual_por_juego.clear()

        try:
            combo_semana.configure(values=[])
            _ajustar_ancho_combo_semana([], "")
        except Exception:
            pass
        try:
            combo_semana.set("")
            combo_semana.state(["disabled"])
        except Exception:
            pass

        try:
            semana_var.set("")
        except Exception:
            pass
        try:
            rango_semana_var.set("")
        except Exception:
            pass
        try:
            desde_var.set("")
        except Exception:
            pass
        try:
            hasta_var.set("")
        except Exception:
            pass

        for juego_reset, w_reset in widgets_por_juego.items():
            try:
                w_reset.semanas = {}
            except Exception:
                pass
            try:
                w_reset.rangos_semana = {}
            except Exception:
                pass
            try:
                w_reset.semana_actual = 0
            except Exception:
                pass
            try:
                w_reset._ultimo_sorteo_editado_manual = None
            except Exception:
                pass

            try:
                _aplicar_filas_a_widget(w_reset, [], clear_rows=True)
            except Exception:
                pass

            try:
                w_reset._ultimo_juego_cargado = ""
                w_reset._ultima_semana_cargada = 0
                w_reset._ultimos_sorteos_cargados = []
            except Exception:
                pass

            try:
                refresh_planilla(w_reset, estado_var=None)
            except Exception:
                pass

            try:
                w_reset.request_diff_render()
            except Exception:
                try:
                    _render_diff_labels(w_reset)
                except Exception:
                    pass

        try:
            guardar_area_recaudacion(widgets_por_juego, estado_var=None)
        except Exception:
            pass

    if not hasattr(app_state, "planilla_global_reset_hooks"):
        app_state.planilla_global_reset_hooks = {}
    app_state.planilla_global_reset_hooks["area_recaudacion"] = _reset_area_recaudacion


    def _on_combo_juego_selected(_e=None):
        _persistir_juego_si_corresponde(_juego_visible.get("nombre", ""))
        mostrar(combo.get())

    combo.bind("<<ComboboxSelected>>", _on_combo_juego_selected)
    combo.set(juegos_nombres[0])
    mostrar(juegos_nombres[0])

    # Inicializa en segundo plano el resto de juegos para que queden
    # registrados sus hooks/snapshots desde el arranque y así Totales
    # pueda recalcular automáticamente sin requerir que el usuario abra
    # cada pestaña manualmente.
    for nombre_juego in juegos_nombres:
        if nombre_juego in built:
            continue
        build_juego(nombre_juego, frames[nombre_juego])
        built.add(nombre_juego)

    return widgets_por_juego

def _mergear_semanas_sin_pisar(
    existentes: dict[int, list[int]] | None,
    nuevas: dict[int, list[int]] | None,
) -> dict[int, list[int]]:
    """
    Merge estable y *no destructivo*.

    - Nunca achica el mapa (no borra semanas existentes).
    - Nunca mueve un sorteo de semana: si ya estaba asignado a una semana en
      'existentes', se respeta esa asignación.
    - Solo agrega sorteos NUEVOS (que no existían en ninguna semana previa).

    Esto blinda la segmentación del PJU y evita que la operatoria (pasar tickets /
    reporte / SFA) termine colapsando semanas.
    """
    prev = _normalizar_mapa_semanas(existentes or {})
    cur = _normalizar_mapa_semanas(nuevas or {})

    # sorteo -> semana asignada (prioridad: existentes)
    asignacion: dict[int, int] = {}
    out: dict[int, list[int]] = {sem: list(vals) for sem, vals in sorted(prev.items())}

    for sem, vals in sorted(prev.items()):
        for v in vals:
            try:
                s = int(v)
            except Exception:
                continue
            asignacion.setdefault(s, sem)

    # Asegurar que existan claves de semanas nuevas (aunque la lista quede vacía)
    for sem in sorted(cur.keys()):
        out.setdefault(sem, list(out.get(sem, [])))

    for sem, vals in sorted(cur.items()):
        bucket = out.setdefault(sem, [])
        bucket_set = {int(v) for v in bucket if str(v).strip()}
        for v in vals:
            try:
                s = int(v)
            except Exception:
                continue
            # Si ya existía en alguna semana previa, NO se mueve.
            if s in asignacion:
                continue
            if s in bucket_set:
                continue
            bucket.append(s)
            bucket_set.add(s)
            asignacion[s] = sem

    out_norm: dict[int, list[int]] = {}
    for sem, vals in sorted(out.items()):
        uniq = sorted({int(v) for v in vals if str(v).strip()})
        if uniq:
            out_norm[int(sem)] = uniq
    return out_norm
