from __future__ import annotations

import copy
import re
import tkinter as tk
from tkinter import ttk

import app_state
import app_state
from tabs.filter_combobox_style import apply_filter_combobox_style
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

SEMANAS = [f"Semana {i}" for i in range(1, 6)]

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
    if txt in SEMANAS:
        return txt
    m = re.fullmatch(r"(?i)\s*semana\s*(\d+)\s*", txt)
    if m:
        try:
            return f"Semana {max(1, min(5, int(m.group(1))))}"
        except Exception:
            pass
    return txt

def _combo_values_semanas() -> list[str]:
    visibles: list[str] = []
    rangos = getattr(app_state, "planilla_rangos_semana_global", {}) or {}
    if not isinstance(rangos, dict):
        return visibles

    for sem_n in range(1, 6):
        rango = rangos.get(sem_n)
        if rango is None:
            rango = rangos.get(str(sem_n))
        if isinstance(rango, dict):
            desde = str(rango.get("desde", "") or "").strip()
            hasta = str(rango.get("hasta", "") or "").strip()
        elif isinstance(rango, (list, tuple)) and len(rango) >= 2:
            desde = str(rango[0] or "").strip()
            hasta = str(rango[1] or "").strip()
        else:
            continue
        if not desde and not hasta:
            continue
        visible = _semana_visible(f"Semana {sem_n}")
        if not visible or "--/--/----" in visible:
            continue
        visibles.append(visible)
    return visibles


FILAS_TOTALES = [
    "Total ventas",
    "Total comisiones",
    "Total premios",
    "Total prescripcion",
    "Total anticipos",
    "Total topes",
    "Total facuni",
    "Total comision agencia amiga",
]


def _parse_importe(v) -> float:
    s = str(v or "").strip()
    if not s:
        return 0.0
    s = s.replace("$", "").replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _fmt_pesos(v: float) -> str:
    s = f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


_DIFF_ROW_WARN_MIN = 20.0
_DIFF_ROW_DANGER_MIN = 40.0
_DIFF_ROW_WARN_BG = "#FFF2CC"
_DIFF_ROW_DANGER_BG = "#FDE2E1"


def _clasificar_diferencia_fila_total(valor: float) -> str | None:
    numero = abs(float(valor or 0.0))
    if numero > _DIFF_ROW_DANGER_MIN:
        return "danger"
    if numero > _DIFF_ROW_WARN_MIN:
        return "warn"
    return None


def _norm_sorteo(valor) -> str:
    fn = getattr(app_state, "_normalizar_sorteo_clave", None)
    if callable(fn):
        return fn(valor)
    txt = str(valor).strip()
    if not txt:
        return "0"
    txt = txt.replace(",", ".")
    try:
        num = float(txt)
        if num.is_integer():
            return str(int(num))
    except Exception:
        pass
    txt = txt.lstrip("0")
    return txt or "0"


def _semana_a_numero(semana_sel: str) -> int:
    try:
        return int(str(semana_sel).replace("Semana", "").strip() or "1")
    except Exception:
        return 1


def _semanas_sorteos_por_juego(semana_sel: str) -> dict[str, set[str]]:
    """
    Devuelve, por juego, el conjunto de sorteos pertenecientes a la semana seleccionada.
    La fuente correcta es el mapa semanal del área recaudación.
    """
    out: dict[str, set[str]] = {}
    sem_n = _semana_a_numero(semana_sel)

    snapshots = getattr(app_state, "obtener_snapshots_area_recaudacion", lambda: {})() or {}

    for juego, snap in snapshots.items():
        semanas = snap.get("semanas", {}) if isinstance(snap, dict) else {}
        filas = snap.get("filas", []) if isinstance(snap, dict) else []
        if isinstance(semanas, dict):
            sorteos = semanas.get(str(sem_n), semanas.get(sem_n, []))
        else:
            sorteos = []

        sorteos_norm = {_norm_sorteo(s) for s in sorteos if str(s).strip()}

        # Fallback sólo cuando el mapa semanal viene vacío o incompleto.
        # No debemos reconstruir por cantidad de sorteos porque juegos como
        # Quiniela pueden tener más de 10 sorteos válidos por semana.
        if not sorteos_norm:
            sorteos_reconstruidos = _reconstruir_sorteos_semana_desde_filas(filas, sem_n)
            if sorteos_reconstruidos:
                sorteos_norm = sorteos_reconstruidos

        out[str(juego)] = sorteos_norm

    return out


def _reconstruir_sorteos_semana_desde_filas(filas: list, semana_n: int) -> set[str]:
    if not isinstance(filas, list) or semana_n < 1:
        return set()

    sorteos_unicos: list[int] = []
    vistos: set[int] = set()

    for fila in filas:
        if not isinstance(fila, list) or not fila:
            continue
        try:
            s = int(float(str(fila[0]).strip().replace(",", ".")))
        except Exception:
            continue
        if s in vistos:
            continue
        vistos.add(s)
        sorteos_unicos.append(s)

    if not sorteos_unicos:
        return set()

    # Fallback defensivo:
    # si no existe el mapa semanal confiable, evitar dividir por un bloque fijo
    # (ej. 7 sorteos) porque hay juegos con más sorteos por semana y eso
    # produce totales incorrectos. En ese escenario devolvemos todos los sorteos
    # conocidos del juego para no recortar datos válidos.
    return {_norm_sorteo(s) for s in sorteos_unicos}


def _sumar_sfa_resumen_por_semana(campo: str, semanas_sorteos: dict[str, set[str]]) -> float:
    """
    Suma desde app_state.sfa_resumen_por_juego filtrando por los sorteos de la semana activa.
    Esto evita depender de la grilla visible del área recaudación.
    """
    total = 0.0
    src = getattr(app_state, "sfa_resumen_por_juego", {}) or {}

    for juego, datos_juego in src.items():
        if not isinstance(datos_juego, dict):
            continue
        juego_key = str(juego)
        if juego_key not in semanas_sorteos:
            continue
        sorteos_permitidos = semanas_sorteos.get(juego_key, set())

        for sorteo, data in datos_juego.items():
            if _norm_sorteo(sorteo) not in sorteos_permitidos:
                continue
            if not isinstance(data, dict):
                continue
            total += _obtener_valor_resumen(data, campo)

    return total


def _sumar_reporte_resumen_por_semana(campo: str, semanas_sorteos: dict[str, set[str]]) -> float:
    total = 0.0
    src = getattr(app_state, "reporte_resumen_por_juego", {}) or {}

    for juego, datos_juego in src.items():
        if not isinstance(datos_juego, dict):
            continue
        juego_key = str(juego)
        if juego_key not in semanas_sorteos:
            continue
        sorteos_permitidos = semanas_sorteos.get(juego_key, set())

        for sorteo, data in datos_juego.items():
            if _norm_sorteo(sorteo) not in sorteos_permitidos:
                continue
            if not isinstance(data, dict):
                continue
            total += _obtener_valor_resumen(data, campo)

    return total


def _obtener_valor_resumen(data: dict, campo: str) -> float:
    """
    Compatibilidad entre nombres legacy y actuales.
    - actuales: recaud / comi / prem
    - legacy: venta / comision / premio
    """
    aliases = {
        "venta": ("venta", "recaud"),
        "recaud": ("recaud", "venta"),
        "comision": ("comision", "comi"),
        "comi": ("comi", "comision"),
        "premio": ("premio", "prem"),
        "prem": ("prem", "premio"),
    }
    keys = aliases.get(str(campo), (str(campo),))
    for key in keys:
        if key in data:
            try:
                return float(data.get(key, 0.0) or 0.0)
            except Exception:
                return 0.0
    return 0.0


def _sumar_area_recaudacion_por_semana(semanas_sorteos: dict[str, set[str]], col_idx: int) -> float:
    """
    Suma una columna numérica de las filas guardadas/visibles del Área Recaudación,
    filtrando por sorteos de la semana activa.
    """
    total = 0.0
    snapshots = getattr(app_state, "obtener_snapshots_area_recaudacion", lambda: {})() or {}

    for juego, sorteos_semana in semanas_sorteos.items():
        # Si para un juego no hay sorteos asociados a la semana activa,
        # no debemos sumar todas las filas (mezcla semanas/filas especiales).
        if not sorteos_semana:
            continue
        snap = snapshots.get(juego, {})
        filas = snap.get("filas", []) if isinstance(snap, dict) else []
        if not isinstance(filas, list):
            continue

        for fila in filas:
            if not isinstance(fila, list) or len(fila) <= col_idx:
                continue
            sorteo = _norm_sorteo(fila[0] if fila else "")
            if sorteo not in sorteos_semana:
                continue
            total += _parse_importe(fila[col_idx])

    return total


def _sumar_fila_totales_area_recaudacion(semanas_sorteos: dict[str, set[str]], col_idx: int) -> float:
    """
    Suma, por juego, el valor de la fila "Totales" de Área Recaudación para la
    columna indicada.

    Si un snapshot no tiene fila "Totales", hace fallback a la suma por sorteos
    de la semana activa para ese juego.
    """
    total = 0.0
    snapshots = getattr(app_state, "obtener_snapshots_area_recaudacion", lambda: {})() or {}

    for juego, sorteos_semana in semanas_sorteos.items():
        snap = snapshots.get(juego, {})
        filas = snap.get("filas", []) if isinstance(snap, dict) else []
        if not isinstance(filas, list):
            continue

        total_juego = None

        for fila in filas:
            if not isinstance(fila, list) or len(fila) <= col_idx:
                continue
            etiqueta = str(fila[0] if fila else "").strip().lower()
            if etiqueta == "totales":
                total_juego = _parse_importe(fila[col_idx])
                break

        usar_fallback_por_sorteos = total_juego is None

        if usar_fallback_por_sorteos:
            total_juego = 0.0
            if not sorteos_semana:
                continue
            for fila in filas:
                if not isinstance(fila, list) or len(fila) <= col_idx:
                    continue
                sorteo = _norm_sorteo(fila[0] if fila else "")
                if sorteo not in sorteos_semana:
                    continue
                total_juego += _parse_importe(fila[col_idx])

        total += total_juego

    return total


def _sorteos_prescripciones_por_juego(semana_sel: str) -> dict[str, set[str]]:
    out: dict[str, set[str]] = {}
    sem_n = _semana_a_numero(semana_sel)

    src = getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", {}) or {}
    for juego, semanas in (src.items() if isinstance(src, dict) else []):
        if not isinstance(semanas, dict):
            continue
        sorteos = semanas.get(sem_n, semanas.get(str(sem_n), []))
        out[str(juego)] = {_norm_sorteo(s) for s in sorteos if str(s).strip()}

    return out


def _sumar_prescripciones_sfa_por_semana(semana_sel: str) -> float:
    """
    Total prescripción:
    suma de la columna Prescripto / Control SFA de la sección Prescripciones
    para la semana activa.
    """
    total = 0.0
    src = getattr(app_state, "sfa_prescripciones_por_juego", {}) or {}
    sorteos_por_juego = _sorteos_prescripciones_por_juego(semana_sel)

    for juego, sorteos in sorteos_por_juego.items():
        datos_juego = src.get(juego, {})
        if not isinstance(datos_juego, dict):
            continue

        for sorteo in sorteos:
            total += float(datos_juego.get(sorteo, 0.0) or 0.0)

    return total


def _sumar_prescripciones_reporte_por_semana(semana_sel: str) -> float:
    total = 0.0
    src = getattr(app_state, "reporte_prescripciones_por_juego", {}) or {}
    sorteos_por_juego = _sorteos_prescripciones_por_juego(semana_sel)

    for juego, sorteos in sorteos_por_juego.items():
        datos_juego = src.get(juego, {})
        if not isinstance(datos_juego, dict):
            continue

        for sorteo in sorteos:
            total += float(datos_juego.get(sorteo, 0.0) or 0.0)

    return total


def _totales_anticipos_topes(semana_sel: str) -> tuple[float, float, float, float]:
    payload = getattr(app_state, "planilla_anticipos_topes_data", {}) or {}
    semanas = payload.get("semanas", {}) if isinstance(payload, dict) else {}
    rows = semanas.get(semana_sel, []) if isinstance(semanas, dict) else []

    anticipos_tobill = 0.0
    topes_tobill = 0.0
    anticipos_sfa = 0.0
    topes_sfa = 0.0

    for row in rows if isinstance(rows, list) else []:
        if not isinstance(row, dict):
            continue

        concepto = str(row.get("concepto", "") or "").strip().lower()
        monto_tobill = _parse_importe(row.get("reporte_prescripto", ""))
        monto_sfa = _parse_importe(row.get("sfa_prescripto", ""))

        if concepto == "anticipos":
            anticipos_tobill += monto_tobill
            anticipos_sfa += monto_sfa
        elif concepto.startswith("anti rec"):
            topes_tobill += monto_tobill
            topes_sfa += monto_sfa

    # Fallback: si en Anticipos/Topes todavía no está cargada la columna SFA
    # (por ejemplo, luego de "Pasar datos a planilla" desde la solapa SFA),
    # usar los totales SFA persistidos por semana.
    data_totales = getattr(app_state, "planilla_totales_data", {}) or {}
    semanas_totales = data_totales.get("semanas", {}) if isinstance(data_totales, dict) else {}
    bucket_semana = semanas_totales.get(semana_sel, {}) if isinstance(semanas_totales, dict) else {}
    bucket_sfa = bucket_semana.get("sfa", {}) if isinstance(bucket_semana, dict) else {}

    if isinstance(bucket_sfa, dict):
        if anticipos_sfa == 0.0:
            anticipos_sfa = float(bucket_sfa.get("Total anticipos", 0.0) or 0.0)
        if topes_sfa == 0.0:
            topes_sfa = float(bucket_sfa.get("Total topes", 0.0) or 0.0)

    return anticipos_tobill, topes_tobill, anticipos_sfa, topes_sfa


def _totales_txt_anticipos_topes(semana_sel: str) -> tuple[float, float]:
    """
    Para Totales TXT:
    - Total anticipos  -> fila "Anticipos", columna "total"
    - Total topes      -> fila "Totales Anti Rec", columna "total"
    """
    payload = getattr(app_state, "planilla_anticipos_topes_data", {}) or {}
    semanas = payload.get("semanas", {}) if isinstance(payload, dict) else {}
    rows = semanas.get(semana_sel, []) if isinstance(semanas, dict) else []

    total_anticipos = 0.0
    total_topes = 0.0

    for row in rows if isinstance(rows, list) else []:
        if not isinstance(row, dict):
            continue
        concepto = str(row.get("concepto", "") or "").strip().lower()
        total = _parse_importe(row.get("total", ""))

        if concepto == "anticipos":
            total_anticipos = total
        elif concepto == "totales anti rec":
            total_topes = total

    return total_anticipos, total_topes


def _total_comision_agencia_amiga_por_semana(semanas_sorteos: dict[str, set[str]]) -> float:
    total = 0.0
    src = getattr(app_state, "sfa_z118_por_juego", {}) or {}

    alias_por_juego = {
        "Quiniela": {"Quiniela", "80"},
        "Quiniela Ya": {"Quiniela Ya", "79"},
        "Poceada": {"Poceada", "82"},
        "Tombolina": {"Tombolina", "74"},
        "Quini 6": {"Quini 6", "66", "67", "68", "69"},
        "Brinco": {"Brinco", "13"},
        "Loto": {"Loto", "7", "8", "9", "10", "11"},
        "Loto 5": {"Loto 5", "5"},
        "LT": {"LT", "41"},
    }

    for juego_planilla, sorteos_semana in semanas_sorteos.items():
        if not sorteos_semana:
            continue

        aliases = alias_por_juego.get(str(juego_planilla), {str(juego_planilla)})

        for alias in aliases:
            datos_juego = src.get(alias, {})
            if not isinstance(datos_juego, dict):
                continue

            for sorteo, importe in datos_juego.items():
                if _norm_sorteo(sorteo) in sorteos_semana:
                    total += float(importe or 0.0)

    return total


def _totales_agencia_amiga_por_semana(semana_sel: str) -> tuple[float, float]:
    data = getattr(app_state, "planilla_agencia_amiga_data", {}) or {}
    juegos = data.get("juegos", {}) if isinstance(data, dict) else {}
    sem_key = str(_semana_a_numero(semana_sel))
    total_tobill = 0.0
    total_sfa = 0.0

    for semanas in (juegos.values() if isinstance(juegos, dict) else []):
        if not isinstance(semanas, dict):
            continue
        bucket = semanas.get(sem_key, {})
        if not isinstance(bucket, dict):
            continue
        for fila in bucket.values():
            if not isinstance(fila, dict):
                continue
            total_tobill += _parse_importe(fila.get("importe_tobill", ""))
            total_sfa += _parse_importe(fila.get("importe_sfa", ""))

    return total_tobill, total_sfa


def _total_agencia_amiga_ntf_por_semana(semana_sel: str) -> float:
    data = getattr(app_state, "planilla_agencia_amiga_data", {}) or {}
    juegos = data.get("juegos", {}) if isinstance(data, dict) else {}
    sem_key = str(_semana_a_numero(semana_sel))
    total_ntf = 0.0

    for semanas in (juegos.values() if isinstance(juegos, dict) else []):
        if not isinstance(semanas, dict):
            continue
        bucket = semanas.get(sem_key, {})
        if not isinstance(bucket, dict):
            continue
        for fila in bucket.values():
            if not isinstance(fila, dict):
                continue
            total_ntf += _parse_importe(fila.get("importe_ntf", ""))

    return total_ntf


def _totales_agencia_amiga_reporte_por_semana(
    semana_sel: str,
    semanas_sorteos: dict[str, set[str]],
) -> tuple[float, float]:
    """
    Calcula Agencia Amiga directo desde lo importado en Reporte Facturación,
    filtrando por sorteos de la semana activa.
    """
    total_tobill = 0.0
    total_sfa = 0.0

    sem_key = f"Semana {_semana_a_numero(semana_sel)}"
    src_tobill_sem = getattr(app_state, "reporte_agencia_amiga_tobill_por_juego_por_semana", {}) or {}
    src_sfa_sem = getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}) or {}
    src_tobill = (
        src_tobill_sem.get(sem_key, {})
        if isinstance(src_tobill_sem, dict) and isinstance(src_tobill_sem.get(sem_key, {}), dict)
        else getattr(app_state, "reporte_agencia_amiga_tobill_por_juego", {}) or {}
    )
    src_sfa = (
        src_sfa_sem.get(sem_key, {})
        if isinstance(src_sfa_sem, dict) and isinstance(src_sfa_sem.get(sem_key, {}), dict)
        else getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego", {}) or {}
    )

    for juego, sorteos_semana in semanas_sorteos.items():
        if not sorteos_semana:
            continue

        datos_tobill = src_tobill.get(juego, {}) if isinstance(src_tobill, dict) else {}
        if isinstance(datos_tobill, dict):
            for sorteo, importe in datos_tobill.items():
                if _norm_sorteo(sorteo) in sorteos_semana:
                    total_tobill += float(importe or 0.0)

        datos_sfa = src_sfa.get(juego, {}) if isinstance(src_sfa, dict) else {}
        if isinstance(datos_sfa, dict):
            for sorteo, importe in datos_sfa.items():
                if _norm_sorteo(sorteo) in sorteos_semana:
                    total_sfa += float(importe or 0.0)

    return total_tobill, total_sfa


def _sumar_prescripciones_tickets_por_semana(semana_sel: str) -> float:
    """
    Total TXT de prescripciones: suma de la columna Prescripto de Consulta Tickets
    para todos los juegos/sorteos de la semana.
    """
    total = 0.0
    src = getattr(app_state, "tickets_prescripciones_por_juego", {}) or {}
    sorteos_por_juego = _sorteos_prescripciones_por_juego(semana_sel)

    for juego, sorteos in sorteos_por_juego.items():
        datos_juego = src.get(juego, {})
        if not isinstance(datos_juego, dict):
            continue
        for sorteo in sorteos:
            total += float(datos_juego.get(sorteo, 0.0) or 0.0)

    return total


def _prescripciones_txt_merge_por_semana(semana: str) -> float:
    """
    Total TXT de prescripción por semana, priorizando la carga manual de
    Prescripciones (t_presc) por sorteo; si no existe, usa Tickets.
    
    Se define a nivel módulo para que el editor y el runtime la resuelvan
    correctamente desde cualquier helper de Totales.
    """
    semana_txt = _semana_interna(semana) or (str(semana or '').strip() or 'Semana 1')
    merged: dict[tuple[str, str], float] = {}
    sorteos_por_juego = _sorteos_prescripciones_por_juego(semana_txt)
    src_tickets = getattr(app_state, 'tickets_prescripciones_por_juego', {}) or {}
    src_manual = getattr(app_state, 'planilla_prescripciones_data', {}) or {}

    for juego, sorteos in (sorteos_por_juego.items() if isinstance(sorteos_por_juego, dict) else []):
        datos_t = src_tickets.get(juego, {}) if isinstance(src_tickets, dict) else {}
        datos_m_juego = src_manual.get(juego, {}) if isinstance(src_manual, dict) else {}
        datos_m_sem = datos_m_juego.get(semana_txt, {}) if isinstance(datos_m_juego, dict) else {}
        for sorteo in sorteos if isinstance(sorteos, set) or isinstance(sorteos, list) else []:
            sorteo_key = _norm_sorteo(sorteo)
            if not sorteo_key:
                continue
            clave = (str(juego), sorteo_key)
            manual_val = None
            if isinstance(datos_m_sem, dict):
                fila = datos_m_sem.get(sorteo_key, {})
                if isinstance(fila, dict):
                    raw = fila.get('t_presc', None)
                    if raw not in (None, ''):
                        try:
                            manual_val = float(raw or 0.0)
                        except Exception:
                            manual_val = 0.0
            if manual_val is not None:
                merged[clave] = manual_val
            elif isinstance(datos_t, dict):
                try:
                    merged[clave] = float(datos_t.get(sorteo_key, 0.0) or 0.0)
                except Exception:
                    merged[clave] = 0.0

    return float(sum(merged.values()))


def _calcular_totales_txt_por_semana(semana_sel: str, valor_facuni_sfa: float) -> list[float]:
    semanas_sorteos = _semanas_sorteos_por_juego(semana_sel)
    ventas = _sumar_fila_totales_area_recaudacion(semanas_sorteos, 1)
    comisiones = _sumar_fila_totales_area_recaudacion(semanas_sorteos, 2)
    premios = _sumar_fila_totales_area_recaudacion(semanas_sorteos, 3)
    prescripcion = _sumar_prescripciones_tickets_por_semana(semana_sel)
    anticipos, topes = _totales_txt_anticipos_topes(semana_sel)
    agencia_amiga_ntf = _total_agencia_amiga_ntf_por_semana(semana_sel)

    return [
        ventas,
        comisiones,
        premios,
        prescripcion,
        anticipos,
        topes,
        float(valor_facuni_sfa or 0.0),
        agencia_amiga_ntf,
    ]


def _leer_totales_importados_por_semana(semana: str) -> tuple[list[float], list[float], list[float]]:
    """
    Lee los totales persistidos por semana.

    Formato preferido (actual):
      planilla_totales_data["semanas"][semana]["Total ventas"] = {"tobill": x, "sfa": y}

    Formato legacy:
      planilla_totales_data["semanas"][semana]["tobill"]["Total ventas"] = x
      planilla_totales_data["semanas"][semana]["sfa"]["Total ventas"] = y
    """
    valores_txt = [0.0 for _ in FILAS_TOTALES]
    valores_tobill = [0.0 for _ in FILAS_TOTALES]
    valores_sfa = [0.0 for _ in FILAS_TOTALES]

    obtener = getattr(app_state, "obtener_totales_importados_semana", None)
    if callable(obtener):
        bucket = obtener(semana) or {}
    else:
        payload = getattr(app_state, "planilla_totales_data", {}) or {}
        semanas = payload.get("semanas", {}) if isinstance(payload, dict) else {}
        bucket = semanas.get(semana, {}) if isinstance(semanas, dict) else {}

    if not isinstance(bucket, dict):
        return valores_txt, valores_tobill, valores_sfa

    # Formato preferido por fila.
    for idx, etiqueta in enumerate(FILAS_TOTALES):
        fila = bucket.get(etiqueta, {})
        if not isinstance(fila, dict):
            continue
        try:
            valores_txt[idx] = float(fila.get("txt", 0.0) or 0.0)
        except Exception:
            valores_txt[idx] = 0.0
        try:
            valores_tobill[idx] = float(fila.get("tobill", 0.0) or 0.0)
        except Exception:
            valores_tobill[idx] = 0.0
        try:
            valores_sfa[idx] = float(fila.get("sfa", 0.0) or 0.0)
        except Exception:
            valores_sfa[idx] = 0.0

    # Compatibilidad con formato legacy por columna.
    bucket_tobill = bucket.get("tobill", {}) if isinstance(bucket.get("tobill", {}), dict) else {}
    bucket_sfa = bucket.get("sfa", {}) if isinstance(bucket.get("sfa", {}), dict) else {}
    if bucket_tobill or bucket_sfa:
        for idx, etiqueta in enumerate(FILAS_TOTALES):
            if etiqueta in bucket_tobill:
                try:
                    valores_tobill[idx] = float(bucket_tobill.get(etiqueta, 0.0) or 0.0)
                except Exception:
                    valores_tobill[idx] = 0.0
            if etiqueta in bucket_sfa:
                try:
                    valores_sfa[idx] = float(bucket_sfa.get(etiqueta, 0.0) or 0.0)
                except Exception:
                    valores_sfa[idx] = 0.0

    bucket_txt = bucket.get("txt", {}) if isinstance(bucket.get("txt", {}), dict) else {}
    if bucket_txt:
        for idx, etiqueta in enumerate(FILAS_TOTALES):
            if etiqueta in bucket_txt:
                try:
                    valores_txt[idx] = float(bucket_txt.get(etiqueta, 0.0) or 0.0)
                except Exception:
                    valores_txt[idx] = 0.0

    return valores_txt, valores_tobill, valores_sfa


def _manuales_totales_por_semana(semana: str) -> dict[str, dict[str, float]]:
    data = getattr(app_state, "planilla_totales_data", {}) or {}
    manuales = data.get("manuales", {}) if isinstance(data, dict) else {}
    bucket = manuales.get(semana, {}) if isinstance(manuales, dict) else {}
    out: dict[str, dict[str, float]] = {}
    if not isinstance(bucket, dict):
        return out

    for etiqueta, payload in bucket.items():
        if str(etiqueta) not in FILAS_TOTALES or not isinstance(payload, dict):
            continue
        fila_manual: dict[str, float] = {}
        if "txt" in payload:
            fila_manual["txt"] = float(payload.get("txt", 0.0) or 0.0)
        if "tobill" in payload:
            fila_manual["tobill"] = float(payload.get("tobill", 0.0) or 0.0)
        if "sfa" in payload:
            fila_manual["sfa"] = float(payload.get("sfa", 0.0) or 0.0)
        if fila_manual:
            out[str(etiqueta)] = fila_manual
    return out


def _set_manual_total(semana: str, etiqueta: str, col: str, valor: float):
    if etiqueta not in FILAS_TOTALES or col not in {"txt", "tobill", "sfa"}:
        return
    if not hasattr(app_state, "planilla_totales_data") or not isinstance(getattr(app_state, "planilla_totales_data", None), dict):
        app_state.planilla_totales_data = {}
    manuales = app_state.planilla_totales_data.setdefault("manuales", {})
    if not isinstance(manuales, dict):
        manuales = {}
        app_state.planilla_totales_data["manuales"] = manuales
    semana_bucket = manuales.setdefault(semana, {})
    if not isinstance(semana_bucket, dict):
        semana_bucket = {}
        manuales[semana] = semana_bucket
    fila = semana_bucket.setdefault(etiqueta, {})
    if not isinstance(fila, dict):
        fila = {}
        semana_bucket[etiqueta] = fila
    fila[col] = float(valor or 0.0)


def _limpiar_manuales_txt(semana: str, etiquetas: list[str] | tuple[str, ...]):
    data = getattr(app_state, "planilla_totales_data", None)
    if not isinstance(data, dict):
        return
    manuales = data.get("manuales", {})
    if not isinstance(manuales, dict):
        return
    bucket = manuales.get(semana, {})
    if not isinstance(bucket, dict):
        return

    hubo_cambios = False
    for etiqueta in etiquetas:
        payload = bucket.get(etiqueta)
        if not isinstance(payload, dict) or "txt" not in payload:
            continue
        payload.pop("txt", None)
        hubo_cambios = True
        if not payload:
            bucket.pop(etiqueta, None)
    if hubo_cambios and not bucket:
        manuales.pop(semana, None)


def _recalcular_txt_desde_area_recaudacion_y_guardar(semana_sel: str) -> dict[str, float]:
    """
    Recalcula TODA la columna Totales TXT para la semana seleccionada y la
    persiste en planilla_totales_data.

    Incluye:
    - ventas / comisiones / premios desde Área Recaudación
    - prescripción TXT mergeando tickets + manuales de Prescripciones
    - anticipos / topes desde Anticipos y Topes
    - facuni desde FACUNI por semana (o valor TXT guardado si no hay dato vigente)
    - comisión agencia amiga desde Agencia Amiga (NTF)
    """
    semana_txt = _semana_interna(semana_sel) or SEMANAS[0]

    bucket_actual = {}
    obtener = getattr(app_state, "obtener_totales_importados_semana", None)
    if callable(obtener):
        try:
            bucket_actual = obtener(semana_txt) or {}
        except Exception:
            bucket_actual = {}

    facuni_txt = 0.0
    try:
        fac_map = getattr(app_state, "facuni_total_por_semana", {}) or {}
        facuni_txt = float(fac_map.get(semana_txt, 0.0) or 0.0)
    except Exception:
        facuni_txt = 0.0

    if facuni_txt == 0.0 and isinstance(bucket_actual, dict):
        fila_facuni = bucket_actual.get("Total facuni", {})
        if isinstance(fila_facuni, dict):
            try:
                facuni_txt = float(fila_facuni.get("txt", fila_facuni.get("sfa", 0.0)) or 0.0)
            except Exception:
                facuni_txt = 0.0

    valores = _calcular_totales_txt_por_semana(semana_txt, facuni_txt)
    payload_txt = {etiqueta: round(float(valor or 0.0), 2) for etiqueta, valor in zip(FILAS_TOTALES, valores)}

    # Prescripción TXT: priorizar merge tickets + manuales de Prescripciones.
    try:
        payload_txt["Total prescripcion"] = round(float(_prescripciones_txt_merge_por_semana(semana_txt) or 0.0), 2)
    except Exception:
        payload_txt["Total prescripcion"] = round(float(payload_txt.get("Total prescripcion", 0.0) or 0.0), 2)

    guardar = getattr(app_state, "guardar_totales_importados", None)
    if callable(guardar):
        guardar(semana_txt, "txt", payload_txt)
    else:
        if not hasattr(app_state, "planilla_totales_data") or not isinstance(getattr(app_state, "planilla_totales_data", None), dict):
            app_state.planilla_totales_data = {}
        semanas = app_state.planilla_totales_data.setdefault("semanas", {})
        if not isinstance(semanas, dict):
            semanas = {}
            app_state.planilla_totales_data["semanas"] = semanas
        bucket = semanas.setdefault(semana_txt, {})
        if not isinstance(bucket, dict):
            bucket = {}
            semanas[semana_txt] = bucket
        for etiqueta, valor in payload_txt.items():
            fila = bucket.setdefault(etiqueta, {})
            if not isinstance(fila, dict):
                fila = {}
                bucket[etiqueta] = fila
            fila["txt"] = float(valor or 0.0)

    # Limpiar cualquier manual TXT previo de todas las filas recalculadas para
    # que al volver a la sección no reaparezcan valores viejos.
    _limpiar_manuales_txt(semana_txt, tuple(payload_txt.keys()))
    return payload_txt


def build_totales(fr_seccion: ttk.Frame, estado_var):
    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(0, weight=0)
    fr_seccion.rowconfigure(1, weight=1)

    top = ttk.Frame(fr_seccion, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=(2, 10), pady=(6, 10))
    top.columnconfigure(2, weight=1)

    toolbar_style = str(getattr(app_state, "planilla_toolbar_combobox_style", "") or "").strip()
    try:
        toolbar_width = int(getattr(app_state, "planilla_toolbar_combobox_width", 30) or 30)
    except Exception:
        toolbar_width = 30

    ttk.Label(top, text="Semana:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")
    semana_var = tk.StringVar(value="")
    combo_semana_kwargs = {
        "state": "disabled",
        "values": _combo_values_semanas(),
        "textvariable": semana_var,
        "width": toolbar_width,
    }
    if toolbar_style:
        combo_semana_kwargs["style"] = toolbar_style
    combo_semana = ttk.Combobox(top, **combo_semana_kwargs)
    combo_semana.grid(row=0, column=1, sticky="w", padx=(8, 10))
    apply_filter_combobox_style(combo_semana, style_name=toolbar_style or "Planilla.Toolbar.TCombobox")

    def _sync_combo_semanas(reset_selection: bool = False, prefer_semana: int = 0):
        values = list(_combo_values_semanas() or [])
        try:
            combo_semana.configure(values=values)
        except Exception:
            pass
        if reset_selection or not values:
            semana_var.set("")
            try:
                combo_semana.set("")
                combo_semana.configure(state="disabled")
            except Exception:
                pass
            return
        try:
            combo_semana.configure(state="readonly")
        except Exception:
            pass
        visible_pref = _semana_visible(f"Semana {prefer_semana}") if 1 <= int(prefer_semana or 0) <= 5 else ""
        actual = str(semana_var.get() or combo_semana.get() or "").strip()
        target = visible_pref if visible_pref in values else (actual if actual in values else values[0])
        semana_var.set(target)
        try:
            combo_semana.set(target)
        except Exception:
            pass

    _week_sync_guard = {"active": False}

    def _publicar_semana_global_desde_combo(semana_txt: str):
        semana_interna = _semana_interna(semana_txt or "")
        if not semana_interna:
            return
        try:
            semana_n = int(str(semana_interna).lower().replace("semana", "").strip() or "0")
        except Exception:
            semana_n = 0
        if semana_n < 1:
            return

        desde = ""
        hasta = ""
        try:
            rangos = getattr(app_state, "planilla_rangos_semana_global", {}) or {}
            rango = rangos.get(semana_n, rangos.get(str(semana_n))) if isinstance(rangos, dict) else None
            if isinstance(rango, dict):
                desde = str(rango.get("desde", "") or "").strip()
                hasta = str(rango.get("hasta", "") or "").strip()
            elif isinstance(rango, (list, tuple)) and len(rango) >= 2:
                desde = str(rango[0] or "").strip()
                hasta = str(rango[1] or "").strip()
        except Exception:
            pass

        juego_actual = ""
        try:
            payload_actual = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            if isinstance(payload_actual, dict):
                juego_actual = str(payload_actual.get("juego", "") or "").strip()
        except Exception:
            pass

        publicar = getattr(app_state, "publicar_filtro_area_recaudacion", None)
        if callable(publicar):
            try:
                publicar(juego_actual, semana_n, desde, hasta)
            except Exception:
                pass

    def _on_semana_combo_selected(_e=None):
        if not _week_sync_guard["active"]:
            _publicar_semana_global_desde_combo(semana_var.get() or combo_semana.get() or "")
        _programar_recalculo(0)

    btn_recalcular_txt = ttk.Button(top, text="Recalcular Totales TXT", style="Marino.TButton")
    # El botón debe permanecer oculto. La acción sólo se expone con click
    # derecho sobre la columna Totales TXT.

    card = tk.Frame(
        fr_seccion,
        bg="#E7EEF6",
        highlightthickness=1,
        highlightbackground="#C7D5E4",
        bd=0,
    )
    card.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
    card.columnconfigure(0, weight=1)
    card.rowconfigure(0, weight=1)

    tabla_wrap = ttk.Frame(card, style="Panel.TFrame")
    tabla_wrap.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
    tabla_wrap.columnconfigure(0, weight=1)
    tabla_wrap.rowconfigure(0, weight=0)
    tabla_wrap.rowconfigure(1, weight=1)

    header_canvas = tk.Canvas(tabla_wrap, height=38, bg="#D3E3F4", highlightthickness=0, bd=0)
    header_canvas.grid(row=0, column=0, sticky="ew")

    tree = ttk.Treeview(
        tabla_wrap,
        columns=("total", "totales_txt", "totales_tobill", "totales_sfa", "diferencias"),
        show="",
        height=len(FILAS_TOTALES) + 1,
    )
    tree.grid(row=1, column=0, sticky="nsew")
    vsb = ttk.Scrollbar(tabla_wrap, orient="vertical", command=tree.yview)
    vsb.grid(row=1, column=1, sticky="ns")

    tree.configure(yscrollcommand=vsb.set)

    style = ttk.Style(fr_seccion)
    style.configure(
        "Totales.Treeview",
        rowheight=30,
        font=("Segoe UI", 9),
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground="#0F172A",
        borderwidth=0,
        relief="flat",
    )
    style.map(
        "Totales.Treeview",
        background=[("selected", "#DCEBFF")],
        foreground=[("selected", "#0F172A")],
    )
    tree.configure(style="Totales.Treeview")


    for col in ("total", "totales_txt", "totales_tobill", "totales_sfa", "diferencias"):
        tree.heading(col, text="")

    base_widths = {
        "total": 360,
        "totales_txt": 185,
        "totales_tobill": 185,
        "totales_sfa": 185,
        "diferencias": 180,
    }
    for col, width in base_widths.items():
        anchor = "w" if col == "total" else "center"
        tree.column(col, width=width, minwidth=120, stretch=False, anchor=anchor)

    tree.tag_configure("total_final", background="#E2F0E8", foreground="#0F172A", font=("Segoe UI Semibold", 9))
    tree.tag_configure("diff_warn", background="#FFF1D6", foreground="#0F172A")
    tree.tag_configure("diff_danger", background="#FDE4E4", foreground="#0F172A")
    tree.tag_configure("fila_par", background="#F7FAFD", foreground="#0F172A")
    tree.tag_configure("fila_impar", background="#FFFFFF", foreground="#0F172A")

    HEADER_STYLES = {
        "total": ("Concepto / Totales", "#D7E3F1"),
        "totales_txt": ("Totales TXT", "#BCD3F3"),
        "totales_tobill": ("Totales Tobill", "#BEDFCF"),
        "totales_sfa": ("Totales SFA", "#F2D2A6"),
        "diferencias": ("Diferencias", "#EFDDA2"),
    }
    BORDER = "#95A8BC"
    BORDER_BOTTOM = "#869AB0"
    TEXT_HEADER = "#142536"

    def _draw_header():
        header_canvas.delete("all")
        x = 0
        h = 38
        columnas = ("total", "totales_txt", "totales_tobill", "totales_sfa", "diferencias")
        for idx, col in enumerate(columnas):
            w = int(tree.column(col, "width") or base_widths[col])
            text_label, fill = HEADER_STYLES[col]

            x1 = x
            x2 = x + w
            if idx == 0:
                x1 = max(0, x1)
            else:
                x1 = x1 - 1  # solapar 1 px para evitar la "costura" blanca entre columnas

            header_canvas.create_rectangle(x1, 0, x2, h, fill=fill, outline=BORDER, width=1)
            header_canvas.create_line(x1, h - 1, x2, h - 1, fill=BORDER_BOTTOM, width=1)

            if col == "total":
                header_canvas.create_text(
                    x + 14,
                    h / 2,
                    text=text_label,
                    font=("Segoe UI Semibold", 10),
                    fill=TEXT_HEADER,
                    anchor="w",
                )
            else:
                header_canvas.create_text(
                    x + w / 2,
                    h / 2,
                    text=text_label,
                    font=("Segoe UI Semibold", 10),
                    fill=TEXT_HEADER,
                    anchor="center",
                )
            x += w

        try:
            header_canvas.configure(width=max(1, int(tree.winfo_width() or x)), height=h)
            header_canvas.lift()
        except Exception:
            pass
        header_canvas.configure(scrollregion=(0, 0, x, h))

    def _ajustar_columnas_a_ancho(_evt=None):
        try:
            ancho_wrap = int(tabla_wrap.winfo_width() or 0)
        except Exception:
            ancho_wrap = 0
        try:
            ancho_seccion = int(fr_seccion.winfo_width() or 0) - 18
        except Exception:
            ancho_seccion = 0

        ancho_disponible = max(ancho_wrap, ancho_seccion, sum(base_widths.values()))

        base_total = sum(base_widths.values())
        extra = max(0, ancho_disponible - base_total)

        proporciones = {
            "total": 0.37,
            "totales_txt": 0.16,
            "totales_tobill": 0.16,
            "totales_sfa": 0.16,
            "diferencias": 0.15,
        }

        widths = {}
        acumulado = 0
        columnas = ("total", "totales_txt", "totales_tobill", "totales_sfa", "diferencias")
        for i, col in enumerate(columnas):
            if i == len(columnas) - 1:
                w = ancho_disponible - acumulado
            else:
                w = int(base_widths[col] + extra * proporciones[col])
                acumulado += w
            widths[col] = max(120, w)

        for col in columnas:
            tree.column(
                col,
                width=widths[col],
                minwidth=120,
                stretch=True,
                anchor=("w" if col == "total" else "center"),
            )

        _draw_header()

    if not hasattr(app_state, "planilla_totales_data") or not isinstance(getattr(app_state, "planilla_totales_data", None), dict):
        app_state.planilla_totales_data = {}

    def _actualizar_label_rango_del_al(semana_txt: str):
        return

    editing = {"widget": None, "commit": None}

    def _destroy_editor():
        editor = editing.get("widget")
        if editor is not None:
            try:
                editor.destroy()
            except Exception:
                pass
        editing["widget"] = None
        editing["commit"] = None

    def _commit_editor():
        commit = editing.get("commit")
        if not callable(commit):
            return False
        try:
            commit()
        except Exception:
            return False
        return True

    def _texto_etiqueta_profesional(etiqueta: str) -> str:
        txt = str(etiqueta or "").strip()
        if txt.upper() == "TOTALES PLANILLA FACTURACION":
            return txt
        return f"  {txt}"

    def _render_tabla_vacia():
        for iid in tree.get_children():
            tree.delete(iid)

        for idx, etiqueta in enumerate(FILAS_TOTALES):
            tags = ("fila_par" if idx % 2 == 0 else "fila_impar",)
            tree.insert(
                "",
                "end",
                values=(_texto_etiqueta_profesional(etiqueta), _fmt_pesos(0.0), _fmt_pesos(0.0), _fmt_pesos(0.0), _fmt_pesos(0.0)),
                tags=tags,
            )

        tree.insert(
            "",
            "end",
            values=("TOTALES PLANILLA FACTURACION", _fmt_pesos(0.0), _fmt_pesos(0.0), _fmt_pesos(0.0), _fmt_pesos(0.0)),
            tags=("total_final",),
        )
        _ajustar_columnas_a_ancho()
    _render_tabla_vacia()

    def _cerrar_editor(save: bool = False):
        if save and _commit_editor():
            return
        _destroy_editor()

    def _recalcular(*_args, force: bool = False):
        if editing.get("widget") is not None and not force:
            return
        seleccion_actual = tree.selection()
        etiqueta_seleccionada = None
        if seleccion_actual:
            valores_sel = tree.item(seleccion_actual[0], "values")
            if valores_sel:
                etiqueta_seleccionada = str(valores_sel[0] or "").strip()

        semana = _semana_interna(semana_var.get())
        if not semana:
            _render_tabla_vacia()
            if estado_var is not None:
                estado_var.set("Totales limpios. La tabla queda disponible para carga manual.")
            return

        _actualizar_label_rango_del_al(semana)

        valores_txt, valores_tobill, valores_sfa = _leer_totales_importados_por_semana(semana)
        manuales = _manuales_totales_por_semana(semana)
        for idx, etiqueta in enumerate(FILAS_TOTALES):
            payload = manuales.get(etiqueta, {})
            if isinstance(payload, dict):
                if "txt" in payload:
                    valores_txt[idx] = float(payload.get("txt", 0.0) or 0.0)
                if "tobill" in payload:
                    valores_tobill[idx] = float(payload.get("tobill", 0.0) or 0.0)
                if "sfa" in payload:
                    valores_sfa[idx] = float(payload.get("sfa", 0.0) or 0.0)

        # Total prescripcion TXT desde tickets + manual, sin duplicar y priorizando manual.
        try:
            idx_presc = FILAS_TOTALES.index("Total prescripcion")
            valores_txt[idx_presc] = _prescripciones_txt_merge_por_semana(semana)
        except Exception:
            pass

        for iid in tree.get_children():
            tree.delete(iid)

        iid_por_etiqueta = {}
        for idx, (etiqueta, valor_txt, valor_tobill, valor_sfa) in enumerate(zip(FILAS_TOTALES, valores_txt, valores_tobill, valores_sfa)):
            diferencia = float(valor_sfa - valor_txt)
            tags = ["fila_par" if idx % 2 == 0 else "fila_impar"]
            alerta = _clasificar_diferencia_fila_total(diferencia)
            if alerta == "danger":
                tags = ["diff_danger"]
            elif alerta == "warn":
                tags = ["diff_warn"]
            iid = tree.insert("", "end", values=(_texto_etiqueta_profesional(etiqueta), _fmt_pesos(valor_txt), _fmt_pesos(valor_tobill), _fmt_pesos(valor_sfa), _fmt_pesos(diferencia)), tags=tuple(tags))
            iid_por_etiqueta[etiqueta] = iid

        factores_planilla = [1, -1, -1, 1, 1, -1, 1, -1]
        total_planilla_facturacion_txt = sum(v * f for v, f in zip(valores_txt, factores_planilla))
        total_planilla_facturacion_tobill = sum(v * f for v, f in zip(valores_tobill, factores_planilla))
        total_planilla_facturacion_sfa = sum(v * f for v, f in zip(valores_sfa, factores_planilla))
        iid_total_final = tree.insert("", "end", values=("TOTALES PLANILLA FACTURACION", _fmt_pesos(total_planilla_facturacion_txt), _fmt_pesos(total_planilla_facturacion_tobill), _fmt_pesos(total_planilla_facturacion_sfa), _fmt_pesos(total_planilla_facturacion_sfa - total_planilla_facturacion_txt)), tags=("total_final",))
        iid_por_etiqueta["TOTALES PLANILLA FACTURACION"] = iid_total_final

        if etiqueta_seleccionada:
            iid_reseleccion = iid_por_etiqueta.get(etiqueta_seleccionada)
            if iid_reseleccion:
                tree.selection_set(iid_reseleccion)
                tree.focus(iid_reseleccion)

        _ajustar_columnas_a_ancho()
        if estado_var is not None:
            estado_var.set(f"Totales calculados para {semana}.")

    editable_columns = {"#2": "txt", "#3": "tobill", "#4": "sfa"}
    editable_col_indices = {1, 2, 3}
    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)

    def _editar_celda(event):
        iid = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not iid or col not in editable_columns:
            return
        vals = tree.item(iid, "values")
        if not vals:
            return
        etiqueta = str(vals[0] or "").strip()
        if etiqueta not in FILAS_TOTALES:
            return
        clipboard_state["cell"] = (iid, int(col.replace("#", "")) - 1)
        bbox = tree.bbox(iid, col)
        if not bbox:
            return
        x, y, width, height = bbox
        if width <= 1 or height <= 1:
            return
        _cerrar_editor()
        editor = ttk.Entry(tree)
        editor.place(x=x, y=y, width=width, height=height)
        col_idx_map = {"#2": 1, "#3": 2, "#4": 3}
        col_idx = col_idx_map[col]
        editor.insert(0, str(vals[col_idx] or ""))
        editor.focus_set()
        editor.select_range(0, "end")
        editing["widget"] = editor

        def _guardar(_evt=None):
            if editing.get("widget") is not editor:
                return "break"
            texto = editor.get()
            valor = _parse_importe(texto)
            semana = _semana_interna(semana_var.get()) or ""
            campo = editable_columns[col]
            valor_prev = _parse_importe(vals[col_idx])
            if valor_prev != valor:
                push_undo_rows(undo_state, tree, [(iid, [str(v) for v in vals])], meta={"semana": semana})
                _set_manual_total(semana, etiqueta, campo, valor)
            _destroy_editor()
            _recalcular()
            return "break"

        editing["commit"] = lambda: _guardar()

        editor.bind("<Return>", _guardar)
        editor.bind("<KP_Enter>", _guardar)
        editor.bind("<Escape>", lambda _e: _cerrar_editor())
        editor.bind("<FocusOut>", _guardar)
        return "break"

    def _copiar_celdas(_evt=None):
        iid, col_idx = get_anchor_cell(tree, clipboard_state, default_col=1)
        if not iid or col_idx not in editable_col_indices:
            return "break"
        seleccion = ordered_selected_rows(tree)
        row_ids = seleccion if len(seleccion) > 1 and iid in seleccion else [iid]
        matrix = []
        for row_iid in row_ids:
            vals = list(tree.item(row_iid, "values"))
            if not vals or str(vals[0] or "") not in FILAS_TOTALES:
                continue
            while len(vals) <= col_idx:
                vals.append("")
            matrix.append([vals[col_idx]])
        set_clipboard_matrix(tree, matrix)
        return "break"

    def _pegar_celdas(_evt=None):
        matrix = get_clipboard_matrix(tree)
        if not matrix:
            return "break"
        anchor_iid, anchor_col = get_anchor_cell(tree, clipboard_state, default_col=1)
        row_ids = [iid for iid in tree.get_children() if str((tree.item(iid, "values") or [""])[0]) in FILAS_TOTALES]
        if not anchor_iid or anchor_iid not in row_ids:
            return "break"
        semana = _semana_interna(semana_var.get()) or ""
        start = row_ids.index(anchor_iid)
        hubo_cambios = False
        undo_rows = []
        idx_to_col = {1: "txt", 2: "tobill", 3: "sfa"}
        for r_off, row_data in enumerate(matrix):
            row_pos = start + r_off
            if row_pos >= len(row_ids):
                break
            iid = row_ids[row_pos]
            vals = list(tree.item(iid, "values"))
            if not vals:
                continue
            etiqueta = str(vals[0] or "").strip()
            if etiqueta not in FILAS_TOTALES:
                continue
            row_before = list(vals)
            row_changed = False
            for c_off, cell_raw in enumerate(row_data):
                col_idx = anchor_col + c_off
                campo = idx_to_col.get(col_idx)
                if not campo:
                    continue
                parsed = _parse_importe(cell_raw)
                _set_manual_total(semana, etiqueta, campo, parsed)
                row_changed = True
            if row_changed:
                undo_rows.append((iid, [str(v) for v in row_before]))
                hubo_cambios = True
        if hubo_cambios:
            push_undo_rows(undo_state, tree, undo_rows, meta={"semana": semana})
            _recalcular()
        return "break"

    def _deshacer_celdas(_evt=None):
        snapshot = pop_undo_snapshot(undo_state)
        if not snapshot:
            return "break"
        semana_actual = _semana_interna(semana_var.get()) or ""
        semana_undo = str((snapshot.get("meta", {}) or {}).get("semana", semana_actual) or semana_actual)
        if semana_undo in SEMANAS and semana_undo != (_semana_interna(semana_var.get()) or ""):
            semana_var.set(semana_undo)
            combo_semana.set(_semana_visible(semana_undo))
        idx_to_col = {1: "txt", 2: "tobill", 3: "sfa"}
        for _iid, vals in (snapshot.get("rows", []) or []):
            vals_list = [str(v) for v in (vals or [])]
            if not vals_list:
                continue
            etiqueta = str(vals_list[0] or "")
            if etiqueta not in FILAS_TOTALES:
                continue
            for col_idx, campo in idx_to_col.items():
                if col_idx < len(vals_list):
                    _set_manual_total(semana_undo, etiqueta, campo, _parse_importe(vals_list[col_idx]))
        _recalcular()
        return "break"

    _refresh_state = {"job": None, "running": False, "pending": False}

    def _ejecutar_recalculo():
        if not fr_seccion.winfo_exists():
            return
        _refresh_state["job"] = None
        if _refresh_state["running"]:
            _refresh_state["pending"] = True
            return
        _refresh_state["running"] = True
        try:
            _recalcular()
        finally:
            _refresh_state["running"] = False
        if _refresh_state["pending"]:
            _refresh_state["pending"] = False
            _programar_recalculo(60)

    def _programar_recalculo(delay: int = 0):
        if not fr_seccion.winfo_exists():
            return
        job = _refresh_state.get("job")
        if job is not None:
            try:
                fr_seccion.after_cancel(job)
            except Exception:
                pass
            _refresh_state["job"] = None
        try:
            _refresh_state["job"] = fr_seccion.after(max(0, int(delay or 0)), _ejecutar_recalculo)
        except Exception:
            _ejecutar_recalculo()

    menu_ctx = tk.Menu(tree, tearoff=0)

    def _accion_recalcular_txt_semana_actual():
        semana = _semana_interna(semana_var.get()) or ""
        payload = _recalcular_txt_desde_area_recaudacion_y_guardar(semana)
        _recalcular(force=True)
        if estado_var is not None:
            estado_var.set(
                f"Totales TXT recalculados para {semana}: ventas={_fmt_pesos(payload.get('Total ventas', 0.0))}, comisiones={_fmt_pesos(payload.get('Total comisiones', 0.0))}, premios={_fmt_pesos(payload.get('Total premios', 0.0))}, prescripción={_fmt_pesos(payload.get('Total prescripcion', 0.0))}."
            )

    def _menu_contextual_totales(event):
        # Debe aparecer sólo sobre la columna Totales TXT (#2).
        try:
            region = tree.identify("region", getattr(event, "x", 0), getattr(event, "y", 0))
            col = tree.identify_column(getattr(event, "x", 0))
        except Exception:
            region, col = "", ""

        permitir = False
        if region in ("cell", "heading") and col == "#2":
            permitir = True
        else:
            try:
                x_canvas = header_canvas.canvasx(getattr(event, "x", 0))
            except Exception:
                x_canvas = getattr(event, "x", 0)
            width_total = int(tree.column("total", "width") or base_widths["total"])
            width_txt = int(tree.column("totales_txt", "width") or base_widths["totales_txt"])
            if width_total <= x_canvas <= (width_total + width_txt):
                permitir = True

        if not permitir:
            return "break"

        try:
            row = tree.identify_row(getattr(event, "y", 0))
            if row:
                tree.selection_set(row)
                tree.focus(row)
        except Exception:
            pass
        try:
            menu_ctx.delete(0, "end")
            menu_ctx.add_command(label="Recalcular Totales TXT", command=_accion_recalcular_txt_semana_actual)
            menu_ctx.tk_popup(event.x_root, event.y_root)
        finally:
            try:
                menu_ctx.grab_release()
            except Exception:
                pass
        return "break"

    btn_recalcular_txt.configure(command=_accion_recalcular_txt_semana_actual)
    combo_semana.bind("<<ComboboxSelected>>", _on_semana_combo_selected)
    tree.bind("<ButtonRelease-1>", _editar_celda)
    tree.bind("<Button-3>", _menu_contextual_totales, add="+")
    header_canvas.bind("<Button-3>", _menu_contextual_totales, add="+")
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")
    tree.bind("<Configure>", _ajustar_columnas_a_ancho, add="+")
    tabla_wrap.bind("<Configure>", _ajustar_columnas_a_ancho, add="+")
    fr_seccion.bind("<Configure>", _ajustar_columnas_a_ancho, add="+")
    tabla_wrap.bind("<Configure>", _ajustar_columnas_a_ancho, add="+")
    fr_seccion.bind("<Configure>", _ajustar_columnas_a_ancho, add="+")
    header_canvas.bind("<Configure>", lambda _e: _draw_header(), add="+")

    if not hasattr(app_state, "planilla_totales_refresh_hooks"):
        app_state.planilla_totales_refresh_hooks = {}
    app_state.planilla_totales_refresh_hooks["totales"] = (lambda: _programar_recalculo(80))

    if not hasattr(app_state, "planilla_bundle_snapshot_hooks"):
        app_state.planilla_bundle_snapshot_hooks = {}
    if not hasattr(app_state, "planilla_bundle_load_hooks"):
        app_state.planilla_bundle_load_hooks = {}

    def _snapshot_totales() -> dict:
        return copy.deepcopy(getattr(app_state, "planilla_totales_data", {}) or {})

    def _load_totales(payload: dict):
        nuevo = copy.deepcopy(payload) if isinstance(payload, dict) else {"semanas": {}, "manuales": {}}
        if not isinstance(getattr(app_state, "planilla_totales_data", None), dict):
            app_state.planilla_totales_data = {}
        app_state.planilla_totales_data.clear()
        app_state.planilla_totales_data.update(nuevo)
        try:
            payload_filtro = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            sem_n = int(payload_filtro.get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        if sem_n < 1:
            _sync_combo_semanas(reset_selection=True)
        else:
            _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
        _programar_recalculo(0)

    app_state.planilla_bundle_snapshot_hooks["totales"] = _snapshot_totales
    app_state.planilla_bundle_load_hooks["totales"] = _load_totales

    if not hasattr(app_state, "planilla_section_commit_hooks"):
        app_state.planilla_section_commit_hooks = {}

    def _commit_hook_totales():
        try:
            _cerrar_editor(save=True)
        except Exception:
            _destroy_editor()

    app_state.planilla_section_commit_hooks["TOTALES"] = _commit_hook_totales

    if not hasattr(app_state, "planilla_semana_filtro_hooks"):
        app_state.planilla_semana_filtro_hooks = {}

    def _aplicar_filtro_area_recaudacion(payload: dict):
        sem_n = 0
        try:
            sem_n = int((payload or {}).get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        _week_sync_guard["active"] = True
        try:
            if sem_n < 1:
                _sync_combo_semanas(reset_selection=True)
                _programar_recalculo(0)
                return
            _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
            _programar_recalculo(0)
        finally:
            _week_sync_guard["active"] = False

    app_state.planilla_semana_filtro_hooks["totales"] = _aplicar_filtro_area_recaudacion

    if not hasattr(app_state, "planilla_global_reset_hooks"):
        app_state.planilla_global_reset_hooks = {}

    def _reset_totales():
        try:
            _sync_combo_semanas(reset_selection=True)
        except Exception:
            pass
        _render_tabla_vacia()

    app_state.planilla_global_reset_hooks["totales"] = _reset_totales


    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    def _refresh_visual_totales():
        try:
            payload = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            sem_n = int(payload.get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        if sem_n < 1:
            _sync_combo_semanas(reset_selection=True)
        else:
            _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
        _programar_recalculo(0)

    app_state.planilla_visual_refresh_hooks["Totales"] = _refresh_visual_totales

    _ajustar_columnas_a_ancho()
    payload_inicial = getattr(app_state, "planilla_semana_filtro_actual", {})
    if isinstance(payload_inicial, dict):
        _aplicar_filtro_area_recaudacion(dict(payload_inicial))
    else:
        _sync_combo_semanas(reset_selection=True)
        _programar_recalculo(0)
