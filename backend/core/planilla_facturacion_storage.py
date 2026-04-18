from __future__ import annotations

import copy
import json
import os
from datetime import datetime

import app_state
from tabs.planilla_area_recaudacion import _ruta_guardado_planilla


def _ruta_bundle_default() -> str:
    appdata = os.environ.get("APPDATA")
    if appdata:
        base = os.path.join(appdata, "ReporteFacturacion")
    else:
        base = os.path.join(os.path.expanduser("~"), ".reporte_facturacion")

    if os.path.exists(base) and not os.path.isdir(base):
        raise RuntimeError(
            f"La ruta de storage existe pero no es una carpeta: {base}"
        )

    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "planilla_facturacion_bundle.json")


def _ruta_bundle_last_path_file() -> str:
    base = os.path.dirname(_ruta_bundle_default())
    if os.path.exists(base) and not os.path.isdir(base):
        raise RuntimeError(
            f"La ruta de storage existe pero no es una carpeta: {base}"
        )
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "planilla_facturacion_last_path.txt")


def _guardar_ruta_bundle(path: str):
    p = os.path.abspath(str(path or "").strip())
    if not p:
        return
    try:
        path_file = _ruta_bundle_last_path_file()
        tmp_file = f"{path_file}.tmp"
        with open(tmp_file, "w", encoding="utf-8") as f:
            f.write(p)
        os.replace(tmp_file, path_file)
    except Exception:
        pass


def _borrar_ruta_bundle_guardada():
    try:
        path_file = _ruta_bundle_last_path_file()
        if os.path.exists(path_file):
            os.remove(path_file)
    except Exception:
        pass


def _leer_ruta_bundle_guardada() -> str:
    path_file = _ruta_bundle_last_path_file()
    if not os.path.exists(path_file):
        return ""
    try:
        with open(path_file, "r", encoding="utf-8") as f:
            p = str(f.read() or "").strip()
    except Exception:
        return ""
    return p if p and os.path.exists(p) else ""


def _leer_json(path: str) -> dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def _escribir_json(path: str, payload: dict):
    folder = os.path.dirname(path) or "."
    os.makedirs(folder, exist_ok=True)
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)


def _replace_dict(name: str, values: dict):
    target = getattr(app_state, name, None)
    if isinstance(target, dict):
        target.clear()
        if isinstance(values, dict):
            target.update(values)


def _normalizar_rangos_area(values: dict) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    if not isinstance(values, dict):
        return out

    for semana_key, rango_raw in values.items():
        try:
            semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
        except Exception:
            continue
        if isinstance(rango_raw, dict):
            desde = str(rango_raw.get("desde", "") or "").strip()
            hasta = str(rango_raw.get("hasta", "") or "").strip()
        elif isinstance(rango_raw, (list, tuple)) and len(rango_raw) >= 2:
            desde = str(rango_raw[0] or "").strip()
            hasta = str(rango_raw[1] or "").strip()
        else:
            continue
        if desde and hasta:
            out[str(semana)] = {"desde": desde, "hasta": hasta}
    return out


def _normalizar_manual_overrides_area(values: dict) -> dict[str, dict[str, str]]:
    out: dict[str, dict[str, str]] = {}
    if not isinstance(values, dict):
        return out

    for sorteo_key, cols_raw in values.items():
        sorteo = str(sorteo_key or "").strip()
        if not sorteo or not isinstance(cols_raw, dict):
            continue
        row_out: dict[str, str] = {}
        for col_key, value in cols_raw.items():
            try:
                col_idx = int(str(col_key).strip())
            except Exception:
                continue
            if col_idx <= 0:
                continue
            row_out[str(col_idx)] = str(value or "")
        if row_out:
            out[sorteo] = row_out
    return out


def _asegurar_claves_semanas_desde_rangos(
    semanas: dict | None,
    rangos: dict | None,
) -> dict[str, list[int]]:
    out: dict[str, list[int]] = {}
    if isinstance(semanas, dict):
        for semana_key, sorteos_raw in semanas.items():
            try:
                semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
            except Exception:
                continue
            sorteos: list[int] = []
            if isinstance(sorteos_raw, list):
                for item in sorteos_raw:
                    try:
                        sorteos.append(int(str(item).strip()))
                    except Exception:
                        continue
            out[str(semana)] = sorted(set(sorteos))

    rangos_norm = _normalizar_rangos_area(rangos)
    for semana_key in rangos_norm.keys():
        out.setdefault(str(int(semana_key)), [])

    return dict(sorted(out.items(), key=lambda it: int(it[0])))


def _mergear_semanas_area(prev: dict | None, cur: dict | None) -> dict[str, list[int]]:
    merged: dict[str, set[int]] = {}
    for src in (prev, cur):
        if not isinstance(src, dict):
            continue
        for semana_key, sorteos_raw in src.items():
            try:
                semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
            except Exception:
                continue
            bucket = merged.setdefault(str(semana), set())
            if not isinstance(sorteos_raw, list):
                continue
            for item in sorteos_raw:
                try:
                    bucket.add(int(str(item).strip()))
                except Exception:
                    continue
    return {sem: sorted(vals) for sem, vals in sorted(merged.items(), key=lambda it: int(it[0]))}


def _mergear_rangos_area(prev: dict | None, cur: dict | None) -> dict[str, dict[str, str]]:
    merged = _normalizar_rangos_area(prev)
    merged.update(_normalizar_rangos_area(cur))
    return dict(sorted(merged.items(), key=lambda it: int(it[0])))


def _mergear_overrides_area(prev: dict | None, cur: dict | None) -> dict[str, dict[str, str]]:
    merged = _normalizar_manual_overrides_area(prev)
    for sorteo, cols_map in _normalizar_manual_overrides_area(cur).items():
        bucket = merged.setdefault(sorteo, {})
        bucket.update(cols_map)
    return merged


def _normalizar_planilla_prescripciones_data(values: dict) -> dict:
    if not isinstance(values, dict):
        return {}

    out: dict[str, dict[int, dict[str, dict[str, float | None]]]] = {}
    for juego, semanas_raw in values.items():
        juego_key = str(juego).strip()
        if not juego_key or not isinstance(semanas_raw, dict):
            continue

        semanas_out: dict[int, dict[str, dict[str, float | None]]] = {}
        for semana_key, rows_raw in semanas_raw.items():
            try:
                semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
            except Exception:
                continue
            if not isinstance(rows_raw, dict):
                continue

            rows_out: dict[str, dict[str, float | None]] = {}
            for sorteo_key, payload_raw in rows_raw.items():
                sorteo = str(sorteo_key).strip()
                if not sorteo or not isinstance(payload_raw, dict):
                    continue
                row_out = {}
                for campo in ("t_presc", "r_presc", "s_presc"):
                    value = payload_raw.get(campo)
                    if value in (None, ""):
                        row_out[campo] = None
                    else:
                        try:
                            row_out[campo] = float(value)
                        except Exception:
                            row_out[campo] = None
                rows_out[sorteo] = row_out
            semanas_out[semana] = rows_out

        out[juego_key] = semanas_out

    return out


def _normalizar_prescripciones_sorteos_por_semana(values: dict) -> dict:
    if not isinstance(values, dict):
        return {}

    out: dict[str, dict[int, list[int]]] = {}
    for juego, semanas_raw in values.items():
        juego_key = str(juego).strip()
        if not juego_key or not isinstance(semanas_raw, dict):
            continue

        semanas_out: dict[int, list[int]] = {}
        for semana_key, sorteos_raw in semanas_raw.items():
            try:
                semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
            except Exception:
                continue
            if not isinstance(sorteos_raw, list):
                continue

            sorteos: list[int] = []
            for item in sorteos_raw:
                try:
                    sorteos.append(int(str(item).strip()))
                except Exception:
                    continue

            semanas_out[semana] = sorted(set(sorteos))

        out[juego_key] = semanas_out

    return out


def _run_hooks(collection_name: str, *args):
    for hook in getattr(app_state, collection_name, {}).values():
        if callable(hook):
            try:
                hook(*args)
            except Exception:
                pass


def _refresh_tabs(reset_anticipos: bool = False):
    _run_hooks("planilla_area_reload_hooks")
    _run_hooks("planilla_refresh_hooks")
    _run_hooks("planilla_presc_refresh_hooks")
    _run_hooks("planilla_totales_refresh_hooks")
    _run_hooks("planilla_agencia_amiga_refresh_hooks")

    if reset_anticipos:
        _run_hooks("planilla_anticipos_reset_hooks")
        _run_hooks("planilla_control_cio_reset_hooks")


def _reset_estado_transitorio_memoria():
    app_state.sfa_bulk_publish = False
    if hasattr(app_state, "limpiar_sfa_resumen") and callable(app_state.limpiar_sfa_resumen):
        try:
            app_state.limpiar_sfa_resumen()
        except Exception:
            pass

    app_state.txt_resumen = {}
    app_state.json_resumen = {}
    app_state.facuni_resumen = []


def _normalizar_area_recaudacion(area_payload: dict) -> dict:
    if not isinstance(area_payload, dict):
        return {"version": 2, "generado": datetime.now().isoformat(timespec="seconds"), "planillas": {}}

    planillas_in = area_payload.get("planillas", {})
    if not isinstance(planillas_in, dict):
        planillas_in = {}

    planillas_out = {}
    for juego, data_juego in planillas_in.items():
        if not isinstance(data_juego, dict):
            continue

        filas = data_juego.get("filas", [])
        if not isinstance(filas, list):
            filas = []

        filas_out = []
        for fila in filas:
            if isinstance(fila, list):
                filas_out.append([str(v) for v in fila])

        semanas_raw = data_juego.get("semanas", {})
        semanas_norm: dict[str, list[int]] = {}
        if isinstance(semanas_raw, dict):
            for k, vals in semanas_raw.items():
                try:
                    sem_n = int(str(k).strip().lower().replace("semana", "").strip())
                except Exception:
                    continue
                if not isinstance(vals, list):
                    continue

                sorteos = []
                for v in vals:
                    try:
                        sorteos.append(int(v))
                    except Exception:
                        continue
                semanas_norm[str(sem_n)] = sorted(set(sorteos))

        try:
            semana_actual = int(data_juego.get("semana_actual", 0) or 0)
        except Exception:
            semana_actual = 0

        rangos_norm = _normalizar_rangos_area(data_juego.get("rangos_semana", {}))
        semanas_norm = _asegurar_claves_semanas_desde_rangos(semanas_norm, rangos_norm)

        planillas_out[str(juego)] = {
            "codigo_juego": data_juego.get("codigo_juego"),
            "columnas": list(data_juego.get("columnas", [])) if isinstance(data_juego.get("columnas", []), list) else [],
            "filas": filas_out,
            "semanas": semanas_norm,
            "rangos_semana": rangos_norm,
            "manual_overrides": _normalizar_manual_overrides_area(data_juego.get("manual_overrides", {})),
            "semana_actual": semana_actual,
            "semana_guardada": str(data_juego.get("semana_guardada", "Semana 1") or "Semana 1"),
        }

    return {
        "version": int(area_payload.get("version", 2) or 2),
        "generado": datetime.now().isoformat(timespec="seconds"),
        "planillas": planillas_out,
    }


def _snapshot_area_recaudacion() -> dict:
    guardar_area_hook = getattr(app_state, "planilla_area_guardar_hook", None)
    if callable(guardar_area_hook):
        try:
            guardar_area_hook()
        except Exception:
            pass

    area_payload = _leer_json(_ruta_guardado_planilla())

    snapshots = {}
    for juego, hook in getattr(app_state, "planilla_area_snapshot_hooks", {}).items():
        if callable(hook):
            try:
                data = hook()
            except Exception:
                data = None
            if isinstance(data, dict):
                snapshots[juego] = data

    if snapshots:
        planillas = area_payload.get("planillas", {}) if isinstance(area_payload, dict) else {}
        if not isinstance(planillas, dict):
            planillas = {}

        for juego, snap in snapshots.items():
            if not isinstance(snap, dict):
                continue

            prev = planillas.get(juego, {}) if isinstance(planillas.get(juego, {}), dict) else {}
            merged = dict(prev)
            merged.update(copy.deepcopy(snap))

            sem_snap = snap.get("semanas") if isinstance(snap, dict) else None
            sem_prev = prev.get("semanas") if isinstance(prev, dict) else None
            merged["semanas"] = _mergear_semanas_area(sem_prev, sem_snap)

            rangos_snap = snap.get("rangos_semana") if isinstance(snap, dict) else None
            rangos_prev = prev.get("rangos_semana") if isinstance(prev, dict) else None
            merged["rangos_semana"] = _mergear_rangos_area(rangos_prev, rangos_snap)
            merged["semanas"] = _asegurar_claves_semanas_desde_rangos(merged.get("semanas", {}), merged.get("rangos_semana", {}))

            overrides_snap = snap.get("manual_overrides") if isinstance(snap, dict) else None
            overrides_prev = prev.get("manual_overrides") if isinstance(prev, dict) else None
            merged["manual_overrides"] = _mergear_overrides_area(overrides_prev, overrides_snap)

            planillas[juego] = merged

        area_payload = {
            "version": area_payload.get("version", 2) if isinstance(area_payload, dict) else 2,
            "generado": datetime.now().isoformat(timespec="seconds"),
            "planillas": planillas,
        }

    return _normalizar_area_recaudacion(area_payload)


def _payload_actual() -> dict:
    area_payload = _snapshot_area_recaudacion()

    bundle_snapshots = {}
    for key, hook in getattr(app_state, "planilla_bundle_snapshot_hooks", {}).items():
        if callable(hook):
            try:
                data = hook()
            except Exception:
                data = None
            if isinstance(data, dict):
                bundle_snapshots[str(key)] = data

    anticipos_payload = copy.deepcopy(
        bundle_snapshots.get("anticipos_topes", getattr(app_state, "planilla_anticipos_topes_data", {}) or {})
    )
    control_cio_payload = copy.deepcopy(
        bundle_snapshots.get("control_cio", getattr(app_state, "planilla_control_cio_data", {}) or {})
    )
    totales_payload = copy.deepcopy(
        bundle_snapshots.get("totales", getattr(app_state, "planilla_totales_data", {}) or {})
    )
    agencia_amiga_payload = copy.deepcopy(
        bundle_snapshots.get("agencia_amiga", getattr(app_state, "planilla_agencia_amiga_data", {}) or {})
    )

    facuni_total_por_semana = copy.deepcopy(getattr(app_state, "facuni_total_por_semana", {}) or {})
    if not isinstance(facuni_total_por_semana, dict):
        facuni_total_por_semana = {}

    facuni_reporte_tobill_por_semana = copy.deepcopy(getattr(app_state, "facuni_reporte_tobill_por_semana", {}) or {})
    if not isinstance(facuni_reporte_tobill_por_semana, dict):
        facuni_reporte_tobill_por_semana = {}

    payload = {
        "version": 3,
        "generado": datetime.now().isoformat(timespec="seconds"),

        "area_recaudacion": area_payload,

        "tickets": {
            "tickets_resumen_por_juego": copy.deepcopy(getattr(app_state, "tickets_resumen_por_juego", {}) or {}),
            "tickets_resumen_por_semana_por_juego": copy.deepcopy(getattr(app_state, "tickets_resumen_por_semana_por_juego", {}) or {}),
            "tickets_prescripciones_por_semana_por_juego": copy.deepcopy(getattr(app_state, "tickets_prescripciones_por_semana_por_juego", {}) or {}),
        },

        "reporte": {
            "reporte_resumen_por_juego": copy.deepcopy(getattr(app_state, "reporte_resumen_por_juego", {}) or {}),
            "reporte_resumen_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_resumen_por_juego_por_semana", {}) or {}),
            "reporte_tobill_facuni_importe": float(getattr(app_state, "reporte_tobill_facuni_importe", 0.0) or 0.0),
            "reporte_tobill_facuni_cargado": bool(getattr(app_state, "reporte_tobill_facuni_cargado", False)),
            "reporte_tobill_sale_limit_por_juego": copy.deepcopy(getattr(app_state, "reporte_tobill_sale_limit_por_juego", {}) or {}),
            "reporte_tobill_sale_limit_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_tobill_sale_limit_por_juego_por_semana", {}) or {}),
            "reporte_tobill_advance_importe": float(getattr(app_state, "reporte_tobill_advance_importe", 0.0) or 0.0),
            "reporte_tobill_advance_importe_por_semana": copy.deepcopy(getattr(app_state, "reporte_tobill_advance_importe_por_semana", {}) or {}),
            "reporte_cash_in_importe": float(getattr(app_state, "reporte_cash_in_importe", 0.0) or 0.0),
            "reporte_cash_out_importe": float(getattr(app_state, "reporte_cash_out_importe", 0.0) or 0.0),
            "reporte_cash_in_importe_por_semana": copy.deepcopy(getattr(app_state, "reporte_cash_in_importe_por_semana", {}) or {}),
            "reporte_cash_out_importe_por_semana": copy.deepcopy(getattr(app_state, "reporte_cash_out_importe_por_semana", {}) or {}),
            "reporte_agencia_amiga_tobill_por_juego": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_tobill_por_juego", {}) or {}),
            "reporte_agencia_amiga_sfa_118_por_juego": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego", {}) or {}),
            "reporte_agencia_amiga_tobill_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_tobill_por_juego_por_semana", {}) or {}),
            "reporte_agencia_amiga_sfa_118_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}) or {}),
        },

        "prescripciones": {
            "reporte_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "reporte_prescripciones_por_juego", {}) or {}),
            "reporte_prescripciones_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_prescripciones_por_juego_por_semana", {}) or {}),
            "tickets_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "tickets_prescripciones_por_juego", {}) or {}),
            "sfa_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "sfa_prescripciones_por_juego", {}) or {}),
            "prescripciones_sorteos_por_semana_por_juego": copy.deepcopy(getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", {}) or {}),
            "planilla_prescripciones_data": copy.deepcopy(getattr(app_state, "planilla_prescripciones_data", {}) or {}),
        },

        "sfa": {
            "sfa_resumen_por_juego": copy.deepcopy(getattr(app_state, "sfa_resumen_por_juego", {}) or {}),
            "sfa_resumen_por_juego_por_semana": copy.deepcopy(getattr(app_state, "sfa_resumen_por_juego_por_semana", {}) or {}),
            "sfa_z118_por_juego": copy.deepcopy(getattr(app_state, "sfa_z118_por_juego", {}) or {}),
        },

        "facuni": {
            "facuni_total": float(getattr(app_state, "facuni_total", 0.0) or 0.0),
            "facuni_total_por_semana": facuni_total_por_semana,
            "facuni_reporte_tobill_por_semana": facuni_reporte_tobill_por_semana,
        },

        "filtro_area_recaudacion": {
            "actual": _normalizar_filtro_actual_para_bundle(
                getattr(app_state, "planilla_semana_filtro_actual", {}) or {},
                getattr(app_state, "planilla_rangos_semana_global", {}) or {},
            ),
            "rangos_semana_global": copy.deepcopy(getattr(app_state, "planilla_rangos_semana_global", {}) or {}),
        },

        "anticipos_topes": copy.deepcopy(anticipos_payload) if isinstance(anticipos_payload, dict) else {},
        "control_cio": copy.deepcopy(control_cio_payload) if isinstance(control_cio_payload, dict) else {},
        "totales": copy.deepcopy(totales_payload) if isinstance(totales_payload, dict) else {},
        "agencia_amiga": copy.deepcopy(agencia_amiga_payload) if isinstance(agencia_amiga_payload, dict) else {},
    }

    return payload


def _normalizar_filtro_actual_para_bundle(actual_raw, rangos_raw) -> dict:
    actual = copy.deepcopy(actual_raw) if isinstance(actual_raw, dict) else {}

    try:
        sem = int(actual.get("semana", 0) or 0)
    except Exception:
        try:
            sem = int(str(actual.get("semana", "0")).lower().replace("semana", "").strip() or 0)
        except Exception:
            sem = 0

    desde = str(actual.get("desde", "") or "").strip()
    hasta = str(actual.get("hasta", "") or "").strip()

    if sem < 0:
        sem = 0
    if sem > 5:
        sem = 5

    rango = None
    if isinstance(rangos_raw, dict):
        rango = rangos_raw.get(sem)
        if rango is None:
            rango = rangos_raw.get(str(sem))

    if isinstance(rango, dict):
        desde = desde or str(rango.get("desde", "") or "").strip()
        hasta = hasta or str(rango.get("hasta", "") or "").strip()
    elif isinstance(rango, (list, tuple)) and len(rango) >= 2:
        desde = desde or str(rango[0] or "").strip()
        hasta = hasta or str(rango[1] or "").strip()

    return {
        "juego": str(actual.get("juego", "") or "").strip(),
        "semana": sem,
        "desde": desde,
        "hasta": hasta,
    }


def guardar_bundle_default(estado_var=None):
    path = _ruta_bundle_default()
    try:
        payload = _payload_actual()
        _escribir_json(path, payload)
        if estado_var is not None:
            estado_var.set(f"Planilla Facturación guardada en: {path}")
    except Exception as e:
        if estado_var is not None:
            estado_var.set(f"Error al guardar Planilla Facturación: {e}")


def guardar_bundle_como(path: str, estado_var=None):
    if not path:
        return

    try:
        payload = _payload_actual()
        _escribir_json(path, payload)
        _escribir_json(_ruta_bundle_default(), payload)

        app_state.planilla_bundle_last_path = path
        _guardar_ruta_bundle(path)

        if estado_var is not None:
            estado_var.set(f"Planilla Facturación guardada en: {path}")
    except Exception as e:
        if estado_var is not None:
            estado_var.set(f"Error al guardar Planilla Facturación: {e}")


def cargar_bundle(path: str, estado_var=None, remember_path: bool = True) -> bool:
    if not path:
        return False

    payload = _leer_json(path)
    if not payload:
        if estado_var is not None:
            estado_var.set("No se pudo cargar el archivo seleccionado.")
        return False

    area_payload = payload.get("area_recaudacion", {}) if isinstance(payload, dict) else {}
    if isinstance(area_payload, dict):
        _escribir_json(_ruta_guardado_planilla(), area_payload)

    _reset_estado_transitorio_memoria()

    tickets_payload = payload.get("tickets", {}) if isinstance(payload, dict) else {}
    _replace_dict("tickets_resumen_por_juego", tickets_payload.get("tickets_resumen_por_juego", {}))
    _replace_dict("tickets_resumen_por_semana_por_juego", tickets_payload.get("tickets_resumen_por_semana_por_juego", {}))
    _replace_dict("tickets_prescripciones_por_semana_por_juego", tickets_payload.get("tickets_prescripciones_por_semana_por_juego", {}))

    reporte_payload = payload.get("reporte", {}) if isinstance(payload, dict) else {}
    _replace_dict("reporte_resumen_por_juego", reporte_payload.get("reporte_resumen_por_juego", {}))
    _replace_dict("reporte_resumen_por_juego_por_semana", reporte_payload.get("reporte_resumen_por_juego_por_semana", {}))
    _replace_dict("reporte_tobill_sale_limit_por_juego", reporte_payload.get("reporte_tobill_sale_limit_por_juego", {}))
    _replace_dict("reporte_tobill_sale_limit_por_juego_por_semana", reporte_payload.get("reporte_tobill_sale_limit_por_juego_por_semana", {}))
    app_state.reporte_tobill_facuni_importe = float(reporte_payload.get("reporte_tobill_facuni_importe", 0.0) or 0.0)
    app_state.reporte_tobill_facuni_cargado = bool(reporte_payload.get("reporte_tobill_facuni_cargado", False))
    app_state.reporte_tobill_advance_importe = float(reporte_payload.get("reporte_tobill_advance_importe", 0.0) or 0.0)
    _replace_dict("reporte_tobill_advance_importe_por_semana", reporte_payload.get("reporte_tobill_advance_importe_por_semana", {}))
    app_state.reporte_cash_in_importe = float(reporte_payload.get("reporte_cash_in_importe", 0.0) or 0.0)
    app_state.reporte_cash_out_importe = float(reporte_payload.get("reporte_cash_out_importe", 0.0) or 0.0)
    _replace_dict("reporte_cash_in_importe_por_semana", reporte_payload.get("reporte_cash_in_importe_por_semana", {}))
    _replace_dict("reporte_cash_out_importe_por_semana", reporte_payload.get("reporte_cash_out_importe_por_semana", {}))
    _replace_dict("reporte_agencia_amiga_tobill_por_juego", reporte_payload.get("reporte_agencia_amiga_tobill_por_juego", {}))
    _replace_dict("reporte_agencia_amiga_sfa_118_por_juego", reporte_payload.get("reporte_agencia_amiga_sfa_118_por_juego", {}))
    _replace_dict("reporte_agencia_amiga_tobill_por_juego_por_semana", reporte_payload.get("reporte_agencia_amiga_tobill_por_juego_por_semana", {}))
    _replace_dict("reporte_agencia_amiga_sfa_118_por_juego_por_semana", reporte_payload.get("reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}))

    presc = payload.get("prescripciones", {}) if isinstance(payload, dict) else {}
    _replace_dict("reporte_prescripciones_por_juego", presc.get("reporte_prescripciones_por_juego", {}))
    _replace_dict("reporte_prescripciones_por_juego_por_semana", presc.get("reporte_prescripciones_por_juego_por_semana", {}))
    _replace_dict("tickets_prescripciones_por_juego", presc.get("tickets_prescripciones_por_juego", {}))
    _replace_dict("sfa_prescripciones_por_juego", presc.get("sfa_prescripciones_por_juego", {}))
    _replace_dict(
        "prescripciones_sorteos_por_semana_por_juego",
        _normalizar_prescripciones_sorteos_por_semana(presc.get("prescripciones_sorteos_por_semana_por_juego", {})),
    )
    _replace_dict(
        "planilla_prescripciones_data",
        _normalizar_planilla_prescripciones_data(presc.get("planilla_prescripciones_data", {})),
    )

    sfa_payload = payload.get("sfa", {}) if isinstance(payload, dict) else {}
    _replace_dict("sfa_resumen_por_juego", sfa_payload.get("sfa_resumen_por_juego", {}))
    _replace_dict("sfa_resumen_por_juego_por_semana", sfa_payload.get("sfa_resumen_por_juego_por_semana", {}))
    _replace_dict("sfa_z118_por_juego", sfa_payload.get("sfa_z118_por_juego", {}))

    facuni_payload = payload.get("facuni", {}) if isinstance(payload, dict) else {}
    app_state.facuni_total = float((facuni_payload.get("facuni_total", 0.0) if isinstance(facuni_payload, dict) else 0.0) or 0.0)

    if not hasattr(app_state, "facuni_total_por_semana") or not isinstance(getattr(app_state, "facuni_total_por_semana", None), dict):
        app_state.facuni_total_por_semana = {}
    else:
        app_state.facuni_total_por_semana.clear()
    if isinstance(facuni_payload, dict):
        app_state.facuni_total_por_semana.update(copy.deepcopy(facuni_payload.get("facuni_total_por_semana", {}) or {}))

    if not hasattr(app_state, "facuni_reporte_tobill_por_semana") or not isinstance(getattr(app_state, "facuni_reporte_tobill_por_semana", None), dict):
        app_state.facuni_reporte_tobill_por_semana = {}
    else:
        app_state.facuni_reporte_tobill_por_semana.clear()
    if isinstance(facuni_payload, dict):
        app_state.facuni_reporte_tobill_por_semana.update(copy.deepcopy(facuni_payload.get("facuni_reporte_tobill_por_semana", {}) or {}))

    anticipos_payload = payload.get("anticipos_topes", {}) if isinstance(payload, dict) else {}
    app_state.planilla_anticipos_topes_data = copy.deepcopy(anticipos_payload) if isinstance(anticipos_payload, dict) else {}
    _run_hooks("planilla_bundle_load_hooks", app_state.planilla_anticipos_topes_data)

    control_cio_payload = payload.get("control_cio", {}) if isinstance(payload, dict) else {}
    app_state.planilla_control_cio_data = copy.deepcopy(control_cio_payload) if isinstance(control_cio_payload, dict) else {}
    _run_hooks("planilla_control_cio_load_hooks", app_state.planilla_control_cio_data)

    totales_payload = payload.get("totales", {}) if isinstance(payload, dict) else {}
    app_state.planilla_totales_data = copy.deepcopy(totales_payload) if isinstance(totales_payload, dict) else {}
    _run_hooks("planilla_bundle_load_hooks", app_state.planilla_totales_data)

    agencia_amiga_payload = payload.get("agencia_amiga", {}) if isinstance(payload, dict) else {}
    app_state.planilla_agencia_amiga_data = copy.deepcopy(agencia_amiga_payload) if isinstance(agencia_amiga_payload, dict) else {}
    _run_hooks("planilla_agencia_amiga_load_hooks", app_state.planilla_agencia_amiga_data)

    filtro_payload = payload.get("filtro_area_recaudacion", {}) if isinstance(payload, dict) else {}
    actual = filtro_payload.get("actual", {}) if isinstance(filtro_payload, dict) else {}
    rangos_raw = filtro_payload.get("rangos_semana_global", {}) if isinstance(filtro_payload, dict) else {}

    rangos_norm = {}
    if isinstance(rangos_raw, dict):
        for k, v in rangos_raw.items():
            try:
                sem = int(str(k).strip())
            except Exception:
                continue

            if isinstance(v, (list, tuple)) and len(v) >= 2:
                desde = str(v[0] or "").strip()
                hasta = str(v[1] or "").strip()
            elif isinstance(v, dict):
                desde = str(v.get("desde", "") or "").strip()
                hasta = str(v.get("hasta", "") or "").strip()
            else:
                continue

            rangos_norm[sem] = (desde, hasta)

    if hasattr(app_state, "planilla_rangos_semana_global"):
        app_state.planilla_rangos_semana_global.clear()
        app_state.planilla_rangos_semana_global.update(rangos_norm)

    actual_norm = _normalizar_filtro_actual_para_bundle(actual, rangos_norm)

    if hasattr(app_state, "planilla_semana_filtro_actual"):
        app_state.planilla_semana_filtro_actual.clear()
        app_state.planilla_semana_filtro_actual.update(actual_norm)

    _refresh_tabs(reset_anticipos=False)

    if hasattr(app_state, "publicar_filtro_area_recaudacion"):
        try:
            app_state.publicar_filtro_area_recaudacion(
                actual_norm.get("juego", ""),
                int(actual_norm.get("semana", 0) or 0),
                actual_norm.get("desde", ""),
                actual_norm.get("hasta", ""),
            )
        except Exception:
            pass

    if remember_path:
        app_state.planilla_bundle_last_path = path
        _guardar_ruta_bundle(path)

    if estado_var is not None:
        estado_var.set(f"Planilla Facturación cargada desde: {path}")
    return True


def cargar_bundle_default_si_existe(estado_var=None):
    ultimo_path = _leer_ruta_bundle_guardada()
    app_state.planilla_bundle_last_path = ultimo_path

    path = _ruta_bundle_default()
    if os.path.exists(path):
        cargar_bundle(path, estado_var, remember_path=False)


def limpiar_planilla_facturacion(estado_var=None):
    try:
        # Señal explícita para que Área Recaudación no conserve semanas/sorteos en memoria.
        app_state.planilla_force_clear_area = True
        ahora = datetime.now().isoformat(timespec="seconds")

        _escribir_json(
            _ruta_guardado_planilla(),
            {"version": 2, "generado": ahora, "planillas": {}},
        )

        _escribir_json(
            _ruta_bundle_default(),
            {
                "version": 3,
                "generado": ahora,
                "area_recaudacion": {
                    "version": 2,
                    "generado": ahora,
                    "planillas": {},
                },
                "tickets": {
                    "tickets_resumen_por_juego": {},
                    "tickets_resumen_por_semana_por_juego": {f"Semana {i}": {} for i in range(1, 6)},
                    "tickets_prescripciones_por_semana_por_juego": {f"Semana {i}": {} for i in range(1, 6)},
                },
                "reporte": {
                    "reporte_resumen_por_juego": {},
                    "reporte_resumen_por_juego_por_semana": {},
                    "reporte_tobill_facuni_importe": 0.0,
                    "reporte_tobill_facuni_cargado": False,
                    "reporte_tobill_sale_limit_por_juego": {},
                    "reporte_tobill_sale_limit_por_juego_por_semana": {},
                    "reporte_tobill_advance_importe": 0.0,
                    "reporte_tobill_advance_importe_por_semana": {},
                    "reporte_cash_in_importe": 0.0,
                    "reporte_cash_out_importe": 0.0,
                    "reporte_cash_in_importe_por_semana": {},
                    "reporte_cash_out_importe_por_semana": {},
                    "reporte_agencia_amiga_tobill_por_juego": {},
                    "reporte_agencia_amiga_sfa_118_por_juego": {},
                    "reporte_agencia_amiga_tobill_por_juego_por_semana": {},
                    "reporte_agencia_amiga_sfa_118_por_juego_por_semana": {},
                },
                "prescripciones": {
                    "reporte_prescripciones_por_juego": {},
                    "reporte_prescripciones_por_juego_por_semana": {},
                    "tickets_prescripciones_por_juego": {},
                    "sfa_prescripciones_por_juego": {},
                    "prescripciones_sorteos_por_semana_por_juego": {},
                    "planilla_prescripciones_data": {},
                },
                "sfa": {
                    "sfa_resumen_por_juego": {},
                    "sfa_resumen_por_juego_por_semana": {},
                    "sfa_z118_por_juego": {},
                },
                "facuni": {
                    "facuni_total": 0.0,
                    "facuni_total_por_semana": {},
                    "facuni_reporte_tobill_por_semana": {},
                },
                "filtro_area_recaudacion": {
                    "actual": {"juego": "", "semana": 0, "desde": "", "hasta": ""},
                    "rangos_semana_global": {},
                },
                "anticipos_topes": {},
                "control_cio": {},
                "agencia_amiga": {},
                "totales": {
                    "semanas": {f"Semana {i}": {} for i in range(1, 6)},
                    "manuales": {},
                },
            },
        )

        app_state.planilla_anticipos_topes_data = {}
        app_state.planilla_control_cio_data = {}
        app_state.planilla_prescripciones_data = {}
        app_state.planilla_totales_data = {
            "semanas": {f"Semana {i}": {} for i in range(1, 6)},
            "manuales": {},
        }
        app_state.planilla_bundle_last_path = ""
        _borrar_ruta_bundle_guardada()

        agencia_amiga_data = getattr(app_state, "planilla_agencia_amiga_data", None)
        if isinstance(agencia_amiga_data, dict):
            agencia_amiga_data.clear()
        else:
            app_state.planilla_agencia_amiga_data = {}

        app_state.reporte_tobill_facuni_importe = 0.0
        app_state.reporte_tobill_facuni_cargado = False
        app_state.reporte_tobill_advance_importe = 0.0
        app_state.reporte_cash_in_importe = 0.0
        app_state.reporte_cash_out_importe = 0.0
        app_state.facuni_total = 0.0

        if not hasattr(app_state, "facuni_total_por_semana") or not isinstance(getattr(app_state, "facuni_total_por_semana", None), dict):
            app_state.facuni_total_por_semana = {}
        else:
            app_state.facuni_total_por_semana.clear()

        if not hasattr(app_state, "facuni_reporte_tobill_por_semana") or not isinstance(getattr(app_state, "facuni_reporte_tobill_por_semana", None), dict):
            app_state.facuni_reporte_tobill_por_semana = {}
        else:
            app_state.facuni_reporte_tobill_por_semana.clear()

        if hasattr(app_state, "planilla_semana_filtro_actual"):
            app_state.planilla_semana_filtro_actual.clear()
            app_state.planilla_semana_filtro_actual.update({"juego": "", "semana": 0, "desde": "", "hasta": ""})

        if hasattr(app_state, "planilla_rangos_semana_global"):
            app_state.planilla_rangos_semana_global.clear()

        for name in (
            "tickets_resumen_por_juego",
            "tickets_resumen_por_semana_por_juego",
            "tickets_prescripciones_por_semana_por_juego",
            "reporte_resumen_por_juego",
            "reporte_resumen_por_juego_por_semana",
            "sfa_resumen_por_juego",
            "sfa_resumen_por_juego_por_semana",
            "sfa_z118_por_juego",
            "reporte_prescripciones_por_juego",
            "reporte_prescripciones_por_juego_por_semana",
            "tickets_prescripciones_por_juego",
            "sfa_prescripciones_por_juego",
            "prescripciones_sorteos_por_semana_por_juego",
            "reporte_tobill_sale_limit_por_juego",
            "reporte_tobill_sale_limit_por_juego_por_semana",
            "reporte_tobill_advance_importe_por_semana",
            "reporte_cash_in_importe_por_semana",
            "reporte_cash_out_importe_por_semana",
            "reporte_agencia_amiga_tobill_por_juego",
            "reporte_agencia_amiga_sfa_118_por_juego",
            "reporte_agencia_amiga_tobill_por_juego_por_semana",
            "reporte_agencia_amiga_sfa_118_por_juego_por_semana",
        ):
            _replace_dict(name, {})

        # Limpiar estado transitorio de otras solapas (SFA/Tickets/Reporte/FACUNI).
        _reset_estado_transitorio_memoria()
        app_state.facuni_total_pendiente = 0.0
        app_state.facuni_ultima_semana_importada = ""

        _run_hooks("planilla_global_reset_hooks")
        _refresh_tabs(reset_anticipos=True)
        _run_hooks("planilla_global_reset_hooks")
        if hasattr(app_state, "publicar_filtro_area_recaudacion"):
            try:
                app_state.publicar_filtro_area_recaudacion("", 0, "", "")
            except Exception:
                pass
        _run_hooks("planilla_totales_refresh_hooks")
        app_state.planilla_force_clear_area = False

        if estado_var is not None:
            estado_var.set("Planilla Facturación limpia (todas las secciones).")
    except Exception as e:
        app_state.planilla_force_clear_area = False
        if estado_var is not None:
            estado_var.set(f"Error al limpiar Planilla Facturación: {e}")
