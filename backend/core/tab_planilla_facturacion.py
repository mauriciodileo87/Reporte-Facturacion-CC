# tabs/tab_planilla_facturacion.py
from __future__ import annotations

import copy
import json
import importlib.util
import os
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog

import app_state
from tabs.planilla_anticipos_topes import build_anticipos_topes
from tabs.planilla_area_recaudacion import build_area_recaudacion
from tabs.planilla_area_recaudacion import _ruta_guardado_planilla
from tabs.planilla_prescripciones import build_prescripciones
from tabs.planilla_control_cio import build_control_cio
from tabs.planilla_agencia_amiga import build_agencia_amiga
from tabs.planilla_totales import build_totales


def _storage_no_disponible(estado_var=None):
    if estado_var is not None:
        estado_var.set("Storage de planilla no disponible. Verificá tabs/planilla_facturacion_storage.py")




def _normalizar_area_recaudacion_para_semana_nueva_fallback(area_payload: dict) -> dict:
    if not isinstance(area_payload, dict):
        return {}

    planillas_in = area_payload.get("planillas", {})
    if not isinstance(planillas_in, dict):
        return copy.deepcopy(area_payload)

    planillas_out = {}
    for juego, data_juego in planillas_in.items():
        if not isinstance(data_juego, dict):
            continue

        filas = data_juego.get("filas", [])
        if not isinstance(filas, list):
            filas = []

        filas_con_datos = []
        for fila in filas:
            if isinstance(fila, list) and any(str(v).strip() for v in fila):
                filas_con_datos.append([str(v) for v in fila])

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
                sorteos: list[int] = []
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

        rangos_raw = data_juego.get("rangos_semana", {}) if isinstance(data_juego, dict) else {}
        rangos_norm = {}
        if isinstance(rangos_raw, dict):
            for rk, rv in rangos_raw.items():
                try:
                    rsem_n = int(str(rk).strip().lower().replace("semana", "").strip())
                except Exception:
                    continue
                if isinstance(rv, dict):
                    desde = str(rv.get("desde", "") or "").strip()
                    hasta = str(rv.get("hasta", "") or "").strip()
                elif isinstance(rv, (list, tuple)) and len(rv) >= 2:
                    desde = str(rv[0] or "").strip()
                    hasta = str(rv[1] or "").strip()
                else:
                    continue
                if desde and hasta:
                    rangos_norm[str(rsem_n)] = {"desde": desde, "hasta": hasta}

        overrides_raw = data_juego.get("manual_overrides", {}) if isinstance(data_juego, dict) else {}
        overrides_norm = {}
        if isinstance(overrides_raw, dict):
            for sorteo_key, cols_raw in overrides_raw.items():
                sorteo = str(sorteo_key or "").strip()
                if not sorteo or not isinstance(cols_raw, dict):
                    continue
                cols_out = {}
                for col_key, value in cols_raw.items():
                    try:
                        col_idx = int(str(col_key).strip())
                    except Exception:
                        continue
                    if col_idx <= 0:
                        continue
                    cols_out[str(col_idx)] = str(value or "")
                if cols_out:
                    overrides_norm[sorteo] = cols_out

        planillas_out[juego] = {
            "codigo_juego": data_juego.get("codigo_juego"),
            "columnas": list(data_juego.get("columnas", [])) if isinstance(data_juego.get("columnas", []), list) else [],
            "filas": filas_con_datos,
            "semanas": semanas_norm,
            "rangos_semana": rangos_norm,
            "manual_overrides": overrides_norm,
            "semana_actual": semana_actual,
            "semana_guardada": str(data_juego.get("semana_guardada", "Semana 1") or "Semana 1"),
        }

    return {
        "version": int(area_payload.get("version", 2) or 2),
        "generado": datetime.now().isoformat(timespec="seconds"),
        "planillas": planillas_out,
    }


def guardar_bundle_default(estado_var=None):
    _storage_no_disponible(estado_var)


def guardar_bundle_como(path: str, estado_var=None):
    _ = path
    _storage_no_disponible(estado_var)


def cargar_bundle(path: str, estado_var=None):
    _ = path
    _storage_no_disponible(estado_var)
    return False


def cargar_bundle_default_si_existe(estado_var=None):
    _storage_no_disponible(estado_var)


def limpiar_planilla_facturacion(estado_var=None):
    _storage_no_disponible(estado_var)


_STORAGE_DISPONIBLE = False
_storage_error = ""

try:
    # Import estático para que PyInstaller incluya este módulo en el ejecutable.
    from tabs import planilla_facturacion_storage as _storage

    _STORAGE_DISPONIBLE = True
except Exception as e_moderno:
    _storage_error = f"Storage moderno: {e_moderno}"

if _STORAGE_DISPONIBLE:
    required_funcs = (
        "guardar_bundle_default",
        "guardar_bundle_como",
        "cargar_bundle",
        "cargar_bundle_default_si_existe",
        "limpiar_planilla_facturacion",
    )
    faltantes = [name for name in required_funcs if not hasattr(_storage, name)]
    if faltantes:
        _STORAGE_DISPONIBLE = False
        detalle = ", ".join(faltantes)
        _storage_error = (
            (_storage_error + " | ") if _storage_error else ""
        ) + f"Storage legado incompleto: faltan funciones: {detalle}"
    else:
        guardar_bundle_default = _storage.guardar_bundle_default
        guardar_bundle_como = _storage.guardar_bundle_como
        cargar_bundle = _storage.cargar_bundle
        cargar_bundle_default_si_existe = _storage.cargar_bundle_default_si_existe
        limpiar_planilla_facturacion = _storage.limpiar_planilla_facturacion
if not _STORAGE_DISPONIBLE:
    # Fallback robusto: permite guardar/cargar aunque falle el import del módulo externo.
    def _ruta_bundle_default_fallback() -> str:
        appdata = os.environ.get("APPDATA")
        if appdata:
            base = os.path.join(appdata, "ReporteFacturacion")
        else:
            base = os.path.join(os.path.expanduser("~"), ".reporte_facturacion")
        os.makedirs(base, exist_ok=True)
        return os.path.join(base, "planilla_facturacion_bundle.json")


    def _leer_json_fallback(path: str) -> dict:
        if not os.path.exists(path):
            return {}
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return data if isinstance(data, dict) else {}
        except Exception:
            return {}


    def _escribir_json_fallback(path: str, payload: dict):
        folder = os.path.dirname(path) or "."
        os.makedirs(folder, exist_ok=True)
        tmp_path = f"{path}.tmp"
        with open(tmp_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        os.replace(tmp_path, path)


    def _ruta_bundle_last_path_file_fallback() -> str:
        base = os.path.dirname(_ruta_bundle_default_fallback())
        os.makedirs(base, exist_ok=True)
        return os.path.join(base, "planilla_facturacion_last_path.txt")


    def _guardar_ruta_bundle_fallback(path: str):
        p = os.path.abspath(str(path or "").strip())
        if not p:
            return
        try:
            path_file = _ruta_bundle_last_path_file_fallback()
            tmp_file = f"{path_file}.tmp"
            with open(tmp_file, "w", encoding="utf-8") as f:
                f.write(p)
            os.replace(tmp_file, path_file)
        except Exception:
            pass


    def _leer_ruta_bundle_guardada_fallback() -> str:
        path_file = _ruta_bundle_last_path_file_fallback()
        if not os.path.exists(path_file):
            return ""
        try:
            with open(path_file, "r", encoding="utf-8") as f:
                p = str(f.read() or "").strip()
        except Exception:
            return ""
        return p if p and os.path.exists(p) else ""


    def _payload_actual_fallback() -> dict:
        guardar_area_hook = getattr(app_state, "planilla_area_guardar_hook", None)
        if callable(guardar_area_hook):
            try:
                guardar_area_hook()
            except Exception:
                pass
        area_payload = _leer_json_fallback(_ruta_guardado_planilla())
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
                merged.update(snap)

                def _merge_semanas(prev_map, cur_map):
                    merged_map = {}
                    for src in (prev_map, cur_map):
                        if not isinstance(src, dict):
                            continue
                        for sem_key, vals in src.items():
                            try:
                                sem_n = int(str(sem_key).strip().lower().replace("semana", "").strip())
                            except Exception:
                                continue
                            bucket = merged_map.setdefault(str(sem_n), set())
                            if not isinstance(vals, list):
                                continue
                            for item in vals:
                                try:
                                    bucket.add(int(str(item).strip()))
                                except Exception:
                                    continue
                    return {k: sorted(v) for k, v in sorted(merged_map.items(), key=lambda it: int(it[0]))}

                def _merge_rangos(prev_map, cur_map):
                    out = {}
                    for src in (prev_map, cur_map):
                        if not isinstance(src, dict):
                            continue
                        for sem_key, raw in src.items():
                            try:
                                sem_n = int(str(sem_key).strip().lower().replace("semana", "").strip())
                            except Exception:
                                continue
                            if isinstance(raw, dict):
                                desde = str(raw.get("desde", "") or "").strip()
                                hasta = str(raw.get("hasta", "") or "").strip()
                            elif isinstance(raw, (list, tuple)) and len(raw) >= 2:
                                desde = str(raw[0] or "").strip()
                                hasta = str(raw[1] or "").strip()
                            else:
                                continue
                            if desde and hasta:
                                out[str(sem_n)] = {"desde": desde, "hasta": hasta}
                    return dict(sorted(out.items(), key=lambda it: int(it[0])))

                def _merge_overrides(prev_map, cur_map):
                    out = {}
                    for src in (prev_map, cur_map):
                        if not isinstance(src, dict):
                            continue
                        for sorteo, cols_raw in src.items():
                            key = str(sorteo or "").strip()
                            if not key or not isinstance(cols_raw, dict):
                                continue
                            bucket = out.setdefault(key, {})
                            for col_key, value in cols_raw.items():
                                try:
                                    col_idx = int(str(col_key).strip())
                                except Exception:
                                    continue
                                if col_idx <= 0:
                                    continue
                                bucket[str(col_idx)] = str(value or "")
                    return out

                merged["semanas"] = _merge_semanas(prev.get("semanas"), snap.get("semanas"))
                merged["rangos_semana"] = _merge_rangos(prev.get("rangos_semana"), snap.get("rangos_semana"))
                merged["manual_overrides"] = _merge_overrides(prev.get("manual_overrides"), snap.get("manual_overrides"))

                planillas[juego] = merged

            area_payload = {
                "version": area_payload.get("version", 2) if isinstance(area_payload, dict) else 2,
                "generado": datetime.now().isoformat(timespec="seconds"),
                "planillas": planillas,
            }

        area_payload = _normalizar_area_recaudacion_para_semana_nueva_fallback(area_payload)

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

        return {
            "version": 3,
            "generado": datetime.now().isoformat(timespec="seconds"),
            "area_recaudacion": area_payload,
            "tickets": {
                "tickets_resumen_por_juego": copy.deepcopy(getattr(app_state, "tickets_resumen_por_juego", {}) or {}),
            },
            "reporte": {
                "reporte_resumen_por_juego": copy.deepcopy(getattr(app_state, "reporte_resumen_por_juego", {}) or {}),
                "reporte_tobill_facuni_importe": float(getattr(app_state, "reporte_tobill_facuni_importe", 0.0) or 0.0),
                "reporte_tobill_facuni_cargado": bool(getattr(app_state, "reporte_tobill_facuni_cargado", False)),
                "reporte_tobill_sale_limit_por_juego": copy.deepcopy(getattr(app_state, "reporte_tobill_sale_limit_por_juego", {}) or {}),
                "reporte_tobill_sale_limit_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_tobill_sale_limit_por_juego_por_semana", {}) or {}),
                "reporte_tobill_advance_importe": float(getattr(app_state, "reporte_tobill_advance_importe", 0.0) or 0.0),
                "reporte_tobill_advance_importe_por_semana": copy.deepcopy(getattr(app_state, "reporte_tobill_advance_importe_por_semana", {}) or {}),
                "reporte_cash_in_importe": float(getattr(app_state, "reporte_cash_in_importe", 0.0) or 0.0),
                "reporte_cash_out_importe": float(getattr(app_state, "reporte_cash_out_importe", 0.0) or 0.0),
                "reporte_agencia_amiga_tobill_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_tobill_por_juego_por_semana", {}) or {}),
                "reporte_agencia_amiga_sfa_118_por_juego_por_semana": copy.deepcopy(getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}) or {}),
            },
            "prescripciones": {
                "reporte_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "reporte_prescripciones_por_juego", {}) or {}),
                "tickets_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "tickets_prescripciones_por_juego", {}) or {}),
                "sfa_prescripciones_por_juego": copy.deepcopy(getattr(app_state, "sfa_prescripciones_por_juego", {}) or {}),
                "prescripciones_sorteos_por_semana_por_juego": copy.deepcopy(getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", {}) or {}),
                "planilla_prescripciones_data": copy.deepcopy(getattr(app_state, "planilla_prescripciones_data", {}) or {}),
            },
            "sfa": {
                "sfa_resumen_por_juego": copy.deepcopy(getattr(app_state, "sfa_resumen_por_juego", {}) or {}),
                "sfa_z118_por_juego": copy.deepcopy(getattr(app_state, "sfa_z118_por_juego", {}) or {}),
            },
            "facuni": {
                "facuni_total": float(getattr(app_state, "facuni_total", 0.0) or 0.0),
                "facuni_total_por_semana": copy.deepcopy(getattr(app_state, "facuni_total_por_semana", {}) or {}),
                "facuni_reporte_tobill_por_semana": copy.deepcopy(getattr(app_state, "facuni_reporte_tobill_por_semana", {}) or {}),
            },
            "filtro_area_recaudacion": {
                "actual": copy.deepcopy(getattr(app_state, "planilla_semana_filtro_actual", {}) or {}),
                "rangos_semana_global": copy.deepcopy(getattr(app_state, "planilla_rangos_semana_global", {}) or {}),
            },
            "anticipos_topes": copy.deepcopy(anticipos_payload) if isinstance(anticipos_payload, dict) else {},
            "control_cio": copy.deepcopy(control_cio_payload) if isinstance(control_cio_payload, dict) else {},
            "totales": copy.deepcopy(totales_payload) if isinstance(totales_payload, dict) else {},
            "agencia_amiga": copy.deepcopy(agencia_amiga_payload) if isinstance(agencia_amiga_payload, dict) else {},
        }


    def _replace_dict_fallback(name: str, values: dict):
        target = getattr(app_state, name, None)
        if isinstance(target, dict):
            target.clear()
            if isinstance(values, dict):
                target.update(values)


    def _normalizar_planilla_prescripciones_data_fallback(values: dict) -> dict:
        if not isinstance(values, dict):
            return {}

        out = {}
        for juego, semanas_raw in values.items():
            juego_key = str(juego).strip()
            if not juego_key or not isinstance(semanas_raw, dict):
                continue

            semanas_out = {}
            for semana_key, rows_raw in semanas_raw.items():
                try:
                    semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
                except Exception:
                    continue
                if not isinstance(rows_raw, dict):
                    continue

                rows_out = {}
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


    def _normalizar_prescripciones_sorteos_por_semana_fallback(values: dict) -> dict:
        if not isinstance(values, dict):
            return {}

        out = {}
        for juego, semanas_raw in values.items():
            juego_key = str(juego).strip()
            if not juego_key or not isinstance(semanas_raw, dict):
                continue

            semanas_out = {}
            for semana_key, sorteos_raw in semanas_raw.items():
                try:
                    semana = int(str(semana_key).strip().lower().replace("semana", "").strip())
                except Exception:
                    continue
                if not isinstance(sorteos_raw, list):
                    continue

                sorteos = []
                for item in sorteos_raw:
                    try:
                        sorteos.append(int(str(item).strip()))
                    except Exception:
                        continue

                semanas_out[semana] = sorted(set(sorteos))

            out[juego_key] = semanas_out

        return out


    def _refresh_tabs_fallback(reset_anticipos: bool = False):
        for hook in getattr(app_state, "planilla_area_reload_hooks", {}).values():
            if callable(hook):
                hook()
        for hook in getattr(app_state, "planilla_refresh_hooks", {}).values():
            if callable(hook):
                hook()
        for hook in getattr(app_state, "planilla_presc_refresh_hooks", {}).values():
            if callable(hook):
                hook()
        for hook in getattr(app_state, "planilla_totales_refresh_hooks", {}).values():
            if callable(hook):
                hook()
        for hook in getattr(app_state, "planilla_agencia_amiga_refresh_hooks", {}).values():
            if callable(hook):
                hook()
        if reset_anticipos:
            for hook in getattr(app_state, "planilla_anticipos_reset_hooks", {}).values():
                if callable(hook):
                    hook()
            for hook in getattr(app_state, "planilla_control_cio_reset_hooks", {}).values():
                if callable(hook):
                    hook()


    def guardar_bundle_default(estado_var=None):
        """
        Guardado automático/default:
        SIEMPRE escribe en la ruta de autosave de la app.
        Nunca debe sobrescribir un archivo elegido por 'Guardar como...'.
        """
        path = _ruta_bundle_default_fallback()
        payload = _payload_actual_fallback()
        _escribir_json_fallback(path, payload)
        if estado_var is not None:
            estado_var.set(f"Planilla Facturación guardada en: {path}")


    def guardar_bundle_como(path: str, estado_var=None):
        if not path:
            return
        payload = _payload_actual_fallback()
        _escribir_json_fallback(path, payload)
        _escribir_json_fallback(_ruta_bundle_default_fallback(), payload)
        app_state.planilla_bundle_last_path = path
        _guardar_ruta_bundle_fallback(path)
        if estado_var is not None:
            estado_var.set(f"Planilla Facturación guardada en: {path}")


    def cargar_bundle(path: str, estado_var=None, remember_path: bool = True):
        if not path:
            return False
        payload = _leer_json_fallback(path)
        if not payload:
            if estado_var is not None:
                estado_var.set("No se pudo cargar el archivo seleccionado.")
            return False

        area_payload = payload.get("area_recaudacion", {}) if isinstance(payload, dict) else {}
        if isinstance(area_payload, dict):
            _escribir_json_fallback(_ruta_guardado_planilla(), area_payload)

        tickets_payload = payload.get("tickets", {}) if isinstance(payload, dict) else {}
        _replace_dict_fallback("tickets_resumen_por_juego", tickets_payload.get("tickets_resumen_por_juego", {}))

        reporte_payload = payload.get("reporte", {}) if isinstance(payload, dict) else {}
        _replace_dict_fallback("reporte_resumen_por_juego", reporte_payload.get("reporte_resumen_por_juego", {}))
        app_state.reporte_tobill_facuni_importe = float((reporte_payload.get("reporte_tobill_facuni_importe", 0.0) if isinstance(reporte_payload, dict) else 0.0) or 0.0)
        app_state.reporte_tobill_facuni_cargado = bool(reporte_payload.get("reporte_tobill_facuni_cargado", False)) if isinstance(reporte_payload, dict) else False
        _replace_dict_fallback("reporte_tobill_sale_limit_por_juego", reporte_payload.get("reporte_tobill_sale_limit_por_juego", {}))
        _replace_dict_fallback("reporte_tobill_sale_limit_por_juego_por_semana", reporte_payload.get("reporte_tobill_sale_limit_por_juego_por_semana", {}))
        app_state.reporte_tobill_advance_importe = float((reporte_payload.get("reporte_tobill_advance_importe", 0.0) if isinstance(reporte_payload, dict) else 0.0) or 0.0)
        _replace_dict_fallback("reporte_tobill_advance_importe_por_semana", reporte_payload.get("reporte_tobill_advance_importe_por_semana", {}))
        app_state.reporte_cash_in_importe = float((reporte_payload.get("reporte_cash_in_importe", 0.0) if isinstance(reporte_payload, dict) else 0.0) or 0.0)
        app_state.reporte_cash_out_importe = float((reporte_payload.get("reporte_cash_out_importe", 0.0) if isinstance(reporte_payload, dict) else 0.0) or 0.0)
        _replace_dict_fallback("reporte_agencia_amiga_tobill_por_juego_por_semana", reporte_payload.get("reporte_agencia_amiga_tobill_por_juego_por_semana", {}))
        _replace_dict_fallback("reporte_agencia_amiga_sfa_118_por_juego_por_semana", reporte_payload.get("reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}))

        presc = payload.get("prescripciones", {}) if isinstance(payload, dict) else {}
        _replace_dict_fallback("reporte_prescripciones_por_juego", presc.get("reporte_prescripciones_por_juego", {}))
        _replace_dict_fallback("tickets_prescripciones_por_juego", presc.get("tickets_prescripciones_por_juego", {}))
        _replace_dict_fallback("sfa_prescripciones_por_juego", presc.get("sfa_prescripciones_por_juego", {}))
        _replace_dict_fallback("prescripciones_sorteos_por_semana_por_juego", _normalizar_prescripciones_sorteos_por_semana_fallback(presc.get("prescripciones_sorteos_por_semana_por_juego", {})))
        _replace_dict_fallback("planilla_prescripciones_data", _normalizar_planilla_prescripciones_data_fallback(presc.get("planilla_prescripciones_data", {})))

        sfa_payload = payload.get("sfa", {}) if isinstance(payload, dict) else {}
        _replace_dict_fallback("sfa_resumen_por_juego", sfa_payload.get("sfa_resumen_por_juego", {}))
        _replace_dict_fallback("sfa_z118_por_juego", sfa_payload.get("sfa_z118_por_juego", {}))

        facuni_payload = payload.get("facuni", {}) if isinstance(payload, dict) else {}
        app_state.facuni_total = float((facuni_payload.get("facuni_total", 0.0) if isinstance(facuni_payload, dict) else 0.0) or 0.0)
        if isinstance(facuni_payload, dict) and hasattr(app_state, "facuni_total_por_semana"):
            app_state.facuni_total_por_semana.clear()
            app_state.facuni_total_por_semana.update(facuni_payload.get("facuni_total_por_semana", {}) or {})
        if isinstance(facuni_payload, dict) and hasattr(app_state, "facuni_reporte_tobill_por_semana"):
            app_state.facuni_reporte_tobill_por_semana.clear()
            app_state.facuni_reporte_tobill_por_semana.update(facuni_payload.get("facuni_reporte_tobill_por_semana", {}) or {})

        anticipos_payload = payload.get("anticipos_topes", {}) if isinstance(payload, dict) else {}
        app_state.planilla_anticipos_topes_data = copy.deepcopy(anticipos_payload) if isinstance(anticipos_payload, dict) else {}
        for _, hook in getattr(app_state, "planilla_bundle_load_hooks", {}).items():
            if callable(hook):
                try:
                    hook(app_state.planilla_anticipos_topes_data)
                except Exception:
                    pass

        control_cio_payload = payload.get("control_cio", {}) if isinstance(payload, dict) else {}
        app_state.planilla_control_cio_data = copy.deepcopy(control_cio_payload) if isinstance(control_cio_payload, dict) else {}
        for _, hook in getattr(app_state, "planilla_control_cio_load_hooks", {}).items():
            if callable(hook):
                try:
                    hook(app_state.planilla_control_cio_data)
                except Exception:
                    pass

        totales_payload = payload.get("totales", {}) if isinstance(payload, dict) else {}
        app_state.planilla_totales_data = copy.deepcopy(totales_payload) if isinstance(totales_payload, dict) else {}

        agencia_amiga_payload = payload.get("agencia_amiga", {}) if isinstance(payload, dict) else {}
        app_state.planilla_agencia_amiga_data = copy.deepcopy(agencia_amiga_payload) if isinstance(agencia_amiga_payload, dict) else {}
        for _, hook in getattr(app_state, "planilla_agencia_amiga_load_hooks", {}).items():
            if callable(hook):
                try:
                    hook(app_state.planilla_agencia_amiga_data)
                except Exception:
                    pass

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

        if hasattr(app_state, "planilla_semana_filtro_actual"):
            app_state.planilla_semana_filtro_actual.clear()
            if isinstance(actual, dict):
                app_state.planilla_semana_filtro_actual.update(actual)

        _refresh_tabs_fallback(reset_anticipos=False)

        if isinstance(actual, dict) and hasattr(app_state, "publicar_filtro_area_recaudacion"):
            try:
                app_state.publicar_filtro_area_recaudacion(
                    actual.get("juego", ""),
                    int(actual.get("semana", 1) or 1),
                    actual.get("desde", ""),
                    actual.get("hasta", ""),
                )
            except Exception:
                pass
        if remember_path:
            app_state.planilla_bundle_last_path = path
            _guardar_ruta_bundle_fallback(path)
        if estado_var is not None:
            estado_var.set(f"Planilla Facturación cargada desde: {path}")
        return True


    def cargar_bundle_default_si_existe(estado_var=None):
        """
        Al iniciar la app, solo cargar el autosave interno.
        No cargar automáticamente el archivo de 'Guardar como...',
        pero sí recordar la última ruta manual para que "Guardar planilla"
        sobrescriba ese JSON cuando corresponda.
        """
        app_state.planilla_bundle_last_path = _leer_ruta_bundle_guardada_fallback()

        path = _ruta_bundle_default_fallback()
        if os.path.exists(path):
            cargar_bundle(path, estado_var, remember_path=False)


    def limpiar_planilla_facturacion(estado_var=None):
        _escribir_json_fallback(
            _ruta_guardado_planilla(),
            {"version": 2, "generado": datetime.now().isoformat(timespec="seconds"), "planillas": {}},
        )
        _escribir_json_fallback(
            _ruta_bundle_default_fallback(),
            {
                "version": 3,
                "generado": datetime.now().isoformat(timespec="seconds"),
                "area_recaudacion": {"version": 2, "generado": datetime.now().isoformat(timespec="seconds"), "planillas": {}},
                "tickets": {"tickets_resumen_por_juego": {}},
                "reporte": {
                    "reporte_resumen_por_juego": {},
                    "reporte_tobill_facuni_importe": 0.0,
                    "reporte_tobill_facuni_cargado": False,
                    "reporte_tobill_sale_limit_por_juego": {},
                    "reporte_tobill_sale_limit_por_juego_por_semana": {},
                    "reporte_tobill_advance_importe": 0.0,
                    "reporte_tobill_advance_importe_por_semana": {},
                    "reporte_cash_in_importe": 0.0,
                    "reporte_cash_out_importe": 0.0,
                },
                "prescripciones": {
                    "reporte_prescripciones_por_juego": {},
                    "tickets_prescripciones_por_juego": {},
                    "sfa_prescripciones_por_juego": {},
                    "prescripciones_sorteos_por_semana_por_juego": {},
                    "planilla_prescripciones_data": {},
                },
                "sfa": {
                    "sfa_resumen_por_juego": {},
                    "sfa_z118_por_juego": {},
                },
                "facuni": {
                    "facuni_total": 0.0,
                    "facuni_total_por_semana": {},
                    "facuni_reporte_tobill_por_semana": {},
                },
                "filtro_area_recaudacion": {
                    "actual": {"juego": "", "semana": 1, "desde": "", "hasta": ""},
                    "rangos_semana_global": {},
                },
                "anticipos_topes": {},
                "control_cio": {},
                "agencia_amiga": {},
                "totales": {"semanas": {}},
            },
        )
        app_state.planilla_anticipos_topes_data = {}
        app_state.planilla_control_cio_data = {}
        app_state.planilla_agencia_amiga_data = {}
        app_state.planilla_totales_data = {}
        app_state.reporte_cash_in_importe = 0.0
        app_state.reporte_cash_out_importe = 0.0
        app_state.facuni_total = 0.0
        _replace_dict_fallback("reporte_tobill_sale_limit_por_juego", {})
        _replace_dict_fallback("reporte_tobill_sale_limit_por_juego_por_semana", {})
        _replace_dict_fallback("reporte_tobill_advance_importe_por_semana", {})
        if hasattr(app_state, "facuni_total_por_semana"):
            app_state.facuni_total_por_semana.clear()
            app_state.facuni_total_por_semana.update({
                "Semana 1": 0.0,
                "Semana 2": 0.0,
                "Semana 3": 0.0,
                "Semana 4": 0.0,
                "Semana 5": 0.0,
            })
        if hasattr(app_state, "facuni_reporte_tobill_por_semana"):
            app_state.facuni_reporte_tobill_por_semana.clear()
            app_state.facuni_reporte_tobill_por_semana.update({
                "Semana 1": 0.0,
                "Semana 2": 0.0,
                "Semana 3": 0.0,
                "Semana 4": 0.0,
                "Semana 5": 0.0,
            })
        if hasattr(app_state, "planilla_semana_filtro_actual"):
            app_state.planilla_semana_filtro_actual.clear()
            app_state.planilla_semana_filtro_actual.update({"juego": "", "semana": 1, "desde": "", "hasta": ""})
        if hasattr(app_state, "planilla_rangos_semana_global"):
            app_state.planilla_rangos_semana_global.clear()

        for name in (
            "tickets_resumen_por_juego",
            "reporte_resumen_por_juego",
            "sfa_resumen_por_juego",
            "sfa_z118_por_juego",
            "reporte_prescripciones_por_juego",
            "tickets_prescripciones_por_juego",
            "sfa_prescripciones_por_juego",
            "prescripciones_sorteos_por_semana_por_juego",
            "reporte_agencia_amiga_tobill_por_juego",
            "reporte_agencia_amiga_sfa_118_por_juego",
            "reporte_agencia_amiga_tobill_por_juego_por_semana",
            "reporte_agencia_amiga_sfa_118_por_juego_por_semana",
        ):
            _replace_dict_fallback(name, {})

        _refresh_tabs_fallback(reset_anticipos=True)
        if estado_var is not None:
            estado_var.set("Planilla Facturación limpia (fallback).")


def _crear_vista_filtrable(parent, titulo: str, opciones: list[str], on_build_frame, estado_var):
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(1, weight=1)
    try:
        parent.grid_rowconfigure(0, minsize=42)
    except Exception:
        pass

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

    top = ttk.Frame(parent, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 10))
    top.columnconfigure(2, weight=1)
    top.columnconfigure(3, minsize=180)

    ttk.Label(top, text=titulo, style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")
    combo = ttk.Combobox(top, state="readonly", values=opciones, width=30, style=filtro_style)
    combo.grid(row=0, column=1, sticky="w", padx=(8, 10))
    app_state.planilla_seccion_combo_widget = combo
    app_state.planilla_toolbar_combobox_style = filtro_style
    app_state.planilla_toolbar_combobox_width = 30
    app_state.planilla_seccion_rango_del_al_var = None
    app_state.planilla_actualizar_rango_del_al_seccion = lambda *args, **kwargs: None

    acciones = ttk.Frame(top)
    acciones.grid(row=0, column=3, sticky="e", padx=(0, 24))
    acciones.columnconfigure(0, weight=1)
    try:
        # No fijar una altura/anchura rígida: si se desactiva la propagación
        # sin definir width, el frame puede colapsar y dejar visible solo una línea.
        acciones.grid_propagate(True)
    except Exception:
        pass
    app_state.planilla_seccion_acciones_widget = acciones

    def _guardar_default():
        guardar_hook = getattr(app_state, "planilla_area_guardar_hook", None)
        if callable(guardar_hook):
            try:
                guardar_hook()
            except Exception:
                pass
        try:
            ultimo_path = str(getattr(app_state, "planilla_bundle_last_path", "") or "").strip()
            if ultimo_path:
                guardar_bundle_como(ultimo_path, estado_var)
            else:
                guardar_bundle_default(estado_var)
        except Exception as e:
            if estado_var is not None:
                estado_var.set(f"Error al guardar planilla: {e}")

    def _guardar_como():
        guardar_hook = getattr(app_state, "planilla_area_guardar_hook", None)
        if callable(guardar_hook):
            try:
                guardar_hook()
            except Exception:
                pass
        path = filedialog.asksaveasfilename(
            title="Guardar planilla facturación como...",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
        )
        if path:
            try:
                guardar_bundle_como(path, estado_var)
            except Exception as e:
                if estado_var is not None:
                    estado_var.set(f"Error en Guardar como: {e}")

    def _cargar():
        guardar_hook = getattr(app_state, "planilla_area_guardar_hook", None)
        if callable(guardar_hook):
            try:
                guardar_hook()
            except Exception:
                pass
        path = filedialog.askopenfilename(
            title="Cargar planilla facturación",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
        )
        if path:
            try:
                cargar_bundle(path, estado_var)
            except Exception as e:
                if estado_var is not None:
                    estado_var.set(f"Error al cargar planilla: {e}")

    def _limpiar():
        guardar_hook = getattr(app_state, "planilla_area_guardar_hook", None)
        if callable(guardar_hook):
            try:
                guardar_hook()
            except Exception:
                pass
        try:
            limpiar_planilla_facturacion(estado_var)
        except Exception as e:
            if estado_var is not None:
                estado_var.set(f"Error al limpiar planilla: {e}")

    def _pasar_ultima_semana_a_nueva():
        hook = getattr(app_state, "planilla_pasarela_ultima_semana_hook", None)

        # Si todavía no existe el hook, intentamos construir la sección Área Recaudación
        # para registrar la pasarela automáticamente.
        if not callable(hook) and "Area Recaudación" in frames:
            _mostrar("Area Recaudación")
            combo.set("Area Recaudación")
            hook = getattr(app_state, "planilla_pasarela_ultima_semana_hook", None)

        if callable(hook):
            try:
                hook()
            except Exception as e:
                if estado_var is not None:
                    estado_var.set(f"Error al pasar datos de última semana: {e}")
        elif estado_var is not None:
            estado_var.set("No se pudo inicializar la acción de última semana.")

    btn_voy_a = ttk.Button(
        acciones,
        text="Voy a...",
        style="Marino.TButton",
        width=14,
        command=lambda: _abrir_menu_voy_a(),
    )

    menu_voy_a = tk.Menu(btn_voy_a, tearoff=0)
    try:
        menu_voy_a.configure(
            font=("Segoe UI", 10),
            background="#F8FAFC",
            foreground="#0F172A",
            activebackground="#1D4ED8",
            activeforeground="#FFFFFF",
            relief="solid",
            borderwidth=1,
        )
    except tk.TclError:
        # Algunos temas/plataformas no soportan todas las opciones visuales del Menu.
        pass

    menu_voy_a.add_command(label="Guardar planilla", command=_guardar_default)
    menu_voy_a.add_command(label="Guardar como...", command=_guardar_como)
    menu_voy_a.add_command(label="Cargar...", command=_cargar)
    menu_voy_a.add_command(label="Limpiar planilla", command=_limpiar)
    menu_voy_a.add_separator()
    menu_voy_a.add_command(label="Pasar datos de última semana a planilla nueva", command=_pasar_ultima_semana_a_nueva)

    def _abrir_menu_voy_a():
        try:
            btn_voy_a.update_idletasks()
            menu_voy_a.update_idletasks()

            # Abrir exactamente debajo del botón y alineado al borde izquierdo.
            x = btn_voy_a.winfo_rootx()
            y = btn_voy_a.winfo_rooty() + btn_voy_a.winfo_height() + 2

            # Mantener el menú dentro de pantalla sin forzarlo a otra posición rara.
            screen_w = btn_voy_a.winfo_screenwidth()
            screen_h = btn_voy_a.winfo_screenheight()
            req_w = max(menu_voy_a.winfo_reqwidth(), 220)
            req_h = max(menu_voy_a.winfo_reqheight(), 10)

            if x + req_w > screen_w - 8:
                x = max(8, screen_w - req_w - 8)
            if y + req_h > screen_h - 8:
                # Si no entra hacia abajo, lo pegamos al borde inferior,
                # pero priorizando siempre que salga desde abajo del botón.
                y = max(btn_voy_a.winfo_rooty() + btn_voy_a.winfo_height() + 2, screen_h - req_h - 8)

            menu_voy_a.tk_popup(x, y)
        finally:
            try:
                menu_voy_a.grab_release()
            except Exception:
                pass

    btn_voy_a.grid(row=0, column=0, sticky="e")

    stack = ttk.Frame(parent)
    stack.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
    stack.columnconfigure(0, weight=1)
    stack.rowconfigure(0, weight=1)

    frames = {}
    built: set[str] = set()
    _visible = {"target": ""}

    def _programar_refresh_visual(opt: str):
        hook = getattr(app_state, "planilla_visual_refresh_hooks", {}).get(opt)
        if not callable(hook):
            return
        try:
            parent.after_idle(hook)
        except Exception:
            try:
                hook()
            except Exception:
                pass

    for opt in opciones:
        fr = ttk.Frame(stack)
        fr.grid(row=0, column=0, sticky="nsew")
        fr.columnconfigure(0, weight=1)
        fr.rowconfigure(1, weight=1)
        frames[opt] = fr

    # Lazy build real: construir cada sección solo cuando se abre por primera vez.
    # Esto baja mucho el tiempo de arranque y evita que la primera apertura de
    # Planilla Facturación se rompa por construir todas las vistas pesadas juntas.

    def _mostrar(opt: str):
        if opt not in frames:
            return

        previo = str(_visible.get("target", "") or "").strip()
        if previo and previo != opt:
            commit_hook = getattr(app_state, "planilla_section_commit_hooks", {}).get(previo)
            if callable(commit_hook):
                try:
                    commit_hook()
                except Exception:
                    pass

            if previo == "Area Recaudación":
                guardar_hook = getattr(app_state, "planilla_area_guardar_hook", None)
                if callable(guardar_hook):
                    try:
                        guardar_hook()
                    except Exception:
                        pass

        if opt not in built:
            try:
                on_build_frame(opt, frames[opt])
                built.add(opt)
            except Exception:
                return
        if _visible["target"] != opt:
            frames[opt].tkraise()
            _visible["target"] = opt
        _programar_refresh_visual(opt)

    def _on_change(_evt=None):
        v = combo.get()
        if v in frames:
            _mostrar(v)

    combo.bind("<<ComboboxSelected>>", _on_change)

    if opciones:
        combo.set(opciones[0])
        try:
            parent.after_idle(lambda: _mostrar(opciones[0]))
        except Exception:
            _mostrar(opciones[0])

    return combo, frames, stack


def crear_tab_planilla_facturacion(frame, root, estado_var):
    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    cont = ttk.Frame(frame)
    cont.grid(row=0, column=0, sticky="nsew")
    cont.columnconfigure(0, weight=1)
    cont.rowconfigure(1, weight=1)

    secciones = ["Area Recaudación", "Prescripciones", "Anticipos y Topes", "Control CIO", "AGENCIA AMIGA", "TOTALES"]

    def build_seccion(nombre: str, fr: ttk.Frame):
        if nombre == "Area Recaudación":
            build_area_recaudacion(fr, estado_var)
        elif nombre == "Prescripciones":
            build_prescripciones(fr, estado_var)
        elif nombre == "Anticipos y Topes":
            build_anticipos_topes(fr, estado_var)
        elif nombre == "Control CIO":
            build_control_cio(fr, estado_var)
        elif nombre == "AGENCIA AMIGA":
            build_agencia_amiga(fr, estado_var)
        else:
            build_totales(fr, estado_var)

    _crear_vista_filtrable(cont, "Sección:", secciones, build_seccion, estado_var)

    if not _STORAGE_DISPONIBLE and estado_var is not None:
        extra = f" ({_storage_error})" if _storage_error else ""
        estado_var.set("Storage externo no disponible: se usa fallback interno." + extra)

    # Levantar el autosave una vez que la solapa ya terminó de montarse.
    # Evita mezclar construcción pesada + hidratación de datos en el mismo ciclo.
    try:
        cont.after_idle(lambda: cargar_bundle_default_si_existe(estado_var))
    except Exception:
        cargar_bundle_default_si_existe(estado_var)
