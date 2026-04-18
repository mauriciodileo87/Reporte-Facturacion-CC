# app_state.py
# Estado compartido + "bus" para pasar datos entre solapas (Tickets/Reporte/SFA -> Planilla)

import re
from datetime import datetime

# Hooks: cada tab de Planilla Facturación registra un callback para refrescarse
planilla_refresh_hooks = {}
planilla_area_snapshot_hooks = {}

# Tickets -> Planilla (solo recaud y premios menores; comisión se calcula en planilla)
tickets_resumen_por_juego = {}

# Tickets -> Planilla por semana (fuente persistible / no mezclar semanas)
tickets_resumen_por_semana_por_juego = {f"Semana {i}": {} for i in range(1, 6)}
tickets_prescripciones_por_semana_por_juego = {f"Semana {i}": {} for i in range(1, 6)}

# Reporte -> Planilla (recaud, comi, prem)
reporte_resumen_por_juego = {}

# JSON/SFA -> Planilla (recaud, comi, prem)
sfa_resumen_por_juego = {}

# JSON/SFA -> Comisiones Agencia Amiga (concepto Z118) por juego/sorteo
sfa_z118_por_juego = {}

# Flag interno para importaciones masivas desde JSON/SFA
sfa_bulk_publish = False

# Estado transitorio de una importación SFA en curso.
# Se usa para que una nueva importación reemplace SOLO los sorteos/conceptos
# presentes en el archivo nuevo, sin borrar semanas anteriores.
_sfa_seen_resumen = set()   # {(juego_tab, sorteo, concepto)}
_sfa_seen_presc = set()     # {(juego_tab, sorteo)}
_sfa_seen_z118 = set()      # {(codigo_juego_norm, sorteo)}

# Hooks diferidos para importación masiva SFA (evita refrescar UI por cada fila).
_sfa_dirty_planilla_hooks = set()
_sfa_dirty_presc_hooks = set()
_sfa_dirty_totales = False

# Prescripciones para Planilla Facturación > Prescripciones
reporte_prescripciones_por_juego = {}
# Estructura semanal: {"Semana N": {"Juego Planilla": {"sorteo_norm": importe_float}}}
reporte_prescripciones_por_juego_por_semana = {
    f"Semana {i}": {} for i in range(1, 6)
}
tickets_prescripciones_por_juego = {}
sfa_prescripciones_por_juego = {}

# Sorteos base por semana para la sección Prescripciones.
prescripciones_sorteos_por_semana_por_juego = {}

# Snapshot persistible de la grilla de Prescripciones por juego/semana.
# Estructura: {juego: {semana: {sorteo: {"t_presc": float|None, "r_presc": float|None, "s_presc": float|None}}}}
planilla_prescripciones_data = {}

# Hooks de refresco para la sub-solapa Prescripciones en Planilla Facturación
planilla_presc_refresh_hooks = {}

# Hooks para resetear la sección Anticipos y Topes
planilla_anticipos_reset_hooks = {}

# Hooks para resetear la sección Control CIO
planilla_control_cio_reset_hooks = {}

# Hooks para refrescar/cargar la sección Control CIO
planilla_control_cio_load_hooks = {}

# Hooks de serialización por sección para Guardar/Cargar planilla completa.
# Cada snapshot hook devuelve dict serializable y cada load hook recibe ese dict.
planilla_bundle_snapshot_hooks = {}
planilla_bundle_load_hooks = {}

# Estado persistible de Anticipos y Topes
planilla_anticipos_topes_data = {}

# Estado persistible de Control CIO
planilla_control_cio_data = {}

# Estado persistible de Agencia Amiga
planilla_agencia_amiga_data = {}

# Reporte Facturación -> Agencia Amiga (Tobill/SFA)
# Estructura: {"Juego Planilla": {"sorteo_norm": importe_float}}
reporte_agencia_amiga_tobill_por_juego = {}
reporte_agencia_amiga_sfa_118_por_juego = {}
# Estructura semanal: {"Semana N": {"Juego Planilla": {"sorteo_norm": importe_float}}}
reporte_agencia_amiga_tobill_por_juego_por_semana = {
    f"Semana {i}": {} for i in range(1, 6)
}
reporte_agencia_amiga_sfa_118_por_juego_por_semana = {
    f"Semana {i}": {} for i in range(1, 6)
}

# Hooks para refrescar la sección Agencia Amiga cuando cambia Reporte
planilla_agencia_amiga_refresh_hooks = {}

# Hooks de carga para Agencia Amiga
planilla_agencia_amiga_load_hooks = {}

# Última ruta elegida por el usuario para guardar/cargar el bundle completo de Planilla.
planilla_bundle_last_path = ""

# Importe del concepto FACUNI detectado en la sub-solapa Tobill de Reporte
reporte_tobill_facuni_importe = 0.0
reporte_tobill_facuni_cargado = False

# Suma de SALE LIMIT por juego detectada en Reporte > Tobill
reporte_tobill_sale_limit_por_juego = {}
# Suma de SALE LIMIT por juego detectada en Reporte > Tobill, separado por semana.
# Estructura: {"Semana N": {"Juego Planilla": importe_float}}
reporte_tobill_sale_limit_por_juego_por_semana = {
    f"Semana {i}": {} for i in range(1, 6)
}

# Suma del concepto ADVANCE detectada en Reporte > Tobill
reporte_tobill_advance_importe = 0.0
# Suma del concepto ADVANCE detectada en Reporte > Tobill, separado por semana.
# Estructura: {"Semana N": importe_float}
reporte_tobill_advance_importe_por_semana = {
    f"Semana {i}": 0.0 for i in range(1, 6)
}

# Importes CASH IN / CASH OUT detectados en Reporte Facturación
reporte_cash_in_importe = 0.0
reporte_cash_out_importe = 0.0

# Última semana inferida desde el archivo de Reporte/Tobill
reporte_tobill_ultima_semana_importada = "Semana 1"

# Hooks para refrescar Anticipos y Topes cuando cambia Reporte > Tobill
planilla_anticipos_refresh_hooks = {}

# Hooks para refrescar Control CIO cuando cambia Reporte Facturación
planilla_control_cio_refresh_hooks = {}

# Hooks para refrescar sección Totales cuando hay cambios de datos
planilla_totales_refresh_hooks = {}

# Hooks para propagar el filtro (Semana + Del/Al) definido en Área Recaudación
planilla_semana_filtro_hooks = {}
planilla_global_reset_hooks = {}
planilla_semana_filtro_actual = {
    "juego": "",
    "semana": 0,
    "desde": "",
    "hasta": "",
}

# Rango Del/Al consolidado por semana (independiente del juego activo)
planilla_rangos_semana_global = {}

# Hook opcional para refrescar la fila "Total reporte tobill" en la solapa FACUNI
facuni_actualizar_total_tobill_hook = None

# Total de afectaciones importado desde FACUNI
facuni_total = 0.0

# Totales FACUNI por semana
facuni_total_por_semana = {
    "Semana 1": 0.0,
    "Semana 2": 0.0,
    "Semana 3": 0.0,
    "Semana 4": 0.0,
    "Semana 5": 0.0,
}

# Total "TOTAL REPORTE TOBILL" de FACUNI por semana
facuni_reporte_tobill_por_semana = {
    "Semana 1": 0.0,
    "Semana 2": 0.0,
    "Semana 3": 0.0,
    "Semana 4": 0.0,
    "Semana 5": 0.0,
}

# Última semana inferida desde archivo FACUNI
facuni_ultima_semana_importada = "Semana 1"

# Último total FACUNI importado pendiente de pasar manualmente a Planilla > Totales
facuni_total_pendiente = 0.0


def _ruta_guardado_planilla_area() -> str:
    import os

    appdata = os.environ.get("APPDATA")
    if appdata:
        base = os.path.join(appdata, "ReporteFacturacion")
    else:
        base = os.path.join(os.path.expanduser("~"), ".reporte_facturacion")

    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "planilla_facturacion_guardada.json")


def _leer_planillas_guardadas_area() -> dict:
    import json
    import os

    path = _ruta_guardado_planilla_area()
    if not os.path.exists(path):
        return {}

    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    if not isinstance(data, dict):
        return {}

    planillas = data.get("planillas", {})
    return planillas if isinstance(planillas, dict) else {}


def _normalizar_mapa_semanas_guardado(raw: dict | None) -> dict[str, list[int]]:
    semanas: dict[str, set[int]] = {}
    if not isinstance(raw, dict):
        return {}

    for k, vals in raw.items():
        try:
            sem = int(str(k).strip().lower().replace("semana", "").strip())
        except Exception:
            continue
        if sem < 1:
            continue
        if not isinstance(vals, list):
            continue

        bucket = semanas.setdefault(str(sem), set())
        for v in vals:
            try:
                bucket.add(int(v))
            except Exception:
                continue

    return {sem: sorted(valores) for sem, valores in sorted(semanas.items(), key=lambda it: int(it[0]))}


def _normalizar_rangos_guardados(raw: dict | None) -> dict[str, dict[str, str]]:
    rangos: dict[str, dict[str, str]] = {}
    if not isinstance(raw, dict):
        return rangos

    for k, v in raw.items():
        try:
            sem = int(str(k).strip().lower().replace("semana", "").strip())
        except Exception:
            continue
        if sem < 1:
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
            rangos[str(sem)] = {"desde": desde, "hasta": hasta}

    return dict(sorted(rangos.items(), key=lambda it: int(it[0])))


def obtener_snapshots_area_recaudacion() -> dict[str, dict]:
    out: dict[str, dict] = {}

    hooks = globals().get("planilla_area_snapshot_hooks", {}) or {}
    for juego, hook in list(hooks.items()):
        if not callable(hook):
            continue
        try:
            payload = hook()
        except Exception:
            continue
        if isinstance(payload, dict):
            out[str(juego)] = payload

    for juego, data_juego in _leer_planillas_guardadas_area().items():
        juego_txt = str(juego).strip()
        if not juego_txt or juego_txt in out or not isinstance(data_juego, dict):
            continue

        out[juego_txt] = {
            "codigo_juego": data_juego.get("codigo_juego"),
            "columnas": data_juego.get("columnas", []),
            "filas": data_juego.get("filas", []),
            "semanas": _normalizar_mapa_semanas_guardado(data_juego.get("semanas", {})),
            "rangos_semana": _normalizar_rangos_guardados(data_juego.get("rangos_semana", {})),
        }

    return out


def sincronizar_rangos_semana_global_desde_storage():
    for _juego, snapshot in obtener_snapshots_area_recaudacion().items():
        rangos = snapshot.get("rangos_semana", {}) if isinstance(snapshot, dict) else {}
        if not isinstance(rangos, dict):
            continue
        for sem, rango in rangos.items():
            try:
                sem_n = int(sem)
            except Exception:
                continue
            if not isinstance(rango, dict):
                continue
            desde = str(rango.get("desde", "") or "").strip()
            hasta = str(rango.get("hasta", "") or "").strip()
            if desde and hasta and sem_n not in planilla_rangos_semana_global:
                planilla_rangos_semana_global[sem_n] = (desde, hasta)


def _obtener_sorteos_planilla_por_juego() -> dict[str, set[str]]:
    """
    Lee los sorteos conocidos por juego desde los snapshot hooks de Planilla.
    Sirve para filtrar importaciones externas y pasar a la planilla únicamente
    filas que realmente coinciden con los números de sorteo cargados allí.
    """
    out: dict[str, set[str]] = {}
    for juego, payload in obtener_snapshots_area_recaudacion().items():
        if not isinstance(payload, dict):
            continue

        sorteos_juego = out.setdefault(str(juego).strip(), set())
        for fila in payload.get("filas", []) or []:
            if not isinstance(fila, list) or not fila:
                continue
            s = _normalizar_sorteo_clave(fila[0])
            if s:
                sorteos_juego.add(s)

        semanas = payload.get("semanas", {}) or {}
        if isinstance(semanas, dict):
            for valores in semanas.values():
                if not isinstance(valores, list):
                    continue
                for sorteo in valores:
                    s = _normalizar_sorteo_clave(sorteo)
                    if s:
                        sorteos_juego.add(s)
    return out


def sorteo_existe_en_planilla(nombre_juego_planilla: str, sorteo: str) -> bool:
    juego = str(nombre_juego_planilla or "").strip()
    if not juego:
        return False
    sorteos_por_juego = _obtener_sorteos_planilla_por_juego()
    if juego not in sorteos_por_juego or not sorteos_por_juego[juego]:
        # Si la planilla todavía no publicó sus sorteos, no bloquear la carga.
        return True
    return _normalizar_sorteo_clave(sorteo) in sorteos_por_juego[juego]


def _obtener_sorteos_prescripciones_por_juego(semana: int | None = None) -> dict[str, set[str]]:
    """
    Lee los sorteos base de Prescripciones por juego/semana.

    Regla de negocio:
    la sección Prescripciones se define EXCLUSIVAMENTE con el botón
    "Importar consulta de prescripciones". No debe heredar sorteos desde
    Área Recaudación ni desde PJU.
    """
    out: dict[str, set[str]] = {}
    src = prescripciones_sorteos_por_semana_por_juego or {}
    if not isinstance(src, dict):
        return out

    for juego, semanas in src.items():
        juego_key = str(juego or "").strip()
        if not juego_key or not isinstance(semanas, dict):
            continue

        bucket = out.setdefault(juego_key, set())
        if semana is not None:
            sorteos_semana = semanas.get(semana, semanas.get(str(semana), []))
            sorteos_fuente = [sorteos_semana]
        else:
            sorteos_fuente = semanas.values()

        for sorteos in sorteos_fuente:
            if not isinstance(sorteos, list):
                continue
            for sorteo in sorteos:
                s = _normalizar_sorteo_clave(sorteo)
                if s:
                    bucket.add(s)

    return out


def sorteo_existe_en_prescripciones(nombre_juego_planilla: str, sorteo: str, semana: int | None = None) -> bool:
    """
    Solo permite enlazar importes de Prescripciones si el sorteo ya existe en la
    base cargada por "Importar consulta de prescripciones" para ese juego/semana.
    """
    juego = str(nombre_juego_planilla or "").strip()
    if not juego:
        return False
    sorteo_norm = _normalizar_sorteo_clave(sorteo)
    if not sorteo_norm:
        return False

    sorteos_por_juego = _obtener_sorteos_prescripciones_por_juego(semana=semana)
    bucket = sorteos_por_juego.get(juego, set())
    if not bucket:
        return False
    return sorteo_norm in bucket


def _mergear_resumen_importado(destino: dict, juego: str, sorteo: str, data: dict, *, con_comision: bool):
    j = str(juego).strip()
    s = _normalizar_sorteo_clave(sorteo)
    if not j or not s:
        return False
    if not sorteo_existe_en_planilla(j, s):
        return False

    bucket = destino.setdefault(j, {})
    actual = bucket.setdefault(s, {})
    recaud_nuevo = float((data or {}).get("recaud", actual.get("recaud", 0.0)) or 0.0)
    prem_nuevo = float((data or {}).get("prem", actual.get("prem", 0.0)) or 0.0)
    comi_nuevo = float((data or {}).get("comi", actual.get("comi", 0.0)) or 0.0) if con_comision else None

    # Si un import externo no trae datos para ese sorteo, suele llegar como todo en 0.
    # En ese caso no pisar lo ya cargado para evitar "borrados" involuntarios.
    if con_comision:
        sin_datos = recaud_nuevo == 0.0 and prem_nuevo == 0.0 and comi_nuevo == 0.0
    else:
        sin_datos = recaud_nuevo == 0.0 and prem_nuevo == 0.0
    if sin_datos and s in bucket:
        return False

    actual["recaud"] = recaud_nuevo
    if con_comision:
        actual["comi"] = comi_nuevo
    actual["prem"] = prem_nuevo
    return True


def _mergear_prescripcion_importada(destino: dict, juego: str, sorteo: str, importe):
    j = str(juego).strip()
    s = _normalizar_sorteo_clave(sorteo)
    if not j or not s:
        return False
    destino.setdefault(j, {})
    importe_nuevo = float(importe or 0.0)

    # Misma regla de no pisar con "vacío" (0) cuando ya había valor cargado.
    if importe_nuevo == 0.0 and s in destino[j]:
        return False

    destino[j][s] = importe_nuevo
    return True


def _refresh_totales():
    hooks = list(planilla_totales_refresh_hooks.items())
    for key, hook in hooks:
        if not callable(hook):
            continue
        try:
            hook()
        except Exception:
            planilla_totales_refresh_hooks.pop(key, None)


def texto_rango_semana_global(semana: int) -> str:
    try:
        sem_n = int(semana)
    except Exception:
        sem_n = 1
    if sem_n < 1:
        sem_n = 1

    desde = ""
    hasta = ""

    try:
        guardado = planilla_rangos_semana_global.get(sem_n)
    except Exception:
        guardado = None

    if isinstance(guardado, (list, tuple)) and len(guardado) >= 2:
        desde = str(guardado[0] or "").strip()
        hasta = str(guardado[1] or "").strip()
    elif isinstance(guardado, dict):
        desde = str(guardado.get("desde", "") or "").strip()
        hasta = str(guardado.get("hasta", "") or "").strip()

    if (not desde or not hasta) and int(planilla_semana_filtro_actual.get("semana", 1) or 1) == sem_n:
        desde = desde or str(planilla_semana_filtro_actual.get("desde", "") or "").strip()
        hasta = hasta or str(planilla_semana_filtro_actual.get("hasta", "") or "").strip()

    if not desde or not hasta:
        return "Del: --/--/---- al: --/--/----"
    return f"Del: {desde} al: {hasta}"


def semana_visible_desde_valor(valor) -> str:
    semana_txt = _normalizar_semana_txt(valor)
    try:
        semana_n = int(re.search(r"(\d+)", semana_txt).group(1))
    except Exception:
        semana_n = 1
    return texto_rango_semana_global(semana_n)


def semana_interna_desde_visible(valor) -> str:
    return _normalizar_semana_txt(valor)


def _normalizar_sorteo_clave(valor) -> str:
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


def _normalizar_concepto_sfa(concepto) -> str:
    txt = str(concepto).strip().upper()
    if not txt:
        return ""
    if txt.startswith("Z"):
        dig = "".join(ch for ch in txt[1:] if ch.isdigit())
        return f"Z{dig.zfill(3)}" if dig else txt
    if txt.isdigit():
        return f"Z{txt.zfill(3)}"
    return txt


def _normalizar_codigo_juego_sfa(codigo_juego: str) -> str:
    raw = str(codigo_juego).strip()
    if not raw:
        return ""

    codigo = _normalizar_sorteo_clave(raw)
    if codigo in {"66", "67", "68", "69"}:
        return "69"
    if codigo in {"7", "8", "9", "10", "11"}:
        return "9"

    m = re.search(r"\d+", raw)
    if m:
        codigo_extraido = _normalizar_sorteo_clave(m.group(0))
        if codigo_extraido in {"66", "67", "68", "69"}:
            return "69"
        if codigo_extraido in {"7", "8", "9", "10", "11"}:
            return "9"

    return raw


def _map_codigo_juego_a_tab_planilla(codigo_juego: str) -> str:
    raw = str(codigo_juego).strip()
    if not raw:
        return ""

    codigo = _normalizar_sorteo_clave(raw)
    mapa = {
        "80": "Quiniela",
        "79": "Quiniela Ya",
        "82": "Poceada",
        "74": "Tombolina",
        "66": "Quini 6",
        "67": "Quini 6",
        "68": "Quini 6",
        "69": "Quini 6",
        "13": "Brinco",
        "9": "Loto",
        "5": "Loto 5",
        "41": "LT",
    }

    if codigo in mapa:
        return mapa[codigo]

    m = re.search(r"\d+", raw)
    if m:
        codigo_extraido = _normalizar_sorteo_clave(m.group(0))
        if codigo_extraido in mapa:
            return mapa[codigo_extraido]

    txt = raw.lower().replace("_", " ").replace("-", " ")
    txt = " ".join(txt.split())
    alias = {
        "quiniela": "Quiniela",
        "quiniela ya": "Quiniela Ya",
        "poceada": "Poceada",
        "tombolina": "Tombolina",
        "quini 6": "Quini 6",
        "quini6": "Quini 6",
        "brinco": "Brinco",
        "loto": "Loto",
        "loto 5": "Loto 5",
        "loto5": "Loto 5",
        "lt": "LT",
    }
    return alias.get(txt, "")


def obtener_semana_activa_txt() -> str:
    try:
        sem_n = int((planilla_semana_filtro_actual or {}).get("semana", 0) or 0)
    except Exception:
        sem_n = 0
    if 1 <= sem_n <= 5:
        return f"Semana {sem_n}"
    return "Semana 1"


def _parse_fecha_texto_ddmmyyyy(texto: str):
    txt = str(texto or "").strip()
    if not txt:
        return None
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(txt, fmt).date()
        except Exception:
            pass
    return None


def _parse_fecha_desde_nombre_archivo(nombre_archivo: str):
    nombre = str(nombre_archivo or "")
    patrones = [
        r"(\d{4})[-_]?(\d{2})[-_]?(\d{2})",
        r"(\d{2})[-_]?(\d{2})[-_]?(\d{4})",
    ]
    for patron in patrones:
        m = re.search(patron, nombre)
        if not m:
            continue
        try:
            if len(m.group(1)) == 4:
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            else:
                d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            return datetime(y, mo, d).date()
        except Exception:
            continue
    return None


def obtener_semana_por_fecha_facuni(nombre_archivo: str) -> str:
    sincronizar_rangos_semana_global_desde_storage()
    nombre = str(nombre_archivo or "")

    # Prioridad 1: semana explícita en el nombre del archivo
    # (ej: "...Semana 2...", "...semana_3...", "...SEM4...").
    m_semana = re.search(r"(?i)\bsem(?:ana)?\s*[-_]?(\d)\b", nombre)
    if m_semana:
        try:
            semana_n = int(m_semana.group(1))
            if 1 <= semana_n <= 5:
                return f"Semana {semana_n}"
        except Exception:
            pass

    fecha_archivo = _parse_fecha_desde_nombre_archivo(nombre_archivo)
    if fecha_archivo is None:
        return obtener_semana_activa_txt()

    for i in range(1, 6):
        rango = planilla_rangos_semana_global.get(i)
        if not rango:
            continue

        if isinstance(rango, dict):
            desde_txt = str(rango.get("desde", "") or "").strip()
            hasta_txt = str(rango.get("hasta", "") or "").strip()
        elif isinstance(rango, (list, tuple)) and len(rango) >= 2:
            desde_txt = str(rango[0] or "").strip()
            hasta_txt = str(rango[1] or "").strip()
        else:
            continue

        desde = _parse_fecha_texto_ddmmyyyy(desde_txt)
        hasta = _parse_fecha_texto_ddmmyyyy(hasta_txt)
        if desde and hasta and desde <= fecha_archivo <= hasta:
            return f"Semana {i}"

    return obtener_semana_activa_txt()



def obtener_semana_por_fecha_archivo_totales(nombre_archivo: str) -> str:
    """
    Resuelve la semana destino para Totales usando la fecha del nombre de archivo
    contra los rangos Del/Al consolidados en Planilla Facturación.

    Ejemplo:
    - fecha 20260412 -> semana 05/04/2026 al 12/04/2026
    """
    return obtener_semana_por_fecha_facuni(nombre_archivo)


def publicar_facuni_total(nombre_archivo: str, total: float):
    global facuni_total
    global facuni_total_pendiente
    semana = obtener_semana_por_fecha_archivo_totales(nombre_archivo)
    global facuni_ultima_semana_importada
    facuni_ultima_semana_importada = semana
    facuni_total = float(total or 0.0)
    facuni_total_pendiente = float(total or 0.0)


def pasar_facuni_a_planilla(semana: str | None = None, total: float | None = None):
    semana_objetivo = str(semana or facuni_ultima_semana_importada or "").strip() or obtener_semana_activa_txt()
    total_objetivo = facuni_total_pendiente if total is None else total

    facuni_total_por_semana[semana_objetivo] = float(total_objetivo or 0.0)
    facuni_reporte_tobill_por_semana[semana_objetivo] = float(reporte_tobill_facuni_importe or 0.0)
    _refresh_totales()


def publicar_reporte_tobill_facuni(nombre_archivo: str, total_facuni: float):
    """
    Publica el total FACUNI detectado en Reporte>Tobill en:
    - variable global (compatibilidad),
    - y mapa por semana inferida desde el nombre del archivo.
    """
    global reporte_tobill_facuni_importe
    global reporte_tobill_facuni_cargado

    reporte_tobill_facuni_importe = round(float(total_facuni or 0.0), 2)
    reporte_tobill_facuni_cargado = True

    semana = obtener_semana_por_fecha_archivo_totales(nombre_archivo)
    facuni_reporte_tobill_por_semana[semana] = float(reporte_tobill_facuni_importe or 0.0)
    _refresh_totales()


def publicar_reporte_tobill_anticipos_topes(
    nombre_archivo: str,
    sale_limit_por_juego: dict | None,
    advance_importe: float,
):
    """
    Publica SALE LIMIT + ADVANCE del reporte Tobill en:
    - variables globales actuales (compatibilidad),
    - y mapa por semana inferido desde la fecha del nombre de archivo.
    """
    global reporte_tobill_sale_limit_por_juego
    global reporte_tobill_advance_importe
    global reporte_tobill_ultima_semana_importada

    sale_limit_normalizado = {}
    for juego, importe in (sale_limit_por_juego or {}).items():
        juego_txt = str(juego or "").strip()
        if not juego_txt:
            continue
        try:
            sale_limit_normalizado[juego_txt] = round(float(importe or 0.0), 2)
        except Exception:
            sale_limit_normalizado[juego_txt] = 0.0

    reporte_tobill_sale_limit_por_juego = sale_limit_normalizado
    reporte_tobill_advance_importe = round(float(advance_importe or 0.0), 2)

    semana = obtener_semana_por_fecha_archivo_totales(nombre_archivo)
    reporte_tobill_ultima_semana_importada = str(semana or "").strip() or "Semana 1"
    if semana not in reporte_tobill_sale_limit_por_juego_por_semana:
        reporte_tobill_sale_limit_por_juego_por_semana[semana] = {}
    reporte_tobill_sale_limit_por_juego_por_semana[semana] = dict(sale_limit_normalizado)
    reporte_tobill_advance_importe_por_semana[semana] = float(reporte_tobill_advance_importe or 0.0)
    _refresh_totales()


def publicar_tickets_resumen(nombre_juego_planilla: str, sorteo: str, recaud: float, premios: float):
    _mergear_resumen_importado(
        tickets_resumen_por_juego,
        nombre_juego_planilla,
        sorteo,
        {"recaud": recaud, "prem": premios},
        con_comision=False,
    )

    j = str(nombre_juego_planilla).strip()
    hook = planilla_refresh_hooks.get(j)
    if callable(hook):
        hook()


def publicar_reporte_resumen(nombre_juego_planilla: str, sorteo: str, recaud: float, comi: float, premios: float):
    _mergear_resumen_importado(
        reporte_resumen_por_juego,
        nombre_juego_planilla,
        sorteo,
        {"recaud": recaud, "comi": comi, "prem": premios},
        con_comision=True,
    )

    j = str(nombre_juego_planilla).strip()
    hook = planilla_refresh_hooks.get(j)
    if callable(hook):
        hook()
    _refresh_totales()


def publicar_reporte_prescripcion(nombre_juego_planilla: str, sorteo: str, importe: float):
    _mergear_prescripcion_importada(reporte_prescripciones_por_juego, nombre_juego_planilla, sorteo, importe)

    j = str(nombre_juego_planilla).strip()
    hook = planilla_presc_refresh_hooks.get(j)
    if callable(hook):
        hook()
    _refresh_totales()


def publicar_ticket_prescripto(nombre_juego_planilla: str, sorteo: str, importe: float):
    _mergear_prescripcion_importada(tickets_prescripciones_por_juego, nombre_juego_planilla, sorteo, importe)

    j = str(nombre_juego_planilla).strip()
    hook = planilla_presc_refresh_hooks.get(j)
    if callable(hook):
        hook()
    _refresh_totales()


def _normalizar_semana_txt(semana) -> str:
    txt = str(semana or "").strip()
    if txt:
        m_sem = re.fullmatch(r"(?i)\s*semana\s*(\d+)\s*", txt)
        if m_sem:
            try:
                n = max(1, min(5, int(m_sem.group(1))))
                return f"Semana {n}"
            except Exception:
                pass

        for _n in range(1, 6):
            try:
                visible = texto_rango_semana_global(_n)
            except Exception:
                visible = ""
            if visible and visible != "Del: --/--/---- al: --/--/----" and txt == visible:
                return f"Semana {_n}"

        if re.fullmatch(r"\d+", txt):
            try:
                n = max(1, min(5, int(txt)))
                return f"Semana {n}"
            except Exception:
                pass

    try:
        n = max(1, min(5, int(semana)))
        return f"Semana {n}"
    except Exception:
        return obtener_semana_activa_txt()


def _asegurar_fuentes_txt_data():
    data = _asegurar_planilla_totales_data()
    fuentes = data.setdefault("fuentes_txt", {})
    if not isinstance(fuentes, dict):
        fuentes = {}
        data["fuentes_txt"] = fuentes
    fuentes_presc = data.setdefault("fuentes_txt_prescripciones", {})
    if not isinstance(fuentes_presc, dict):
        fuentes_presc = {}
        data["fuentes_txt_prescripciones"] = fuentes_presc
    for i in range(1, 6):
        fuentes.setdefault(f"Semana {i}", {})
        fuentes_presc.setdefault(f"Semana {i}", {})
    return data, fuentes, fuentes_presc


def _reconstruir_tickets_globales_desde_fuentes_txt():
    global tickets_resumen_por_juego, tickets_prescripciones_por_juego
    tickets_resumen_por_juego.clear()
    tickets_prescripciones_por_juego.clear()

    _data, fuentes, fuentes_presc = _asegurar_fuentes_txt_data()

    for semana_key, juegos in (fuentes.items() if isinstance(fuentes, dict) else []):
        if not isinstance(juegos, dict):
            continue
        sem_bucket = tickets_resumen_por_semana_por_juego.setdefault(semana_key, {})
        sem_bucket.clear()
        for juego, sorteos in juegos.items():
            if not isinstance(sorteos, dict):
                continue
            sem_bucket[juego] = {}
            for sorteo, payload in sorteos.items():
                s = _normalizar_sorteo_clave(sorteo)
                if not s:
                    continue
                row = payload if isinstance(payload, dict) else {}
                rec = float(row.get("recaud", 0.0) or 0.0)
                prem = float(row.get("prem", 0.0) or 0.0)
                sem_bucket[juego][s] = {"recaud": rec, "prem": prem}
                tickets_resumen_por_juego.setdefault(juego, {})[s] = {"recaud": rec, "prem": prem}

    for semana_key, juegos in (fuentes_presc.items() if isinstance(fuentes_presc, dict) else []):
        if not isinstance(juegos, dict):
            continue
        sem_bucket = tickets_prescripciones_por_semana_por_juego.setdefault(semana_key, {})
        sem_bucket.clear()
        for juego, sorteos in juegos.items():
            if not isinstance(sorteos, dict):
                continue
            sem_bucket[juego] = {}
            for sorteo, importe in sorteos.items():
                s = _normalizar_sorteo_clave(sorteo)
                if not s:
                    continue
                try:
                    imp = float(importe or 0.0)
                except Exception:
                    imp = 0.0
                sem_bucket[juego][s] = imp
                tickets_prescripciones_por_juego.setdefault(juego, {})[s] = imp


def inferir_semana_por_sorteo_en_planilla(nombre_juego_planilla: str, sorteo: str, semana_preferida=None) -> str:
    juego = str(nombre_juego_planilla or "").strip()
    s = _normalizar_sorteo_clave(sorteo)
    if not juego or not s:
        return _normalizar_semana_txt(semana_preferida)

    candidata = _normalizar_semana_txt(semana_preferida)
    snapshots = obtener_snapshots_area_recaudacion() or {}
    payload = snapshots.get(juego, {}) if isinstance(snapshots, dict) else {}
    semanas = payload.get("semanas", {}) if isinstance(payload, dict) else {}
    if isinstance(semanas, dict):
        for sem, sorteos in semanas.items():
            if not isinstance(sorteos, list):
                continue
            if s in {_normalizar_sorteo_clave(x) for x in sorteos}:
                return _normalizar_semana_txt(sem)

    for j, data_juego in _leer_planillas_guardadas_area().items():
        if str(j).strip() != juego or not isinstance(data_juego, dict):
            continue
        semanas_g = _normalizar_mapa_semanas_guardado(data_juego.get("semanas", {}))
        for sem, sorteos in semanas_g.items():
            if s in {_normalizar_sorteo_clave(x) for x in sorteos}:
                return _normalizar_semana_txt(sem)

    return candidata


def calcular_totales_txt_tickets_por_semana(semana) -> dict:
    semana_key = _normalizar_semana_txt(semana)
    _data, fuentes, fuentes_presc = _asegurar_fuentes_txt_data()
    juegos = fuentes.get(semana_key, {}) if isinstance(fuentes, dict) else {}
    juegos_presc = fuentes_presc.get(semana_key, {}) if isinstance(fuentes_presc, dict) else {}

    tasas = {
        "Quiniela": 0.20,
        "Quiniela Ya": 0.20,
        "Poceada": 0.20,
        "Tombolina": 0.20,
        "LT": 0.20,
        "Loto": 0.14,
        "Loto 5": 0.14,
        "Quini 6": 0.14,
        "Brinco": 0.14,
    }

    ventas = 0.0
    premios = 0.0
    comisiones = 0.0
    prescripcion = 0.0

    for juego, sorteos in (juegos.items() if isinstance(juegos, dict) else []):
        tasa = tasas.get(str(juego), 0.0)
        for _s, payload in (sorteos.items() if isinstance(sorteos, dict) else []):
            if not isinstance(payload, dict):
                continue
            rec = float(payload.get("recaud", 0.0) or 0.0)
            prem = float(payload.get("prem", 0.0) or 0.0)
            ventas += rec
            premios += prem
            comisiones += rec * tasa

    # OJO: prescripción TXT sólo se suma desde los sorteos que realmente
    # quedaron linkeados en Prescripciones para la semana.
    for _juego, sorteos in (juegos_presc.items() if isinstance(juegos_presc, dict) else []):
        for _s, importe in (sorteos.items() if isinstance(sorteos, dict) else []):
            try:
                prescripcion += float(importe or 0.0)
            except Exception:
                pass

    return {
        "Total ventas": round(ventas, 2),
        "Total comisiones": round(comisiones, 2),
        "Total premios": round(premios, 2),
        "Total prescripcion": round(prescripcion, 2),
    }


def recalcular_y_guardar_totales_txt_semana(semana, overrides: dict | None = None) -> dict:
    """
    Recalcula la parte TXT que proviene de Tickets para una semana y la
    persiste en Totales, preservando las demás filas TXT que pertenecen a
    otras secciones (Anticipos/Topes, FACUNI, Agencia Amiga, etc.).
    """
    semana_key = _normalizar_semana_txt(semana)
    valores_txt = calcular_totales_txt_tickets_por_semana(semana_key) or {}

    obtener = globals().get("obtener_totales_importados_semana")
    bucket_actual = obtener(semana_key) if callable(obtener) else {}
    if isinstance(bucket_actual, dict):
        for etiqueta in ("Total anticipos", "Total topes", "Total facuni", "Total comision agencia amiga"):
            fila = bucket_actual.get(etiqueta, {})
            if isinstance(fila, dict):
                try:
                    valores_txt[etiqueta] = float(fila.get("txt", 0.0) or 0.0)
                except Exception:
                    valores_txt[etiqueta] = 0.0

    if isinstance(overrides, dict):
        for etiqueta, valor in overrides.items():
            try:
                valores_txt[str(etiqueta)] = round(float(valor or 0.0), 2)
            except Exception:
                valores_txt[str(etiqueta)] = 0.0

    guardar = globals().get("guardar_totales_importados")
    if callable(guardar):
        guardar(semana_key, "txt", valores_txt)

    return valores_txt


def reemplazar_tickets_importado(resumen_por_juego: dict, prescripciones_por_juego: dict, semana: int | str | None = None):
    semana_key = _normalizar_semana_txt(semana)
    _reconstruir_tickets_globales_desde_fuentes_txt()
    _data, fuentes, fuentes_presc = _asegurar_fuentes_txt_data()

    bucket_semana = fuentes.setdefault(semana_key, {})
    bucket_presc_semana = fuentes_presc.setdefault(semana_key, {})

    sorteos_importados: dict[str, set[str]] = {}

    # Un sorteo pertenece a una sola semana. Si reimportan/corrigen el mismo
    # sorteo, hay que pisarlo en su semana destino y removerlo del resto.
    for juego, sorteos in (resumen_por_juego or {}).items():
        j = str(juego).strip()
        if not j or not isinstance(sorteos, dict):
            continue
        bucket_semana.setdefault(j, {})
        for sorteo, data_sorteo in sorteos.items():
            s = _normalizar_sorteo_clave(sorteo)
            if not s:
                continue
            sorteos_importados.setdefault(j, set()).add(s)
            for sem_otra, juegos_otra in (fuentes.items() if isinstance(fuentes, dict) else []):
                if sem_otra == semana_key or not isinstance(juegos_otra, dict):
                    continue
                juego_otro = juegos_otra.get(j)
                if isinstance(juego_otro, dict):
                    juego_otro.pop(s, None)
            payload = data_sorteo if isinstance(data_sorteo, dict) else {}
            bucket_semana[j][s] = {
                "recaud": float(payload.get("recaud", 0.0) or 0.0),
                "prem": float(payload.get("prem", 0.0) or 0.0),
            }

    # Si el mismo sorteo se reimporta y ahora NO linkea en Prescripciones,
    # debe eliminarse cualquier prescripción TXT vieja de ese sorteo.
    for juego, sorteos in (prescripciones_por_juego or {}).items():
        j = str(juego).strip()
        if not j or not isinstance(sorteos, dict):
            continue
        for sorteo in sorteos.keys():
            s = _normalizar_sorteo_clave(sorteo)
            if s:
                sorteos_importados.setdefault(j, set()).add(s)

    for juego, sorteos_set in sorteos_importados.items():
        for sem_otra, juegos_otra in (fuentes_presc.items() if isinstance(fuentes_presc, dict) else []):
            if not isinstance(juegos_otra, dict):
                continue
            juego_otro = juegos_otra.get(juego)
            if not isinstance(juego_otro, dict):
                continue
            for s in list(sorteos_set):
                juego_otro.pop(s, None)

    for juego, sorteos in (prescripciones_por_juego or {}).items():
        j = str(juego).strip()
        if not j or not isinstance(sorteos, dict):
            continue
        bucket_presc_semana.setdefault(j, {})
        for sorteo, importe in sorteos.items():
            s = _normalizar_sorteo_clave(sorteo)
            if not s:
                continue
            try:
                bucket_presc_semana[j][s] = float(importe or 0.0)
            except Exception:
                bucket_presc_semana[j][s] = 0.0

    _reconstruir_tickets_globales_desde_fuentes_txt()

    for hook in planilla_refresh_hooks.values():
        if callable(hook):
            hook()

    for hook in planilla_presc_refresh_hooks.values():
        if callable(hook):
            hook()

    recalcular = globals().get("recalcular_y_guardar_totales_txt_semana")
    if callable(recalcular):
        recalcular(semana_key)
    else:
        _refresh_totales()


def reemplazar_reporte_importado(resumen_por_juego: dict, prescripciones_por_juego: dict, semana: int | None = None):
    for juego, sorteos in (resumen_por_juego or {}).items():
        for sorteo, data in (sorteos or {}).items():
            _mergear_resumen_importado(reporte_resumen_por_juego, juego, sorteo, data or {}, con_comision=True)

    semana_key = None
    if semana is not None:
        try:
            sem_n = max(1, min(5, int(semana)))
            semana_key = f"Semana {sem_n}"
            reporte_prescripciones_por_juego_por_semana[semana_key] = {}
        except Exception:
            semana_key = None

    for juego, sorteos in (prescripciones_por_juego or {}).items():
        for sorteo, importe in (sorteos or {}).items():
            _mergear_prescripcion_importada(reporte_prescripciones_por_juego, juego, sorteo, importe)
            if semana_key is not None:
                _mergear_prescripcion_importada(
                    reporte_prescripciones_por_juego_por_semana[semana_key], juego, sorteo, importe
                )

    for hook in planilla_refresh_hooks.values():
        if callable(hook):
            hook()

    for hook in planilla_presc_refresh_hooks.values():
        if callable(hook):
            hook()

    _refresh_totales()


def limpiar_sfa_resumen():
    global _sfa_seen_resumen, _sfa_seen_presc, _sfa_seen_z118
    global _sfa_dirty_planilla_hooks, _sfa_dirty_presc_hooks, _sfa_dirty_totales
    _sfa_seen_resumen = set()
    _sfa_seen_presc = set()
    _sfa_seen_z118 = set()
    _sfa_dirty_planilla_hooks = set()
    _sfa_dirty_presc_hooks = set()
    _sfa_dirty_totales = False
    # Importante: NO borrar los dicts persistidos. Solo reiniciar marcas de esta importación.


def flush_sfa_bulk_updates():
    global _sfa_dirty_planilla_hooks, _sfa_dirty_presc_hooks, _sfa_dirty_totales

    for juego in list(_sfa_dirty_planilla_hooks):
        hook = planilla_refresh_hooks.get(juego)
        if callable(hook):
            try:
                hook()
            except Exception:
                pass

    for juego in list(_sfa_dirty_presc_hooks):
        hook = planilla_presc_refresh_hooks.get(juego)
        if callable(hook):
            try:
                hook()
            except Exception:
                pass

    if _sfa_dirty_totales:
        _refresh_totales()

    _sfa_dirty_planilla_hooks.clear()
    _sfa_dirty_presc_hooks.clear()
    _sfa_dirty_totales = False


def publicar_sfa_resumen(codigo_juego: str, sorteo: str, concepto: str, importe: float):
    global _sfa_seen_resumen, _sfa_seen_presc, _sfa_seen_z118
    global _sfa_dirty_planilla_hooks, _sfa_dirty_presc_hooks, _sfa_dirty_totales

    codigo_juego_norm = _normalizar_codigo_juego_sfa(codigo_juego)
    s = _normalizar_sorteo_clave(sorteo)
    c = _normalizar_concepto_sfa(concepto)

    if c == "Z118":
        juego_key = str(codigo_juego_norm).strip() or "SIN_JUEGO"
        sfa_z118_por_juego.setdefault(juego_key, {})
        clave_z118 = (juego_key, s)
        if clave_z118 not in _sfa_seen_z118:
            sfa_z118_por_juego[juego_key][s] = 0.0
            _sfa_seen_z118.add(clave_z118)
        sfa_z118_por_juego[juego_key][s] = sfa_z118_por_juego[juego_key].get(s, 0.0) + float(importe)
        if sfa_bulk_publish:
            _sfa_dirty_totales = True
        else:
            _refresh_totales()
        return

    juego_tab = _map_codigo_juego_a_tab_planilla(codigo_juego_norm)
    if not juego_tab:
        return

    if c == "Z081":
        sfa_prescripciones_por_juego.setdefault(juego_tab, {})
        clave_presc = (juego_tab, s)
        if clave_presc not in _sfa_seen_presc:
            sfa_prescripciones_por_juego[juego_tab][s] = 0.0
            _sfa_seen_presc.add(clave_presc)
        sfa_prescripciones_por_juego[juego_tab][s] = sfa_prescripciones_por_juego[juego_tab].get(s, 0.0) + float(importe)

        if sfa_bulk_publish:
            _sfa_dirty_presc_hooks.add(juego_tab)
            _sfa_dirty_totales = True
        else:
            hook_presc = planilla_presc_refresh_hooks.get(juego_tab)
            if callable(hook_presc):
                hook_presc()
            _refresh_totales()
        return

    if not sorteo_existe_en_planilla(juego_tab, s):
        return

    campo_por_concepto = {
        "Z005": "recaud",
        "Z013": "comi",
        "Z046": "prem",
    }
    campo = campo_por_concepto.get(c)
    if not campo:
        return

    sfa_resumen_por_juego.setdefault(juego_tab, {})
    sfa_resumen_por_juego[juego_tab].setdefault(s, {"recaud": 0.0, "comi": 0.0, "prem": 0.0})

    clave_resumen = (juego_tab, s, campo)
    if clave_resumen not in _sfa_seen_resumen:
        sfa_resumen_por_juego[juego_tab][s][campo] = 0.0
        _sfa_seen_resumen.add(clave_resumen)

    sfa_resumen_por_juego[juego_tab][s][campo] += float(importe)

    if sfa_bulk_publish:
        _sfa_dirty_planilla_hooks.add(juego_tab)
        _sfa_dirty_totales = True
    else:
        hook = planilla_refresh_hooks.get(juego_tab)
        if callable(hook):
            hook()
        _refresh_totales()


def publicar_filtro_area_recaudacion(juego: str, semana: int, desde: str = "", hasta: str = ""):
    try:
        semana_n = int(semana)
    except Exception:
        semana_n = 0
    if semana_n < 0:
        semana_n = 0
    if semana_n > 5:
        semana_n = 5

    payload = {
        "juego": str(juego or "").strip(),
        "semana": semana_n,
        "desde": str(desde or "").strip(),
        "hasta": str(hasta or "").strip(),
    }

    if semana_n >= 1 and payload["desde"] and payload["hasta"]:
        planilla_rangos_semana_global[semana_n] = (payload["desde"], payload["hasta"])

    planilla_semana_filtro_actual.clear()
    planilla_semana_filtro_actual.update(payload)

    for key, hook in list(planilla_semana_filtro_hooks.items()):
        if not callable(hook):
            continue
        try:
            hook(dict(payload))
        except Exception:
            planilla_semana_filtro_hooks.pop(key, None)


# -----------------------------
# Totales persistidos por semana/columna
# -----------------------------
TOT_FILAS_PLANILLA = [
    "Total ventas",
    "Total comisiones",
    "Total premios",
    "Total prescripcion",
    "Total anticipos",
    "Total topes",
    "Total facuni",
    "Total comision agencia amiga",
]


def _asegurar_planilla_totales_data():
    global planilla_totales_data
    if not isinstance(globals().get("planilla_totales_data"), dict):
        planilla_totales_data = {}
    planilla_totales_data.setdefault("semanas", {})
    planilla_totales_data.setdefault("manuales", {})
    for i in range(1, 6):
        planilla_totales_data["semanas"].setdefault(f"Semana {i}", {})
        planilla_totales_data["manuales"].setdefault(f"Semana {i}", {})
    return planilla_totales_data


def guardar_totales_importados(semana: str, columna: str, valores: dict):
    data = _asegurar_planilla_totales_data()
    semana_key = str(semana or "").strip() or obtener_semana_activa_txt()
    col_raw = str(columna or "").strip().lower()
    if col_raw == "sfa":
        col_key = "sfa"
    elif col_raw == "txt":
        col_key = "txt"
    else:
        col_key = "tobill"

    bucket = data["semanas"].setdefault(semana_key, {})
    manuales_semana = data["manuales"].setdefault(semana_key, {})
    valores_dict = valores if isinstance(valores, dict) else {}

    # Mantener compatibilidad con el layout histórico, pero permitir nuevas
    # filas/columnas agregadas en Totales sin requerir cambios adicionales.
    etiquetas = list(dict.fromkeys([
        *TOT_FILAS_PLANILLA,
        *[k for k in bucket.keys() if isinstance(k, str)],
        *[k for k in manuales_semana.keys() if isinstance(k, str)],
        *[k for k in valores_dict.keys() if isinstance(k, str)],
    ]))

    for etiqueta in etiquetas:
        row = bucket.setdefault(etiqueta, {})
        if not isinstance(row, dict):
            row = {}
            bucket[etiqueta] = row
        try:
            row[col_key] = round(float(valores_dict.get(etiqueta, row.get(col_key, 0.0)) or 0.0), 2)
        except Exception:
            row[col_key] = 0.0

        # Si llega un valor importado para esta fila/columna, ese dato debe
        # prevalecer sobre cualquier edición manual previa para no "pisar"
        # futuras importaciones (ej.: Total facuni al volver a pasar datos).
        if etiqueta in valores_dict:
            fila_manual = manuales_semana.get(etiqueta, {})
            if isinstance(fila_manual, dict) and col_key in fila_manual:
                fila_manual.pop(col_key, None)
                if not fila_manual:
                    manuales_semana.pop(etiqueta, None)

    _refresh_totales()


def obtener_totales_importados_semana(semana: str) -> dict:
    data = _asegurar_planilla_totales_data()
    return dict((data.get("semanas", {}) or {}).get(str(semana or "").strip() or obtener_semana_activa_txt(), {}) or {})
