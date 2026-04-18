from __future__ import annotations

import re
import copy
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tkinter.font as tkfont

import app_state
import app_state
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
    "Quiniela",
    "Quiniela Ya",
    "Poceada",
    "Tombolina",
    "Quini 6",
    "Brinco",
    "Loto",
    "Loto 5",
    "LT",
]

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
    helper = getattr(app_state, "semana_interna_desde_visible", None)
    if callable(helper):
        try:
            return str(helper(valor) or "Semana 1")
        except Exception:
            pass
    txt = str(valor or "").strip()
    if txt in SEMANAS:
        return txt
    m = re.fullmatch(r"(?i)\s*semana\s*(\d+)\s*", txt)
    if m:
        try:
            return f"Semana {max(1, min(5, int(m.group(1))))}"
        except Exception:
            pass
    return "Semana 1"

def _combo_values_semanas() -> list[str]:
    visibles: list[str] = []
    vistos: set[str] = set()

    rangos = getattr(app_state, "planilla_rangos_semana_global", {})
    if not isinstance(rangos, dict):
        return visibles

    for key, value in rangos.items():
        try:
            sem_n = int(str(key).strip())
        except Exception:
            continue
        if sem_n < 1 or sem_n > 5:
            continue

        desde = ""
        hasta = ""
        if isinstance(value, dict):
            desde = str(value.get("desde", "") or "").strip()
            hasta = str(value.get("hasta", "") or "").strip()
        elif isinstance(value, (list, tuple)) and len(value) >= 2:
            desde = str(value[0] or "").strip()
            hasta = str(value[1] or "").strip()

        if not desde and not hasta:
            continue

        visible = _semana_visible(f"Semana {sem_n}")
        if not visible or "--/--/----" in visible or visible in vistos:
            continue
        vistos.add(visible)
        visibles.append(visible)

    return visibles



COLUMNAS = ("sorteos", "importe_ntf", "importe_tobill", "importe_sfa", "diferencias")
TITULOS = {
    "sorteos": "Sorteos",
    "importe_ntf": "Importe NTF",
    "importe_tobill": "Importe Tobill",
    "importe_sfa": "Importe SFA",
    "diferencias": "Diferencias",
}

UMBRAL_DIF_WARN = 20.0
UMBRAL_DIF_DANGER = 40.0
FILA_TOTALES = "Totales"
PANEL_BG = "#D3E6F7"
BADGE_BG = "#F7FAFE"
BADGE_BORDER = "#C7D2E0"
BADGE_TEXT = "#425466"

COLOR_HEAD_SORTEOS = "#E9EEF5"
COLOR_HEAD_TICKETS = "#BFD7F6"
COLOR_HEAD_REPORTE = "#BFEBD3"
COLOR_HEAD_SFA = "#FFD7AA"
COLOR_HEAD_DIF = "#FCE9A7"
COLOR_DIF_WARN_BG = "#FFF2CC"
COLOR_DIF_DANGER_BG = "#FDE2E1"
COLOR_TEXTO = "#0F172A"


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


def _crear_badge_redondeado(parent, textvariable: tk.StringVar, *, width=280, height=28):
    canvas = tk.Canvas(parent, width=width, height=height, bg=PANEL_BG, highlightthickness=0, bd=0, relief="flat")
    font = tkfont.Font(family="Segoe UI", size=9)

    def _redraw(*_args):
        texto = str(textvariable.get() or "").strip()
        canvas.delete("all")
        if not texto:
            return
        pad_x = 14
        badge_w = max(175, min(width, font.measure(texto) + pad_x * 2))
        _rounded_rect(canvas, 2, 2, 2 + badge_w, height - 2, r=10, fill=BADGE_BG, outline=BADGE_BORDER, width=1)
        canvas.create_text(2 + pad_x, height // 2, text=texto, anchor="w", font=font, fill=BADGE_TEXT)

    textvariable.trace_add("write", _redraw)
    _redraw()
    return canvas


def _norm_sorteo(valor) -> str:
    fn = getattr(app_state, "_normalizar_sorteo_clave", None)
    if callable(fn):
        return fn(valor)
    txt = str(valor or "").strip()
    if not txt:
        return ""
    try:
        num = float(txt.replace(",", "."))
        if num.is_integer():
            return str(int(num))
    except Exception:
        pass
    return txt


def _obtener_sorteos_semana(juego: str, semana: int) -> list[str]:
    snaps = getattr(app_state, "obtener_snapshots_area_recaudacion", lambda: {})() or {}
    snap = snaps.get(juego, {}) if isinstance(snaps, dict) else {}
    semanas = snap.get("semanas", {}) if isinstance(snap, dict) else {}
    sorteos = []
    if isinstance(semanas, dict):
        sorteos = semanas.get(str(semana), semanas.get(semana, []))
    out = []
    for s in sorteos if isinstance(sorteos, list) else []:
        n = _norm_sorteo(s)
        if n:
            out.append(n)
    return out


def _to_float(valor) -> float:
    txt = str(valor or "").strip()
    if not txt:
        return 0.0
    txt = txt.replace("$", "").replace(" ", "").replace(",", ".").replace("!", "").replace("🔴", "")
    try:
        return float(txt)
    except Exception:
        return 0.0


def _fmt_importe(valor) -> str:
    txt = str(valor or "").strip()
    if not txt:
        return ""
    return f"{_to_float(txt):.2f}"


def _calcular_diferencia(importe_ntf, importe_sfa) -> str:
    dif = _to_float(importe_sfa) - _to_float(importe_ntf)
    return f"{dif:.2f}"


def _fmt_diferencia(importe_ntf, importe_sfa) -> str:
    dif = _to_float(importe_sfa) - _to_float(importe_ntf)
    return f"{dif:.2f}"


def _seed_data() -> dict:
    data = getattr(app_state, "planilla_agencia_amiga_data", {}) or {}
    if not isinstance(data, dict):
        data = {}
    data.setdefault("juegos", {})
    return data


def _ensure_bucket(data: dict, juego: str, semana: int) -> dict:
    juegos = data.setdefault("juegos", {})
    juego_map = juegos.setdefault(juego, {})
    return juego_map.setdefault(str(semana), {})


def _parse_txt_ntf(path: str) -> dict[str, str]:
    def _extraer(texto: str, desde: int, hasta: int) -> str:
        if not texto:
            return ""
        if len(texto) < desde:
            return ""
        return str(texto[desde - 1:hasta]).strip()

    def _limpiar_numerico(texto: str) -> str:
        return "".join(ch for ch in str(texto or "").strip() if ch.isdigit() or ch == "-")

    def _convertir_importe_centavos(texto: str) -> float:
        limpio = _limpiar_numerico(texto)
        if not limpio or limpio == "-":
            return 0.0
        try:
            return float(limpio) / 100.0
        except Exception:
            return 0.0

    def _tiene_dato(texto: str) -> bool:
        return bool(str(texto or "").strip())

    def _tiene_agencia_amiga_valida(texto: str) -> bool:
        raw = str(texto or "").strip()
        if not raw:
            return False
        limpio = _limpiar_numerico(raw)
        if not limpio:
            return False
        try:
            return float(limpio) != 0.0
        except Exception:
            return False

    def _nombre_juego(codigo: str) -> str:
        c = str(codigo or "").strip()
        mapa = {
            "80": "Quiniela",
            "82": "Poceada",
            "79": "Quiniela Ya",
            "74": "Tombolina",
            "9": "Loto",
            "09": "Loto",
            "5": "Loto 5",
            "05": "Loto 5",
        }
        return mapa.get(c, "")

    def _calcular_comision_por_juego(juego: str, total: float) -> float:
        j = str(juego or "").strip().upper()
        if j in ("QUINIELA", "POCEADA", "QUINIELA YA", "TOMBOLINA"):
            return float(total) * 0.08
        if j in ("LOTO", "LOTO 5"):
            return float(total) * 0.04
        return 0.0

    # Acumulación por juego+sorteo para poder ubicar importe en su fila exacta.
    totales: dict[tuple[str, str], float] = {}
    canceladas: dict[tuple[str, str], float] = {}

    with open(path, "r", encoding="utf-8", errors="replace") as f:
        for linea in f:
            if len(linea) < 131:
                continue

            juego_cod = _extraer(linea, 3, 4)
            juego = _nombre_juego(juego_cod)
            if not juego:
                continue

            sorteo = _norm_sorteo(_extraer(linea, 5, 10))
            if not sorteo:
                continue

            fecha_cancelacion = _extraer(linea, 71, 78)
            agencia_amiga_web = _extraer(linea, 114, 121)
            valor_apuesta = _convertir_importe_centavos(_extraer(linea, 122, 131))

            if not _tiene_agencia_amiga_valida(agencia_amiga_web):
                continue

            key = (juego, sorteo)
            totales[key] = totales.get(key, 0.0) + valor_apuesta

            if _tiene_dato(fecha_cancelacion):
                canceladas[key] = canceladas.get(key, 0.0) + valor_apuesta

    out: dict[tuple[str, str], str] = {}
    for key, total_base in totales.items():
        juego, _sorteo = key
        total_cancelado = canceladas.get(key, 0.0)
        neto = _calcular_comision_por_juego(juego, total_base) - _calcular_comision_por_juego(juego, total_cancelado)
        out[key] = f"{neto:.2f}"

    return out




def _mapear_sfa_z118_a_juego_planilla() -> dict[str, dict[str, float]]:
    """Convierte app_state.sfa_z118_por_juego (código SFA) a nombre de juego de planilla."""
    src = getattr(app_state, "sfa_z118_por_juego", {}) or {}
    out: dict[str, dict[str, float]] = {}
    map_juego = getattr(app_state, "_map_codigo_juego_a_tab_planilla", None)

    for codigo_juego, sorteos in (src.items() if isinstance(src, dict) else []):
        if not isinstance(sorteos, dict):
            continue

        juego_planilla = map_juego(codigo_juego) if callable(map_juego) else str(codigo_juego or "").strip()
        juego_planilla = str(juego_planilla or "").strip()
        if not juego_planilla:
            continue

        out.setdefault(juego_planilla, {})
        for sorteo, importe in sorteos.items():
            s = _norm_sorteo(sorteo)
            if not s:
                continue
            out[juego_planilla][s] = out[juego_planilla].get(s, 0.0) + float(importe or 0.0)

    return out

def build_agencia_amiga(fr_seccion: ttk.Frame, estado_var):
    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(1, weight=1)

    top = ttk.Frame(fr_seccion, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 4))
    top.columnconfigure(8, weight=1)

    ttk.Label(top, text="Juego:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")
    juego_var = tk.StringVar(value=JUEGOS[0])
    combo_juego = ttk.Combobox(top, state="readonly", values=JUEGOS, textvariable=juego_var, width=30, style="Planilla.Toolbar.TCombobox")
    combo_juego.grid(row=0, column=1, sticky="w", padx=(8, 10))

    ttk.Label(top, text="Semana:", style="PanelLabel.TLabel").grid(row=0, column=2, sticky="e", padx=(12, 6))
    semana_var = tk.StringVar(value="")
    combo_semana = ttk.Combobox(top, state="disabled", values=_combo_values_semanas(), textvariable=semana_var, width=28, style="Planilla.Toolbar.TCombobox")
    combo_semana.grid(row=0, column=3, sticky="w")
    rango_semana_var = tk.StringVar(value="Del: --/--/---- al: --/--/----")
    badge_rango_semana = _crear_badge_redondeado(top, rango_semana_var, width=280, height=28)
    badge_rango_semana.grid_remove()

    tabla_wrap = ttk.Frame(fr_seccion, style="Panel.TFrame")
    tabla_wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(4, 8))
    tabla_wrap.columnconfigure(0, weight=1)
    tabla_wrap.rowconfigure(1, weight=1)

    header_canvas = tk.Canvas(tabla_wrap, height=30, bg="#E9EEF5", highlightthickness=0)
    header_canvas.grid(row=0, column=0, sticky="ew")

    tree = ttk.Treeview(tabla_wrap, columns=COLUMNAS, show="tree", height=16)
    tree.grid(row=1, column=0, sticky="nsew")
    tree.column("#0", width=0, stretch=False)
    tree.heading("#0", text="")
    vsb = ttk.Scrollbar(tabla_wrap, orient="vertical")
    vsb.grid(row=1, column=1, sticky="ns")

    # Estilo visual profesional y consistente con el resto de secciones.
    tree.tag_configure("even", background="#FFFFFF", foreground=COLOR_TEXTO)
    tree.tag_configure("odd", background="#F8FAFC", foreground=COLOR_TEXTO)
    tree.tag_configure("totales", background="#FFF9CC", foreground="#1F2937", font=("Segoe UI Semibold", 9))
    tree.tag_configure("diff_warn", background=COLOR_DIF_WARN_BG, foreground=COLOR_TEXTO)
    tree.tag_configure("diff_danger", background=COLOR_DIF_DANGER_BG, foreground=COLOR_TEXTO)

    base_widths = {
        "sorteos": 160,
        "importe_ntf": 160,
        "importe_tobill": 160,
        "importe_sfa": 160,
        "diferencias": 160,
    }
    for col in COLUMNAS:
        tree.heading(col, text="")
        anchor = "center"
        tree.column(col, width=base_widths[col], minwidth=110, stretch=False, anchor=anchor)

    ntf_btn_win_id = None

    def _ajustar_columnas_a_ancho():
        min_col = 110
        total_base = sum(base_widths[c] for c in COLUMNAS)
        ancho_disponible = max(total_base, int(tree.winfo_width() or total_base))

        acumulado = 0
        for idx, col in enumerate(COLUMNAS):
            if idx == len(COLUMNAS) - 1:
                w = max(min_col, ancho_disponible - acumulado)
            else:
                proporcion = base_widths[col] / total_base if total_base else 1 / len(COLUMNAS)
                w = max(min_col, int(ancho_disponible * proporcion))
                acumulado += w
            tree.column(col, width=w, minwidth=min_col, stretch=False)

    def _draw_headers():
        nonlocal ntf_btn_win_id
        header_canvas.delete("all")
        if ntf_btn_win_id is not None:
            try:
                header_canvas.delete(ntf_btn_win_id)
            except Exception:
                pass
            ntf_btn_win_id = None

        color_por_col = {
            "sorteos": COLOR_HEAD_SORTEOS,
            "importe_ntf": COLOR_HEAD_TICKETS,
            "importe_tobill": COLOR_HEAD_REPORTE,
            "importe_sfa": COLOR_HEAD_SFA,
            "diferencias": COLOR_HEAD_DIF,
        }

        x = 0
        for col in COLUMNAS:
            w = int(tree.column(col, "width") or base_widths[col])
            color = color_por_col.get(col, COLOR_HEAD_SORTEOS)
            header_canvas.create_rectangle(x, 1, x + w, 29, fill=color, outline="#AAB6C5", width=1)
            if col == "importe_ntf":
                btn = tk.Button(
                    header_canvas,
                    text=TITULOS[col],
                    command=_importar_ntf_txt,
                    relief="flat",
                    bd=0,
                    font=("Segoe UI Semibold", 9),
                    fg="#0F172A",
                    bg=COLOR_HEAD_TICKETS,
                    activebackground=COLOR_HEAD_TICKETS,
                    activeforeground="#0F172A",
                    cursor="hand2",
                    padx=1,
                    pady=0,
                )
                ntf_btn_win_id = header_canvas.create_window(x + (w / 2), 15, window=btn, anchor="center", width=max(90, w - 8), height=22)
            else:
                header_canvas.create_text(x + w / 2, 15, text=TITULOS[col], font=("Segoe UI Semibold", 9), fill="#1F2937")
            x += w
        header_canvas.configure(scrollregion=(0, 0, max(1, x), 30))
        header_canvas.xview_moveto(0)

    diff_labels: list[tk.Label] = []

    def _fondo_fila(iid: str) -> str:
        if iid in tree.selection():
            return "#DCEBFF"
        tags = tree.item(iid, "tags")
        if "totales" in tags:
            return "#FFF9CC"
        if "diff_danger" in tags:
            return COLOR_DIF_DANGER_BG
        if "diff_warn" in tags:
            return COLOR_DIF_WARN_BG
        if "odd" in tags:
            return "#F8FAFC"
        return "#FFFFFF"

    def _clear_diff_labels(*, destroy: bool = False):
        for label in diff_labels:
            try:
                if destroy:
                    label.destroy()
                else:
                    label.place_forget()
            except Exception:
                pass
        if destroy:
            diff_labels.clear()

    def _ensure_diff_label_pool(count: int):
        while len(diff_labels) < count:
            diff_labels.append(
                tk.Label(
                    tree,
                    anchor="center",
                    padx=6,
                    pady=0,
                    borderwidth=0,
                    highlightthickness=0,
                    font=("Segoe UI", 9),
                )
            )

    def _render_diferencias_labels():
        _clear_diff_labels()
        return

    data = _seed_data()
    app_state.planilla_agencia_amiga_data = data
    def _aplicar_estilo_filas():
        idx_visible = 0
        for iid in tree.get_children():
            vals = [str(v) for v in tree.item(iid, "values")]
            if vals and str(vals[0]).strip().lower() == FILA_TOTALES.lower():
                tree.item(iid, tags=("totales",))
                continue

            zebra = "even" if (idx_visible % 2 == 0) else "odd"
            dif = _to_float(vals[4]) if len(vals) > 4 else 0.0
            dif_abs = abs(dif)

            if dif_abs > UMBRAL_DIF_DANGER:
                tree.item(iid, tags=("diff_danger",))
            elif dif_abs > UMBRAL_DIF_WARN:
                tree.item(iid, tags=("diff_warn",))
            else:
                tree.item(iid, tags=(zebra,))
            idx_visible += 1
        _render_diferencias_labels()

    def _semana_num() -> int:
        try:
            interna = _semana_interna(str(semana_var.get() or combo_semana.get() or "").strip())
            return int(str(interna).replace("Semana", "").strip() or "1")
        except Exception:
            return 1

    def _guardar_fila_actual(iid: str):
        vals = tree.item(iid, "values")
        if len(vals) != len(COLUMNAS):
            return
        sorteo = _norm_sorteo(vals[0])
        if not sorteo or sorteo.lower() == FILA_TOTALES.lower():
            return
        juego = juego_var.get().strip()
        semana = _semana_num()
        dif = _calcular_diferencia(vals[1], vals[3])

        bucket = _ensure_bucket(data, juego, semana)
        bucket[sorteo] = {
            "importe_ntf": _fmt_importe(vals[1]),
            "importe_tobill": _fmt_importe(vals[2]),
            "importe_sfa": _fmt_importe(vals[3]),
            "diferencias": dif,
        }

    def _actualizar_rango_del_al():
        if hasattr(app_state, "texto_rango_semana_global"):
            rango_semana_var.set(app_state.texto_rango_semana_global(_semana_num()))
        else:
            rango_semana_var.set("Del: --/--/---- al: --/--/----")

    def _guardar_filtro_actual():
        ui = data.setdefault("_ui", {})
        if not isinstance(ui, dict):
            ui = {}
            data["_ui"] = ui
        ui["juego"] = juego_var.get().strip()
        ui["semana"] = _semana_num()

    def _sync_totales_txt_agencia_amiga(semana_val=None):
        try:
            semana_n = int(str(semana_val or _semana_num()).replace("Semana", "").strip() or "1")
        except Exception:
            semana_n = _semana_num()
        semana_n = max(1, min(5, int(semana_n or 1)))
        sem_key = str(semana_n)
        total_ntf = 0.0

        juegos = data.get("juegos", {}) if isinstance(data, dict) else {}
        for semanas in (juegos.values() if isinstance(juegos, dict) else []):
            if not isinstance(semanas, dict):
                continue
            bucket = semanas.get(sem_key, {})
            if not isinstance(bucket, dict):
                continue
            for fila in bucket.values():
                if not isinstance(fila, dict):
                    continue
                total_ntf += _to_float(fila.get("importe_ntf", ""))

        recalcular = getattr(app_state, "recalcular_y_guardar_totales_txt_semana", None)
        if callable(recalcular):
            recalcular(f"Semana {semana_n}", {"Total comision agencia amiga": round(total_ntf, 2)})
        else:
            guardar = getattr(app_state, "guardar_totales_importados", None)
            if callable(guardar):
                guardar(f"Semana {semana_n}", "txt", {"Total comision agencia amiga": round(total_ntf, 2)})


    def _insertar_fila_totales():
        for iid in list(tree.get_children()):
            vals = list(tree.item(iid, "values"))
            if vals and str(vals[0]).strip().lower() == FILA_TOTALES.lower():
                tree.delete(iid)
        total_ntf = 0.0
        total_tobill = 0.0
        total_sfa = 0.0
        total_dif = 0.0
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            if not vals or str(vals[0]).strip().lower() == FILA_TOTALES.lower():
                continue
            total_ntf += _to_float(vals[1] if len(vals) > 1 else 0.0)
            total_tobill += _to_float(vals[2] if len(vals) > 2 else 0.0)
            total_sfa += _to_float(vals[3] if len(vals) > 3 else 0.0)
            total_dif += _to_float(vals[4] if len(vals) > 4 else 0.0)

        tree.insert(
            "",
            "end",
            values=(FILA_TOTALES, f"{total_ntf:.2f}", f"{total_tobill:.2f}", f"{total_sfa:.2f}", f"{total_dif:.2f}"),
            tags=("totales",),
        )
        _aplicar_estilo_filas()
        _sync_totales_txt_agencia_amiga(_semana_num())

    def _sync_importes_desde_reporte(juego: str, semana: int, bucket: dict):
        semana_key = f"Semana {int(semana or 1)}"
        tobill_por_semana = getattr(app_state, "reporte_agencia_amiga_tobill_por_juego_por_semana", {}) or {}
        sfa118_por_semana = getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego_por_semana", {}) or {}
        tobill = (
            tobill_por_semana.get(semana_key, {})
            if isinstance(tobill_por_semana, dict) and isinstance(tobill_por_semana.get(semana_key, {}), dict)
            else getattr(app_state, "reporte_agencia_amiga_tobill_por_juego", {}) or {}
        )
        sfa118 = (
            sfa118_por_semana.get(semana_key, {})
            if isinstance(sfa118_por_semana, dict) and isinstance(sfa118_por_semana.get(semana_key, {}), dict)
            else getattr(app_state, "reporte_agencia_amiga_sfa_118_por_juego", {}) or {}
        )
        sfa118_txt_json = _mapear_sfa_z118_a_juego_planilla()

        tobill_juego = tobill.get(juego, {}) if isinstance(tobill, dict) else {}
        sfa_juego = sfa118.get(juego, {}) if isinstance(sfa118, dict) else {}
        sfa_juego_txt_json = sfa118_txt_json.get(juego, {}) if isinstance(sfa118_txt_json, dict) else {}
        if not isinstance(tobill_juego, dict):
            tobill_juego = {}
        if not isinstance(sfa_juego, dict):
            sfa_juego = {}

        sorteos_semana = _obtener_sorteos_semana(juego, semana)
        for s in sorteos_semana:
            item = bucket.setdefault(s, {}) if isinstance(bucket, dict) else {}
            if not isinstance(item, dict):
                item = {}
                bucket[s] = item

            if s in tobill_juego:
                item["importe_tobill"] = f"{float(tobill_juego.get(s, 0.0) or 0.0):.2f}"
            if s in sfa_juego:
                item["importe_sfa"] = f"{float(sfa_juego.get(s, 0.0) or 0.0):.2f}"
            elif s in sfa_juego_txt_json:
                item["importe_sfa"] = f"{float(sfa_juego_txt_json.get(s, 0.0) or 0.0):.2f}"

    _week_sync_guard = {"active": False}

    def _publicar_semana_global_desde_combo(semana_txt: str, juego_txt: str = ""):
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

        try:
            payload_actual = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            juego_payload = str(payload_actual.get("juego", "") or "").strip() if isinstance(payload_actual, dict) else ""
        except Exception:
            juego_payload = ""

        juego_final = str(juego_txt or "").strip() or juego_payload
        publicar = getattr(app_state, "publicar_filtro_area_recaudacion", None)
        if callable(publicar):
            try:
                publicar(juego_final, semana_n, desde, hasta)
            except Exception:
                pass

    def _cargar_grilla(*_args):
        _guardar_filtro_actual()
        juego = juego_var.get().strip()
        semana = _semana_num()
        _actualizar_rango_del_al()

        for iid in tree.get_children():
            tree.delete(iid)

        sorteos = _obtener_sorteos_semana(juego, semana)
        bucket = _ensure_bucket(data, juego, semana)
        _sync_importes_desde_reporte(juego, semana, bucket)

        known = set()
        for s in sorteos:
            item = bucket.get(s, {}) if isinstance(bucket.get(s), dict) else {}
            iid = tree.insert(
                "",
                "end",
                values=(
                    s,
                    _fmt_importe(item.get("importe_ntf", "")),
                    _fmt_importe(item.get("importe_tobill", "")),
                    _fmt_importe(item.get("importe_sfa", "")),
                    _fmt_diferencia(item.get("importe_ntf", ""), item.get("importe_sfa", "")),
                ),
            )
            known.add(s)

        extras = sorted([k for k in bucket.keys() if str(k) not in known], key=lambda x: int(_norm_sorteo(x) or 0))
        for s in extras:
            item = bucket.get(s, {}) if isinstance(bucket.get(s), dict) else {}
            tree.insert(
                "",
                "end",
                values=(
                    s,
                    _fmt_importe(item.get("importe_ntf", "")),
                    _fmt_importe(item.get("importe_tobill", "")),
                    _fmt_importe(item.get("importe_sfa", "")),
                    _fmt_diferencia(item.get("importe_ntf", ""), item.get("importe_sfa", "")),
                ),
            )

        _insertar_fila_totales()

        if estado_var is not None:
            total_sorteos = max(0, len(tree.get_children()) - 1)
            estado_var.set(f"Agencia Amiga: {juego} - Semana {semana} ({total_sorteos} sorteos).")

    editable_col_indices = {0, 1, 2, 3}
    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)

    def _normalizar_valor_celda(idx: int, raw: str) -> str:
        valor = str(raw or "").strip()
        if idx == 0:
            return _norm_sorteo(valor)
        if idx in (1, 2, 3):
            return _fmt_importe(valor)
        return valor

    def _editar_celda(event):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return

        row = tree.identify_row(event.y)
        col_id = tree.identify_column(event.x)
        if not row or not col_id:
            return

        idx = int(col_id.replace("#", "")) - 1
        if idx < 0 or idx >= len(COLUMNAS):
            return
        clipboard_state["cell"] = (row, idx)
        row_vals = list(tree.item(row, "values"))
        if row_vals and str(row_vals[0]).strip().lower() == FILA_TOTALES.lower():
            return
        if COLUMNAS[idx] == "diferencias":
            return

        x, y, w, h = tree.bbox(row, col_id)
        old_vals = list(tree.item(row, "values"))
        while len(old_vals) < len(COLUMNAS):
            old_vals.append("")

        editor = ttk.Entry(tree)
        editor.place(x=x, y=y, width=w, height=h)
        editor.insert(0, old_vals[idx])
        editor.focus_set()
        editor.select_range(0, "end")

        def _commit(_e=None):
            before_vals = list(tree.item(row, "values"))
            old_vals[idx] = _normalizar_valor_celda(idx, editor.get())
            old_vals[4] = _fmt_diferencia(old_vals[1], old_vals[3])
            if old_vals != before_vals:
                push_undo_rows(undo_state, tree, [(row, list(before_vals))], meta={"juego": juego_var.get().strip(), "semana": combo_semana.get().strip()})
                tree.item(row, values=old_vals)
                _guardar_fila_actual(row)
                _insertar_fila_totales()
                _render_diferencias_labels()
            editor.destroy()

        def _cancel(_e=None):
            editor.destroy()

        editor.bind("<Return>", _commit)
        editor.bind("<FocusOut>", _commit)
        editor.bind("<Escape>", _cancel)

    def _copiar_celdas(_evt=None):
        iid, col_idx = get_anchor_cell(tree, clipboard_state, default_col=0)
        if not iid or col_idx not in editable_col_indices:
            return "break"
        seleccion = ordered_selected_rows(tree)
        row_ids = seleccion if len(seleccion) > 1 and iid in seleccion else [iid]
        matrix = []
        for row_iid in row_ids:
            vals = [str(v) for v in tree.item(row_iid, "values")]
            if vals and str(vals[0]).strip().lower() == FILA_TOTALES.lower():
                continue
            while len(vals) < len(COLUMNAS):
                vals.append("")
            matrix.append([vals[col_idx]])
        set_clipboard_matrix(tree, matrix)
        return "break"

    def _pegar_celdas(_evt=None):
        matrix = get_clipboard_matrix(tree)
        if not matrix:
            return "break"
        anchor_iid, anchor_col = get_anchor_cell(tree, clipboard_state, default_col=0)
        row_ids = [iid for iid in tree.get_children() if str((tree.item(iid, "values") or [""])[0]).strip().lower() != FILA_TOTALES.lower()]
        if not anchor_iid or anchor_iid not in row_ids:
            return "break"
        start = row_ids.index(anchor_iid)
        hubo_cambios = False
        undo_rows = []
        for r_off, row_data in enumerate(matrix):
            row_pos = start + r_off
            if row_pos >= len(row_ids):
                break
            iid = row_ids[row_pos]
            vals = list(tree.item(iid, "values"))
            before_vals = list(vals)
            while len(vals) < len(COLUMNAS):
                vals.append("")
            while len(before_vals) < len(COLUMNAS):
                before_vals.append("")
            row_changed = False
            for c_off, cell_raw in enumerate(row_data):
                col_idx = anchor_col + c_off
                if col_idx not in editable_col_indices:
                    continue
                vals[col_idx] = _normalizar_valor_celda(col_idx, cell_raw)
                row_changed = True
            if row_changed:
                vals[4] = _fmt_diferencia(vals[1], vals[3])
                if vals != before_vals:
                    undo_rows.append((iid, before_vals))
                    tree.item(iid, values=vals)
                    _guardar_fila_actual(iid)
                    hubo_cambios = True
        if hubo_cambios:
            push_undo_rows(undo_state, tree, undo_rows, meta={"juego": juego_var.get().strip(), "semana": combo_semana.get().strip()})
            _insertar_fila_totales()
            _render_diferencias_labels()
        return "break"

    def _deshacer_celdas(_evt=None):
        snapshot = pop_undo_snapshot(undo_state)
        if not snapshot:
            return "break"
        meta = snapshot.get("meta", {}) if isinstance(snapshot, dict) else {}
        juego_actual = juego_var.get().strip()
        semana_actual = combo_semana.get().strip()
        juego_undo = str(meta.get("juego", juego_actual) or juego_actual)
        semana_undo = str(meta.get("semana", semana_actual) or semana_actual)
        need_reload = (juego_undo and juego_undo != juego_actual) or (semana_undo and semana_undo != semana_actual)
        if juego_undo and juego_undo != juego_actual:
            combo_juego.set(juego_undo)
        if semana_undo and semana_undo != semana_actual:
            combo_semana.set(_semana_visible(semana_undo))
        if need_reload:
            _cargar_grilla()
        restore_undo_snapshot(tree, snapshot)
        for iid, _vals in (snapshot.get("rows", []) or []):
            _guardar_fila_actual(str(iid))
        _insertar_fila_totales()
        _render_diferencias_labels()
        return "break"

    def _importar_ntf_txt():
        path = filedialog.askopenfilename(
            title="Importar Importe NTF (TXT)",
            filetypes=[("TXT", "*.txt"), ("Todos", "*.*")],
        )
        if not path:
            return

        try:
            mapeo = _parse_txt_ntf(path)
        except Exception as e:
            messagebox.showerror("Importe NTF", f"No pude leer el archivo TXT.\n{e}")
            return

        if not mapeo:
            messagebox.showwarning("Importe NTF", "No se detectaron filas válidas para Agencia Amiga.")
            return

        juego_actual = juego_var.get().strip()
        actualizados = 0
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            while len(vals) < len(COLUMNAS):
                vals.append("")
            if str(vals[0]).strip().lower() == FILA_TOTALES.lower():
                continue
            sorteo = _norm_sorteo(vals[0])
            key = (juego_actual, sorteo)
            if key in mapeo:
                vals[1] = _fmt_importe(mapeo[key])
                vals[4] = _fmt_diferencia(vals[1], vals[3])
                tree.item(iid, values=vals)
                _guardar_fila_actual(iid)
                actualizados += 1

        _insertar_fila_totales()

        if estado_var is not None:
            estado_var.set(f"Importe NTF: {actualizados} sorteos actualizados desde TXT.")

    def _on_tree_yscroll(first, last):
        vsb.set(first, last)
        _render_diferencias_labels()

    vsb.configure(command=tree.yview)
    tree.configure(yscrollcommand=_on_tree_yscroll)

    tree.bind("<Double-1>", _editar_celda)
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")
    def _on_tree_configure(_e=None):
        _ajustar_columnas_a_ancho()
        _draw_headers()
        _render_diferencias_labels()

    tree.bind("<Configure>", _on_tree_configure)
    tree.bind("<<TreeviewSelect>>", lambda _e: _render_diferencias_labels(), add="+")
    combo_juego.bind("<<ComboboxSelected>>", _cargar_grilla)

    def _on_semana_combo_selected(_e=None):
        if not _week_sync_guard["active"]:
            _publicar_semana_global_desde_combo(combo_semana.get(), juego_var.get().strip())
        _cargar_grilla()

    combo_semana.bind("<<ComboboxSelected>>", _on_semana_combo_selected)

    if not hasattr(app_state, "planilla_semana_filtro_hooks"):
        app_state.planilla_semana_filtro_hooks = {}

    def _aplicar_filtro_area_recaudacion(payload: dict):
        juego = str((payload or {}).get("juego", "") or "").strip()
        values = _combo_values_semanas()
        try:
            combo_semana.configure(values=values)
        except Exception:
            pass

        try:
            semana = int((payload or {}).get("semana", 0) or 0)
        except Exception:
            semana = 0

        if juego in JUEGOS:
            juego_var.set(juego)

        _week_sync_guard["active"] = True
        try:
            if semana < 1 or not values:
                semana_var.set("")
                try:
                    combo_semana.set("")
                    combo_semana.configure(state="disabled")
                except Exception:
                    pass
                _cargar_grilla()
                return

            semana = max(1, min(5, semana))
            visible = _semana_visible(f"Semana {semana}")
            semana_var.set(visible)
            try:
                combo_semana.configure(state="readonly")
                combo_semana.set(visible)
            except Exception:
                pass
            _cargar_grilla()
        finally:
            _week_sync_guard["active"] = False

    app_state.planilla_semana_filtro_hooks["agencia_amiga"] = _aplicar_filtro_area_recaudacion

    if not hasattr(app_state, "planilla_bundle_snapshot_hooks"):
        app_state.planilla_bundle_snapshot_hooks = {}

    def _snapshot() -> dict:
        snapshot = copy.deepcopy(data if isinstance(data, dict) else {})
        juegos_out = snapshot.setdefault("juegos", {})
        if not isinstance(juegos_out, dict):
            juegos_out = {}
            snapshot["juegos"] = juegos_out

        for juego in JUEGOS:
            juego_map = juegos_out.setdefault(juego, {})
            if not isinstance(juego_map, dict):
                juego_map = {}
                juegos_out[juego] = juego_map

            for semana in range(1, 6):
                key_semana = str(semana)
                bucket = juego_map.get(key_semana, {})
                if not isinstance(bucket, dict):
                    bucket = {}

                _sync_importes_desde_reporte(juego, semana, bucket)

                # Garantiza persistir todos los sorteos visibles de Área Recaudación
                # para cada juego/semana, incluso si el usuario no navegó esa semana.
                for sorteo in _obtener_sorteos_semana(juego, semana):
                    item = bucket.setdefault(sorteo, {})
                    if not isinstance(item, dict):
                        item = {}
                        bucket[sorteo] = item
                    item.setdefault("importe_ntf", _fmt_importe(item.get("importe_ntf", "")))
                    item.setdefault("importe_tobill", _fmt_importe(item.get("importe_tobill", "")))
                    item.setdefault("importe_sfa", _fmt_importe(item.get("importe_sfa", "")))
                    item["diferencias"] = _fmt_diferencia(item.get("importe_ntf", ""), item.get("importe_sfa", ""))

                juego_map[key_semana] = bucket

        return snapshot

    app_state.planilla_bundle_snapshot_hooks["agencia_amiga"] = _snapshot

    if not hasattr(app_state, "planilla_agencia_amiga_load_hooks"):
        app_state.planilla_agencia_amiga_load_hooks = {}

    def _load_hook(payload: dict):
        nuevo = payload if isinstance(payload, dict) else {"juegos": {}}
        data.clear()
        data.update(copy.deepcopy(nuevo))
        data.setdefault("juegos", {})

        values = _combo_values_semanas()
        try:
            combo_semana.configure(values=values)
        except Exception:
            pass

        payload_filtro = getattr(app_state, "planilla_semana_filtro_actual", {})
        if isinstance(payload_filtro, dict):
            try:
                _aplicar_filtro_area_recaudacion(dict(payload_filtro))
                return
            except Exception:
                pass

        if values:
            try:
                combo_semana.configure(state="readonly")
            except Exception:
                pass
        else:
            try:
                combo_semana.set("")
                combo_semana.configure(state="disabled")
            except Exception:
                pass
        _cargar_grilla()

    app_state.planilla_agencia_amiga_load_hooks["agencia_amiga"] = _load_hook
    app_state.planilla_agencia_amiga_refresh_hooks["agencia_amiga"] = _cargar_grilla

    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    def _refresh_visual_agencia_amiga():
        try:
            _clear_diff_labels(destroy=True)
        except Exception:
            pass

        try:
            _cargar_grilla()
        except Exception:
            pass

        try:
            tree.after_idle(_draw_headers)
        except Exception:
            try:
                _draw_headers()
            except Exception:
                pass

    app_state.planilla_visual_refresh_hooks["AGENCIA AMIGA"] = _refresh_visual_agencia_amiga


    payload_inicial = getattr(app_state, "planilla_semana_filtro_actual", {})
    ui_guardada = data.get("_ui", {}) if isinstance(data, dict) else {}

    _aplicar_filtro_area_recaudacion(payload_inicial if isinstance(payload_inicial, dict) else {})

    if isinstance(ui_guardada, dict):
        juego_ui = str(ui_guardada.get("juego", "") or "").strip()
        if juego_ui in JUEGOS:
            juego_var.set(juego_ui)

        try:
            sem_ini = int(ui_guardada.get("semana", 0) or 0)
        except Exception:
            sem_ini = 0

        values = _combo_values_semanas()
        try:
            combo_semana.configure(values=values)
        except Exception:
            pass

        if 1 <= sem_ini <= 5 and values:
            visible = _semana_visible(f"Semana {sem_ini}")
            semana_var.set(visible)
            try:
                combo_semana.configure(state="readonly")
                combo_semana.set(visible)
            except Exception:
                pass

        _cargar_grilla()

    _ajustar_columnas_a_ancho()
    _draw_headers()
