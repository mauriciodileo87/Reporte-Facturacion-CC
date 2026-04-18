from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
import re

import app_state
from json_sfa import leer_json_sfa
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

FILAS_BASE = ["CASH-IN", "CASH-OUT"]
SEMANAS = ["Semana 1", "Semana 2", "Semana 3", "Semana 4", "Semana 5"]

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



DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado"]


def _lighten(hex_color: str, factor: float = 1.08) -> str:
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = max(0, min(255, int(r * factor)))
    g = max(0, min(255, int(g * factor)))
    b = max(0, min(255, int(b * factor)))
    return f"#{r:02X}{g:02X}{b:02X}"


def _darken(hex_color: str, factor: float = 0.82) -> str:
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = max(0, min(255, int(r * factor)))
    g = max(0, min(255, int(g * factor)))
    b = max(0, min(255, int(b * factor)))
    return f"#{r:02X}{g:02X}{b:02X}"


def _aplicar_hover_boton(btn: tk.Button, base_bg: str):
    hover_bg = _lighten(base_bg, 1.08)

    def _on_enter(_evt=None):
        btn.configure(bg=hover_bg, activebackground=hover_bg, relief="raised", bd=1)

    def _on_leave(_evt=None):
        btn.configure(bg=base_bg, activebackground=base_bg, relief="flat", bd=0)

    btn.bind("<Enter>", _on_enter, add="+")
    btn.bind("<Leave>", _on_leave, add="+")


def _get_ttk_background(widget: tk.Widget) -> str:
    try:
        style = ttk.Style(widget)
        bg = style.lookup("TFrame", "background")
        if bg:
            return bg
    except Exception:
        pass
    try:
        return widget.winfo_toplevel().cget("bg")
    except Exception:
        return "#D9D9D9"


def _parse_importe_fijo(raw: str) -> float:
    """
    Campo Monto (pos 34-51, largo 18) según layout recibido.
    - Si viene sin separadores: se interpreta como entero con 2 decimales implícitos.
    - Si viene con coma/punto: se respeta el separador decimal.
    """
    s = str(raw or "").strip()
    if not s:
        return 0.0

    neg = False
    if s.startswith("-"):
        neg = True
        s = s[1:].strip()
    elif s.endswith("-"):
        neg = True
        s = s[:-1].strip()

    if "," in s or "." in s:
        # Usar el último separador como decimal y limpiar el resto como miles.
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        sep_idx = max(last_comma, last_dot)
        int_part = "".join(ch for ch in s[:sep_idx] if ch.isdigit())
        dec_part = "".join(ch for ch in s[sep_idx + 1 :] if ch.isdigit())
        num_str = f"{int_part}.{dec_part or '0'}" if int_part else f"0.{dec_part or '0'}"
        try:
            v = float(num_str)
            return -v if neg else v
        except Exception:
            return 0.0

    only_digits = "".join(ch for ch in s if ch.isdigit())
    if not only_digits:
        return 0.0

    try:
        v = int(only_digits) / 100.0
        return -v if neg else v
    except Exception:
        return 0.0


def _normalizar_movimiento(raw: str) -> str:
    s = "".join(ch for ch in str(raw or "").upper() if ch.isalnum())
    if not s:
        return ""
    if s.startswith("Z"):
        codigo = "".join(ch for ch in s[1:] if ch.isdigit())
        return f"Z{codigo}" if codigo else ""
    if s.isdigit():
        return f"Z{s}"
    return s


def _extraer_movimiento_monto(row: str) -> tuple[str, str]:
    """
    Extrae (movimiento, monto_raw) de una línea Control CIO.
    1) Intenta layout fijo (30-33 y 34-51).
    2) Fallback: busca primer patrón Z110..Z115 seguido de 18 dígitos.
    """
    mov = _normalizar_movimiento(row[29:33]) if len(row) >= 33 else ""
    monto_raw = row[33:51] if len(row) >= 51 else ""
    if mov in {"Z110", "Z111", "Z112", "Z113", "Z114", "Z115"}:
        return mov, monto_raw

    m = re.search(r"(Z11[0-5])\s*([0-9]{18})", row.upper())
    if m:
        return m.group(1), m.group(2)

    return "", ""


def _leer_control_cio_desde_txt(ruta: str) -> dict[str, float]:
    """
    Parsea TXT Control CIO.
    Soporta:
      - layout fijo por línea (mov 30-33, monto 34-51)
      - líneas corridas
      - archivos que traen '\n' escapado en vez de saltos reales
      - múltiples ocurrencias Z11x + monto en una misma línea
    """
    totales = {"Z110": 0.0, "Z111": 0.0, "Z112": 0.0, "Z113": 0.0, "Z114": 0.0, "Z115": 0.0}

    with open(ruta, "r", encoding="utf-8", errors="ignore") as f:
        content = f.read() or ""

    # Algunos orígenes guardan texto con '\n' literal (dos chars) en lugar de salto real.
    content = content.replace("\r\n", "\n").replace("\r", "\n").replace("\\n", "\n")

    for row in content.splitlines():
        row = str(row or "").strip()
        if not row:
            continue

        # Si hay varias operaciones concatenadas en la línea, tomar todas.
        matches = list(re.finditer(r"(Z11[0-5])\s*([0-9]{18})", row.upper()))
        if matches:
            for m in matches:
                mov = m.group(1)
                monto = _parse_importe_fijo(m.group(2))
                if mov in totales and monto != 0.0:
                    totales[mov] += monto
            continue

        # Fallback a extracción tradicional por posiciones fijas.
        mov, monto_raw = _extraer_movimiento_monto(row)
        if mov not in totales:
            continue

        monto = _parse_importe_fijo(monto_raw)
        if monto == 0.0:
            continue

        totales[mov] += monto

    return totales


def _parse_valor(v: str) -> float:
    s = str(v or "").strip()
    if not s:
        return 0.0
    s = s.replace("$", "").strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


_DIFF_ROW_WARN_MIN = 20.0
_DIFF_ROW_DANGER_MIN = 40.0
_DIFF_ROW_WARN_BG = "#FFF2CC"
_DIFF_ROW_DANGER_BG = "#FDE2E1"


def _clasificar_diferencia(valor: str) -> str | None:
    numero = abs(float(_parse_valor(valor)))
    if numero > _DIFF_ROW_DANGER_MIN:
        return "danger"
    if numero > _DIFF_ROW_WARN_MIN:
        return "warn"
    return None


def _diff_foreground(valor: str) -> str:
    return "#0F172A"


def _fmt_pesos(v: float) -> str:
    s = f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


def _estado_inicial_control_cio() -> dict:
    return {
        "version": 1,
        "current_semana": SEMANAS[0],
        "semanas": {
            s: [
                {
                    "concepto": fila,
                    "lunes": "",
                    "martes": "",
                    "miercoles": "",
                    "jueves": "",
                    "viernes": "",
                    "sabado": "",
                    "total": "$ 0,00",
                    "reporte_prescripto": "",
                    "diferencia_reporte": "",
                    "sfa_prescripto": "",
                    "diferencia": "",
                }
                for fila in FILAS_BASE
            ]
            for s in SEMANAS
        },
    }


def _normalizar_payload_control_cio(payload: dict | None) -> dict:
    base = _estado_inicial_control_cio()
    if not isinstance(payload, dict):
        return base

    semanas_in = payload.get("semanas", {})
    semanas_out = {}
    for sem in SEMANAS:
        rows_in = semanas_in.get(sem, []) if isinstance(semanas_in, dict) else []
        filas = []
        for i, concepto in enumerate(FILAS_BASE):
            row_in = rows_in[i] if isinstance(rows_in, list) and i < len(rows_in) and isinstance(rows_in[i], dict) else {}
            filas.append(
                {
                    "concepto": str(row_in.get("concepto", concepto) or concepto),
                    "lunes": str(row_in.get("lunes", "") or ""),
                    "martes": str(row_in.get("martes", "") or ""),
                    "miercoles": str(row_in.get("miercoles", "") or ""),
                    "jueves": str(row_in.get("jueves", "") or ""),
                    "viernes": str(row_in.get("viernes", "") or ""),
                    "sabado": str(row_in.get("sabado", "") or ""),
                    "total": str(row_in.get("total", "$ 0,00") or "$ 0,00"),
                    "reporte_prescripto": str(row_in.get("reporte_prescripto", "") or ""),
                    "diferencia_reporte": str(row_in.get("diferencia_reporte", "") or ""),
                    "sfa_prescripto": str(row_in.get("sfa_prescripto", "") or ""),
                    "diferencia": str(row_in.get("diferencia", "") or ""),
                }
            )
        semanas_out[sem] = filas

    current = str(payload.get("current_semana", SEMANAS[0]) or SEMANAS[0])
    if current not in SEMANAS:
        current = SEMANAS[0]

    return {"version": 1, "current_semana": current, "semanas": semanas_out}


def build_control_cio(fr_seccion: ttk.Frame, estado_var):
    cols = [
        "concepto",
        "lunes",
        "martes",
        "miercoles",
        "jueves",
        "viernes",
        "sabado",
        "total",
        "reporte_prescripto",
        "diferencia_reporte",
        "sfa_prescripto",
        "diferencia",
    ]
    col_labels = [
        "Concepto", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado",
        "Total", "Totales", "Diferencia", "Totales", "Diferencia",
    ]

    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(0, weight=0)
    fr_seccion.rowconfigure(1, weight=1)

    seed = _normalizar_payload_control_cio(getattr(app_state, "planilla_control_cio_data", None))
    app_state.planilla_control_cio_data = seed
    datos_por_semana = seed["semanas"]
    current_semana = seed.get("current_semana", SEMANAS[0])

    toolbar_style = str(getattr(app_state, "planilla_toolbar_combobox_style", "") or "").strip()
    try:
        toolbar_width = int(getattr(app_state, "planilla_toolbar_combobox_width", 30) or 30)
    except Exception:
        toolbar_width = 30

    top = ttk.Frame(fr_seccion, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=(2, 10), pady=(6, 10))
    top.columnconfigure(2, weight=1)

    ttk.Label(top, text="Semana:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")
    combo_semana_kwargs = {"state": "disabled", "values": _combo_values_semanas(), "width": toolbar_width}
    if toolbar_style:
        combo_semana_kwargs["style"] = toolbar_style
    combo_semana = ttk.Combobox(top, **combo_semana_kwargs)
    combo_semana.grid(row=0, column=1, sticky="w", padx=(8, 10))

    cont = ttk.Frame(fr_seccion)
    cont.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
    cont.columnconfigure(0, weight=1)
    cont.rowconfigure(3, weight=1)
    try:
        cont.grid_rowconfigure(1, minsize=30)
        cont.grid_rowconfigure(2, minsize=28)
    except Exception:
        pass

    COLOR_FONDO = _get_ttk_background(fr_seccion)

    header1 = tk.Canvas(cont, height=28, bg=COLOR_FONDO, highlightthickness=0, bd=0)
    header1.grid(row=1, column=0, sticky="ew")
    header2 = tk.Canvas(cont, height=26, bg=COLOR_FONDO, highlightthickness=0, bd=0)
    header2.grid(row=2, column=0, sticky="ew")

    def _ajustar_ancho_combo_semana(valores: list[str] | None = None, texto_actual: str = ""):
        try:
            combo_semana.configure(width=toolbar_width)
        except Exception:
            pass

    widths = [340, 132, 132, 132, 132, 132, 132, 150, 150, 150, 150, 150]
    total_width = sum(widths)

    COLOR_GRUPO_CIO = "#BFD7F6"
    COLOR_GRUPO_REPORTE = "#BFEBD3"
    COLOR_GRUPO_SFA = "#FFD7AA"
    COLOR_GRUPO_DIF = "#FCE9A7"

    COLOR_COL_CIO = _lighten(COLOR_GRUPO_CIO, 1.10)
    COLOR_COL_REPORTE = _lighten(COLOR_GRUPO_REPORTE, 1.10)
    COLOR_COL_SFA = _lighten(COLOR_GRUPO_SFA, 1.08)
    COLOR_COL_DIF = _lighten(COLOR_GRUPO_DIF, 1.06)

    BORDE_CIO = _darken(COLOR_COL_CIO, 0.78)
    BORDE_GRUPO_CIO = _darken(COLOR_GRUPO_CIO, 0.78)
    BORDE_REPORTE = _darken(COLOR_COL_REPORTE, 0.80)
    BORDE_GRUPO_REPORTE = _darken(COLOR_GRUPO_REPORTE, 0.78)
    BORDE_SFA = _darken(COLOR_COL_SFA, 0.80)
    BORDE_GRUPO_SFA = _darken(COLOR_GRUPO_SFA, 0.78)
    BORDE_DIF = _darken(COLOR_COL_DIF, 0.80)

    style = ttk.Style(fr_seccion)
    tree_style = "PlanillaControlCIO.Treeview"
    style.configure(tree_style, rowheight=27, font=("Segoe UI", 9), background="#FFFFFF", fieldbackground="#FFFFFF", foreground="#0F172A")
    style.map(tree_style, background=[("selected", "#DCEBFF")], foreground=[("selected", "#0F172A")])

    tree = ttk.Treeview(cont, columns=cols, show="tree", height=18, style=tree_style)
    tree.grid(row=3, column=0, sticky="nsew")
    tree.column("#0", width=0, stretch=False)
    for c, w in zip(cols, widths):
        tree.column(c, width=w, anchor=("w" if c == "concepto" else "e"), stretch=False)

    diff_labels: list[tk.Label] = []
    diff_render_job = None
    diff_render_after_job = None

    def _clear_diff_labels(*, destroy: bool = False):
        nonlocal diff_labels
        for label in diff_labels:
            try:
                if destroy:
                    label.destroy()
                else:
                    label.place_forget()
            except Exception:
                pass
        if destroy:
            diff_labels = []



    def _ensure_diff_label_pool(count: int):
        nonlocal diff_labels
        while len(diff_labels) < count:
            diff_labels.append(tk.Label(
                tree,
                anchor="e",
                padx=6,
                pady=0,
                borderwidth=0,
                highlightthickness=0,
                font=("Segoe UI", 9),
            ))
    def _row_background(iid: str) -> str:
        if iid in tree.selection():
            return "#DCEBFF"
        tags = tree.item(iid, "tags")
        if "odd" in tags:
            return "#F8FAFC"
        return "#FFFFFF"

    def _render_diff_labels():
        return
        _clear_diff_labels()
        visibles = tree.winfo_height()
        if visibles <= 0:
            return
        diff_cells: list[tuple[int, int, int, int, str, str, str]] = []
        for iid in tree.get_children():
            if not tree.bbox(iid, "#1"):
                continue
            valores = [str(v) for v in tree.item(iid, "values")]
            fondo = _row_background(iid)
            for col_idx in (9, 11):
                if col_idx >= len(valores):
                    continue
                valor = valores[col_idx].strip()
                if _clasificar_diferencia(valor) not in ("warn", "danger"):
                    continue
                bbox = tree.bbox(iid, f"#{col_idx + 1}")
                if not bbox:
                    continue
                x, y, width, height = bbox
                if y >= visibles or (y + height) <= 0 or width <= 2 or height <= 2:
                    continue
                diff_cells.append((x, y, width, height, valor, fondo, _diff_foreground(valor)))

        _ensure_diff_label_pool(len(diff_cells))
        for idx, (x, y, width, height, valor, fondo, foreground) in enumerate(diff_cells):
            label = diff_labels[idx]
            label.configure(text=valor, bg=fondo, fg=foreground)
            label.place(x=x + 1, y=y + 1, width=width - 2, height=height - 2)

    def _render_diff_labels_idle():
        nonlocal diff_render_job, diff_render_after_job
        diff_render_after_job = None
        diff_render_job = None
        _render_diff_labels()

    def _request_diff_render(_evt=None):
        return

    vs = ttk.Scrollbar(cont, orient="vertical", command=tree.yview)
    vs.grid(row=3, column=1, sticky="ns")
    hs = ttk.Scrollbar(cont, orient="horizontal", command=tree.xview)
    hs.grid(row=4, column=0, sticky="ew")
    tree.configure(yscrollcommand=vs.set)

    def _sync_headers(*_):
        x0, _x1 = tree.xview()
        header1.xview_moveto(x0)
        header2.xview_moveto(x0)

    def _xscroll(f, l):
        hs.set(f, l)
        _sync_headers()
        _request_diff_render()

    def _tree_xview(*args):
        tree.xview(*args)
        _sync_headers()
        _request_diff_render()

    def _tree_yview(*args):
        tree.yview(*args)
        _request_diff_render()

    def _on_tree_yscroll(f, l):
        vs.set(f, l)
        _request_diff_render()

    tree.configure(xscrollcommand=_xscroll, yscrollcommand=_on_tree_yscroll)
    hs.configure(command=_tree_xview)
    vs.configure(command=_tree_yview)
    bind_smooth_mousewheel(
        tree=tree,
        targets=(tree, cont, header1, header2),
        on_scroll=_request_diff_render,
    )
    tree.bind("<Configure>", _request_diff_render, add="+")
    tree.bind("<<TreeviewSelect>>", _request_diff_render, add="+")
    for event_name in (
        "<MouseWheel>",
        "<Shift-MouseWheel>",
        "<Button-4>",
        "<Button-5>",
        "<Shift-Button-4>",
        "<Shift-Button-5>",
        "<KeyRelease-Up>",
        "<KeyRelease-Down>",
        "<KeyRelease-Left>",
        "<KeyRelease-Right>",
        "<KeyRelease-Prior>",
        "<KeyRelease-Next>",
        "<KeyRelease-Home>",
        "<KeyRelease-End>",
    ):
        tree.bind(event_name, _request_diff_render, add="+")

    def _actualizar_label_rango_del_al(sem: str):
        try:
            sem_n = int(str(sem).replace("Semana", "").strip() or "1")
        except Exception:
            sem_n = 1

        hook = getattr(app_state, "planilla_actualizar_rango_del_al_seccion", None)
        if callable(hook):
            hook(sem_n)

    _week_sync_guard = {"active": False}

    def _publicar_semana_global_desde_combo(sem: str):
        sem_txt = _semana_interna(sem or SEMANAS[0])
        try:
            sem_n = int(str(sem_txt).replace("Semana", "").strip() or "1")
        except Exception:
            sem_n = 1

        rangos = getattr(app_state, "planilla_rangos_semana_global", {}) or {}
        rango = rangos.get(sem_n)
        if rango is None:
            rango = rangos.get(str(sem_n))

        desde = ""
        hasta = ""
        if isinstance(rango, dict):
            desde = str(rango.get("desde", "") or "").strip()
            hasta = str(rango.get("hasta", "") or "").strip()
        elif isinstance(rango, (list, tuple)) and len(rango) >= 2:
            desde = str(rango[0] or "").strip()
            hasta = str(rango[1] or "").strip()

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
                publicar(juego_actual, sem_n, desde, hasta)
            except Exception:
                pass

    def _persistir_estado():
        app_state.planilla_control_cio_data = _normalizar_payload_control_cio({
            "version": 1,
            "current_semana": current_semana,
            "semanas": datos_por_semana,
        })

    def _guardar_semana(sem: str):
        rows = []
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            rows.append({k: str(v) for k, v in zip(cols, vals)})
        datos_por_semana[sem] = rows
        _persistir_estado()

    def _cargar_semana(sem: str):
        _actualizar_label_rango_del_al(sem)
        for iid in tree.get_children():
            tree.delete(iid)
        for row in datos_por_semana.get(sem, []):
            tree.insert("", "end", values=[row.get(c, "") for c in cols])
        _aplicar_zebra()
        _request_diff_render()

    def _aplicar_zebra():
        tree.tag_configure("even", background="#FFFFFF", foreground="#0F172A")
        tree.tag_configure("odd", background="#F8FAFC", foreground="#0F172A")
        tree.tag_configure("diff_warn", background=_DIFF_ROW_WARN_BG, foreground="#0F172A")
        tree.tag_configure("diff_danger", background=_DIFF_ROW_DANGER_BG, foreground="#0F172A")
        for idx, iid in enumerate(tree.get_children()):
            vals = [str(v) for v in tree.item(iid, "values")]
            base_tag = "even" if idx % 2 == 0 else "odd"
            diffs = []
            if len(vals) > 9:
                diffs.append(_clasificar_diferencia(vals[9]))
            if len(vals) > 11:
                diffs.append(_clasificar_diferencia(vals[11]))
            if "danger" in diffs:
                tree.item(iid, tags=("diff_danger",))
            elif "warn" in diffs:
                tree.item(iid, tags=("diff_warn",))
            else:
                tree.item(iid, tags=(base_tag,))

    def _recalc_row(iid: str):
        vals = list(tree.item(iid, "values"))
        total = sum(_parse_valor(vals[i]) for i in range(1, 7))
        vals[7] = _fmt_pesos(total)
        control_reporte = _parse_valor(vals[8])
        vals[9] = _fmt_pesos(control_reporte - total) if str(vals[8]).strip() else ""
        control_sfa = _parse_valor(vals[10])
        vals[11] = _fmt_pesos(control_sfa - total) if str(vals[10]).strip() else ""
        tree.item(iid, values=vals)
        _aplicar_zebra()
        _request_diff_render()

    def _buscar_iid_por_concepto(concepto: str):
        c = str(concepto).strip().lower()
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            if vals and str(vals[0]).strip().lower() == c:
                return iid
        return None

    def _set_valor(iid: str, col_idx: int, valor: float | None):
        vals = list(tree.item(iid, "values"))
        vals[col_idx] = "" if valor is None else _fmt_pesos(valor)
        tree.item(iid, values=vals)
        _recalc_row(iid)

    def _cargar_control_reporte_desde_reporte():
        nonlocal current_semana
        semana_objetivo = current_semana

        col_rep_idx = cols.index("reporte_prescripto")
        col_diff_rep_idx = cols.index("diferencia_reporte")

        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            vals[col_rep_idx] = ""
            vals[col_diff_rep_idx] = ""
            tree.item(iid, values=vals)
            _recalc_row(iid)

        cash_in = float(getattr(app_state, "reporte_cash_in_importe", 0.0) or 0.0)
        cash_out = float(getattr(app_state, "reporte_cash_out_importe", 0.0) or 0.0)

        iid_in = _buscar_iid_por_concepto("CASH-IN")
        iid_out = _buscar_iid_por_concepto("CASH-OUT")

        if iid_in is not None:
            _set_valor(iid_in, col_rep_idx, cash_in if cash_in != 0.0 else None)
        if iid_out is not None:
            _set_valor(iid_out, col_rep_idx, cash_out if cash_out != 0.0 else None)

        _guardar_semana(semana_objetivo)
        if estado_var is not None:
            estado_var.set(f"{semana_objetivo}: actualizado Control Reporte (CASH IN / CASH OUT desde Reporte Facturación).")

    def _importar_resumen_dia(nombre_dia: str):
        nonlocal current_semana
        ruta = filedialog.askopenfilename(
            title=f"Importar CONTROL CIO ({nombre_dia.capitalize()})",
            filetypes=[("JSON o TXT", "*.json;*.JSON;*.txt;*.TXT"), ("Todos", "*.*")],
        )
        if not ruta:
            return
        ext = Path(ruta).suffix.lower()
        z110 = z111 = z112 = z113 = z114 = z115 = 0.0

        if ext == ".txt":
            try:
                totales = _leer_control_cio_desde_txt(ruta)
            except Exception as e:
                messagebox.showerror("Importación", f"No pude leer el TXT de Control CIO:\n{e}")
                return

            z110 = float(totales.get("Z110", 0.0) or 0.0)
            z111 = float(totales.get("Z111", 0.0) or 0.0)
            z112 = float(totales.get("Z112", 0.0) or 0.0)
            z113 = float(totales.get("Z113", 0.0) or 0.0)
            z114 = float(totales.get("Z114", 0.0) or 0.0)
            z115 = float(totales.get("Z115", 0.0) or 0.0)
        else:
            try:
                resumen = leer_json_sfa(ruta)
            except Exception as e:
                messagebox.showerror("Importación", f"No pude leer el JSON/TXT SFA:\n{e}")
                return

            for (_juego, _sorteo, concepto), importe in (resumen or {}).items():
                c = str(concepto).strip().upper().replace(" ", "")
                imp = float(importe or 0.0)
                if c in {"Z110", "110"}:
                    z110 += imp
                elif c in {"Z111", "111"}:
                    z111 += imp
                elif c in {"Z112", "112"}:
                    z112 += imp
                elif c in {"Z113", "113"}:
                    z113 += imp
                elif c in {"Z114", "114"}:
                    z114 += imp
                elif c in {"Z115", "115"}:
                    z115 += imp

        cash_in = z110 + z111
        cash_out = (z112 * 0.9) + (z113 * 0.9) + z114 + z115

        col_idx = cols.index(nombre_dia)
        iid_in = _buscar_iid_por_concepto("CASH-IN")
        iid_out = _buscar_iid_por_concepto("CASH-OUT")
        if iid_in is not None:
            _set_valor(iid_in, col_idx, cash_in)
        if iid_out is not None:
            _set_valor(iid_out, col_idx, cash_out)

        _guardar_semana(current_semana)
        if estado_var is not None:
            estado_var.set(f"{current_semana}: importado {nombre_dia.capitalize()} para Control CIO.")

    _day_buttons: dict[str, tk.Button] = {}

    def _draw_headers():
        header1.delete("all")
        header2.delete("all")

        for b in list(_day_buttons.values()):
            try:
                b.destroy()
            except Exception:
                pass
        _day_buttons.clear()

        x = 0
        header1.create_rectangle(x, 0, x + widths[0], 28, fill=COLOR_FONDO, outline=COLOR_FONDO, width=0)
        x += widths[0]
        w_control = sum(widths[1:8])
        header1.create_rectangle(x, 0, x + w_control, 28, fill=COLOR_GRUPO_CIO, outline=BORDE_GRUPO_CIO)
        header1.create_text(x + w_control / 2, 14, text="CONTROL CIO", font=("Segoe UI Semibold", 9))
        x += w_control
        header1.create_rectangle(x, 0, x + widths[8], 28, fill=COLOR_GRUPO_REPORTE, outline=BORDE_GRUPO_REPORTE)
        header1.create_text(x + widths[8] / 2, 14, text="Control Reporte", font=("Segoe UI Semibold", 9))
        x += widths[8]
        header1.create_rectangle(x, 0, x + widths[9], 28, fill=COLOR_GRUPO_DIF, outline=BORDE_DIF)
        header1.create_text(x + widths[9] / 2, 14, text="Diferencia", font=("Segoe UI Semibold", 9))
        x += widths[9]
        header1.create_rectangle(x, 0, x + widths[10], 28, fill=COLOR_GRUPO_SFA, outline=BORDE_GRUPO_SFA)
        header1.create_text(x + widths[10] / 2, 14, text="Control SFA", font=("Segoe UI Semibold", 9))
        x += widths[10]
        header1.create_rectangle(x, 0, x + widths[11], 28, fill=COLOR_GRUPO_DIF, outline=BORDE_DIF)
        header1.create_text(x + widths[11] / 2, 14, text="Diferencia", font=("Segoe UI Semibold", 9))

        x = 0
        for i, (label, w) in enumerate(zip(col_labels, widths)):
            if i == 0:
                fill, outline = COLOR_COL_CIO, BORDE_CIO
            elif 1 <= i <= 7:
                fill, outline = COLOR_COL_CIO, BORDE_CIO
            elif i == 8:
                fill, outline = COLOR_COL_REPORTE, BORDE_REPORTE
            elif i == 10:
                fill, outline = COLOR_COL_SFA, BORDE_SFA
            else:
                fill, outline = COLOR_COL_DIF, BORDE_DIF

            header2.create_rectangle(x, 0, x + w, 26, fill=fill, outline=outline, width=1)
            if 1 <= i <= 6:
                dia = DIAS[i - 1]
                btn = tk.Button(
                    header2,
                    text=label,
                    command=lambda d=dia: _importar_resumen_dia(d),
                    bg=fill,
                    activebackground=fill,
                    fg="black",
                    activeforeground="black",
                    relief="flat",
                    bd=0,
                    highlightthickness=0,
                    font=("Segoe UI Semibold", 9),
                    cursor="hand2",
                )
                _aplicar_hover_boton(btn, fill)
                _day_buttons[dia] = btn
                header2.create_window(x + w / 2, 13, window=btn, width=max(20, w - 2), height=24)
            else:
                header2.create_text(x + w / 2, 13, text=label, font=("Segoe UI Semibold", 9))
            x += w

        header1.configure(scrollregion=(0, 0, total_width, 28))
        header2.configure(scrollregion=(0, 0, total_width, 26))

    def _on_week_change(_evt=None):
        nonlocal current_semana
        nueva = _semana_interna(combo_semana.get() or SEMANAS[0])
        _guardar_semana(current_semana)
        current_semana = nueva
        _cargar_semana(current_semana)
        _cargar_control_reporte_desde_reporte()
        _persistir_estado()
        if not _week_sync_guard["active"]:
            _publicar_semana_global_desde_combo(current_semana)

    editable_col_indices = {1, 2, 3, 4, 5, 6, 8, 10}
    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)
    edit_entry = None

    def _close_editor(save=True):
        nonlocal edit_entry
        if edit_entry is None:
            return
        if save:
            try:
                iid, col = edit_entry._cell
                col_idx = int(col.replace("#", "")) - 1
                vals = list(tree.item(iid, "values"))
                nuevo_valor = str(edit_entry.get()).strip()
                if vals[col_idx] != nuevo_valor:
                    push_undo_rows(undo_state, tree, [(iid, list(vals))], meta={"semana": current_semana})
                    vals[col_idx] = nuevo_valor
                    tree.item(iid, values=vals)
                    _recalc_row(iid)
                    _request_diff_render()
            except Exception:
                pass
            _guardar_semana(current_semana)
        edit_entry.destroy()
        edit_entry = None

    def _copiar_celdas(_evt=None):
        _close_editor(save=True)
        iid, col_idx = get_anchor_cell(tree, clipboard_state, default_col=1)
        if not iid or col_idx not in editable_col_indices:
            return "break"
        seleccion = ordered_selected_rows(tree)
        row_ids = seleccion if len(seleccion) > 1 and iid in seleccion else [iid]
        matrix = []
        for row_iid in row_ids:
            vals = [str(v) for v in tree.item(row_iid, "values")]
            while len(vals) <= col_idx:
                vals.append("")
            matrix.append([vals[col_idx]])
        set_clipboard_matrix(tree, matrix)
        return "break"

    def _pegar_celdas(_evt=None):
        _close_editor(save=True)
        matrix = get_clipboard_matrix(tree)
        if not matrix:
            return "break"
        anchor_iid, anchor_col = get_anchor_cell(tree, clipboard_state, default_col=1)
        row_ids = list(tree.get_children())
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
            row_changed = False
            for c_off, cell_raw in enumerate(row_data):
                col_idx = anchor_col + c_off
                if col_idx not in editable_col_indices:
                    continue
                vals[col_idx] = str(cell_raw or "").strip()
                row_changed = True
            if row_changed and vals != before_vals:
                undo_rows.append((iid, before_vals))
                tree.item(iid, values=vals)
                _recalc_row(iid)
                hubo_cambios = True
        if hubo_cambios:
            push_undo_rows(undo_state, tree, undo_rows, meta={"semana": current_semana})
            _request_diff_render()
            _guardar_semana(current_semana)
        return "break"

    def _deshacer_celdas(_evt=None):
        nonlocal edit_entry, current_semana
        if edit_entry is not None:
            _close_editor(save=False)
        snapshot = pop_undo_snapshot(undo_state)
        if not snapshot:
            return "break"
        meta = snapshot.get("meta", {}) if isinstance(snapshot, dict) else {}
        semana_undo = str(meta.get("semana", current_semana) or current_semana)
        if semana_undo in SEMANAS and semana_undo != current_semana:
            _guardar_semana(current_semana)
            current_semana = semana_undo
            combo_semana.set(_semana_visible(current_semana))
            _cargar_semana(current_semana)
            _cargar_control_reporte_desde_reporte()
        restore_undo_snapshot(tree, snapshot)
        for iid, _vals in (snapshot.get("rows", []) or []):
            _recalc_row(str(iid))
        _request_diff_render()
        _guardar_semana(current_semana)
        return "break"

    def _on_double_click(evt):
        nonlocal edit_entry
        _close_editor(save=True)
        if tree.identify("region", evt.x, evt.y) != "cell":
            return
        col = tree.identify_column(evt.x)
        iid = tree.identify_row(evt.y)
        if not iid:
            return
        col_idx = int(col.replace("#", "")) - 1
        if col_idx not in editable_col_indices:
            return
        clipboard_state["cell"] = (iid, col_idx)
        x, y, w, h = tree.bbox(iid, col)
        if w <= 0 or h <= 0:
            return
        edit_entry = ttk.Entry(tree)
        edit_entry.place(x=x, y=y, width=w, height=h)
        edit_entry.insert(0, tree.item(iid, "values")[col_idx])
        edit_entry.focus_set()
        edit_entry._cell = (iid, col)
        edit_entry.bind("<Return>", lambda _e: _close_editor(save=True))
        edit_entry.bind("<Escape>", lambda _e: _close_editor(save=False))
        edit_entry.bind("<FocusOut>", lambda _e: _close_editor(save=True))

    def _reset_control_cio():
        nonlocal current_semana
        nuevo = _estado_inicial_control_cio()
        datos_por_semana.clear()
        datos_por_semana.update(nuevo["semanas"])
        current_semana = SEMANAS[0]
        combo_semana.set(_semana_visible(current_semana))
        _cargar_semana(current_semana)
        _persistir_estado()

    def _cargar_snapshot_control_cio(payload: dict):
        nonlocal current_semana
        normalizado = _normalizar_payload_control_cio(payload)
        app_state.planilla_control_cio_data = normalizado
        datos_por_semana.clear()
        datos_por_semana.update(normalizado["semanas"])
        current_semana = normalizado.get("current_semana", SEMANAS[0])
        combo_semana.set(_semana_visible(current_semana))
        _cargar_semana(current_semana)

    if not hasattr(app_state, "planilla_control_cio_reset_hooks"):
        app_state.planilla_control_cio_reset_hooks = {}
    app_state.planilla_control_cio_reset_hooks["control_cio"] = _reset_control_cio

    if not hasattr(app_state, "planilla_control_cio_load_hooks"):
        app_state.planilla_control_cio_load_hooks = {}
    app_state.planilla_control_cio_load_hooks["control_cio"] = _cargar_snapshot_control_cio

    if not hasattr(app_state, "planilla_control_cio_refresh_hooks"):
        app_state.planilla_control_cio_refresh_hooks = {}
    app_state.planilla_control_cio_refresh_hooks["control_cio"] = _cargar_control_reporte_desde_reporte

    if not hasattr(app_state, "planilla_bundle_snapshot_hooks"):
        app_state.planilla_bundle_snapshot_hooks = {}

    def _snapshot_control_cio() -> dict:
        _guardar_semana(current_semana)
        return _normalizar_payload_control_cio(app_state.planilla_control_cio_data)

    app_state.planilla_bundle_snapshot_hooks["control_cio"] = _snapshot_control_cio

    _draw_headers()
    combo_semana.bind("<<ComboboxSelected>>", _on_week_change)
    try:
        _vals_ini = _combo_values_semanas()
        combo_semana.configure(values=_vals_ini)
        if _vals_ini:
            combo_semana.configure(state="readonly")
            combo_semana.set(_semana_visible(current_semana))
        else:
            combo_semana.configure(state="disabled")
            combo_semana.set("")
    except Exception:
        combo_semana.set("")
    _cargar_semana(current_semana)
    _cargar_control_reporte_desde_reporte()
    tree.bind("<Double-1>", _on_double_click)
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")
    _sync_headers()

    if not hasattr(app_state, "planilla_semana_filtro_hooks"):
        app_state.planilla_semana_filtro_hooks = {}

    def _aplicar_filtro_area_recaudacion(payload: dict):
        values = _combo_values_semanas()
        try:
            combo_semana.configure(values=values)
            _ajustar_ancho_combo_semana(values, str(combo_semana.get() or "").strip())
        except Exception:
            pass

        sem_n = 0
        try:
            sem_n = int((payload or {}).get("semana", 0) or 0)
        except Exception:
            sem_n = 0

        if sem_n < 1 or not values:
            try:
                combo_semana.set("")
                combo_semana.configure(state="disabled")
                _ajustar_ancho_combo_semana([], "")
            except Exception:
                pass
            return

        sem_n = max(1, min(5, sem_n))
        visible = _semana_visible(f"Semana {sem_n}")
        try:
            combo_semana.configure(state="readonly")
        except Exception:
            pass

        _ajustar_ancho_combo_semana(values, visible)

        if combo_semana.get() == visible and current_semana == f"Semana {sem_n}":
            return

        combo_semana.set(visible)
        _week_sync_guard["active"] = True
        try:
            _on_week_change()
        finally:
            _week_sync_guard["active"] = False

    app_state.planilla_semana_filtro_hooks["control_cio"] = _aplicar_filtro_area_recaudacion

    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    _refresh_visual_job = {"id": None}

    def _refresh_visual_control_cio():
        try:
            _vals_refresh = _combo_values_semanas()
            combo_semana.configure(values=_vals_refresh)
            payload = getattr(app_state, "planilla_semana_filtro_actual", {})
            sem_n = int((payload or {}).get("semana", 0) or 0) if isinstance(payload, dict) else 0
            if sem_n < 1 or not _vals_refresh:
                combo_semana.set("")
                combo_semana.configure(state="disabled")
            else:
                combo_semana.configure(state="readonly")
        except Exception:
            pass

        if _refresh_visual_job["id"] is not None:
            return

        def _run_visual():
            _refresh_visual_job["id"] = None
            try:
                _sync_headers()
            except Exception:
                pass

        try:
            _refresh_visual_job["id"] = cont.after_idle(_run_visual)
        except Exception:
            _run_visual()

    app_state.planilla_visual_refresh_hooks["Control CIO"] = _refresh_visual_control_cio

    payload_inicial = getattr(app_state, "planilla_semana_filtro_actual", {})
    if isinstance(payload_inicial, dict):
        _aplicar_filtro_area_recaudacion(dict(payload_inicial))

    if estado_var is not None:
        estado_var.set("Control CIO listo.")  
