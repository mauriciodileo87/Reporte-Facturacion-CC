# tabs/planilla_anticipos_topes.py
from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from json_sfa import leer_json_sfa
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


FILAS_BASE = [
    "Anticipos",
    "Anti Rec QNL",
    "Anti Rec PCD",
    "Anti Rec Loto",
    "Anti Rec Q6",
    "Anti Rec BRN",
    "Anti Rec TMB",
    "Anti Rec LT5",
    "Anti Rec QNY",
    "Totales Anti Rec",
]

FILA_TOTALES_ANTI_REC = "Totales Anti Rec"

SEMANAS = ["Semana 1", "Semana 2", "Semana 3", "Semana 4", "Semana 5"]
DIAS = ["lunes", "martes", "miercoles", "jueves", "viernes", "sabado"]

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



# Mapeo oficial (Z062 por código -> fila)
CODIGO_A_FILA_Z062 = {
    "80": "Anti Rec QNL",
    "79": "Anti Rec QNY",
    "74": "Anti Rec TMB",
    "82": "Anti Rec PCD",
    "69": "Anti Rec Q6",
    "13": "Anti Rec BRN",
    "9": "Anti Rec Loto",
    "5": "Anti Rec LT5",
}


JUEGO_PLANILLA_A_FILA_ANTI_REC = {
    "Quiniela": "Anti Rec QNL",
    "Poceada": "Anti Rec PCD",
    "Loto": "Anti Rec Loto",
    "Loto 5": "Anti Rec LT5",
    "Quini 6": "Anti Rec Q6",
    "Brinco": "Anti Rec BRN",
    "Tombolina": "Anti Rec TMB",
    "Quiniela Ya": "Anti Rec QNY",
}


def _parse_valor(v: str) -> float:
    s = str(v or "").strip()
    if not s:
        return 0.0
    s = s.replace("$", "").strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def _fmt_pesos(v: float) -> str:
    s = f"{float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"$ {s}"


_DIFF_COLOR_ALERT_MIN = 20.0
_DIFF_FOREGROUND_ALERT = "#B91C1C"


def _clasificar_diferencia(valor: str) -> str | None:
    numero = abs(float(_parse_valor(valor)))
    if numero > _DIFF_COLOR_ALERT_MIN:
        return "alert"
    return None


def _diff_foreground(valor: str) -> str:
    clasificacion = _clasificar_diferencia(valor)
    if clasificacion == "alert":
        return _DIFF_FOREGROUND_ALERT
    return "#0F172A"


def _darken(hex_color: str, factor: float = 0.82) -> str:
    h = hex_color.lstrip("#")
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    r = max(0, min(255, int(r * factor)))
    g = max(0, min(255, int(g * factor)))
    b = max(0, min(255, int(b * factor)))
    return f"#{r:02X}{g:02X}{b:02X}"


def _lighten(hex_color: str, factor: float = 1.08) -> str:
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


def _aplicar_zebra_anticipos(tree: ttk.Treeview):
    tree.tag_configure("even", background="#FFFFFF")
    tree.tag_configure("odd", background="#F8FAFC")
    for idx, iid in enumerate(tree.get_children()):
        vals = tree.item(iid, "values")
        base_tag = "even" if idx % 2 == 0 else "odd"
        tree.item(iid, tags=(base_tag,))


def _estado_inicial_anticipos_topes() -> dict:
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


def _normalizar_payload_anticipos_topes(payload: dict | None) -> dict:
    base = _estado_inicial_anticipos_topes()
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

def build_anticipos_topes(fr_seccion: ttk.Frame, estado_var):
    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(0, weight=0)
    fr_seccion.rowconfigure(1, weight=1)

    # Fondo real del programa (para que NO se note que arriba de Concepto/Diferencia no hay nada)
    COLOR_FONDO = _get_ttk_background(fr_seccion)

    style = ttk.Style(fr_seccion)
    tree_style = "PlanillaAnticipos.Treeview"
    style.configure(
        tree_style,
        rowheight=27,
        font=("Segoe UI", 9),
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        bordercolor="#D5DEE8",
        lightcolor="#D5DEE8",
        darkcolor="#D5DEE8",
    )
    style.map(
        tree_style,
        background=[("selected", "#DCEBFF")],
        foreground=[("selected", "#0F172A")],
    )

    # Paleta alineada a Área Recaudación (misma tonalidad)
    COLOR_GRUPO_ANTICIPOS = "#BFD7F6"  # equivalente a Tickets
    COLOR_GRUPO_REPORTE = "#BFEBD3"    # equivalente a Reporte
    COLOR_GRUPO_SFA = "#FFD7AA"        # equivalente a SFA
    COLOR_GRUPO_DIF = "#FCE9A7"        # equivalente a Diferencias

    COLOR_COL_ANTICIPOS = _lighten(COLOR_GRUPO_ANTICIPOS, 1.10)
    COLOR_COL_REPORTE = _lighten(COLOR_GRUPO_REPORTE, 1.10)
    COLOR_COL_SFA = _lighten(COLOR_GRUPO_SFA, 1.08)
    COLOR_COL_DIF = _lighten(COLOR_GRUPO_DIF, 1.06)

    BORDE_ANT = _darken(COLOR_COL_ANTICIPOS, 0.78)
    BORDE_GRUPO_ANT = _darken(COLOR_GRUPO_ANTICIPOS, 0.78)
    BORDE_REPORTE = _darken(COLOR_COL_REPORTE, 0.80)
    BORDE_GRUPO_REPORTE = _darken(COLOR_GRUPO_REPORTE, 0.78)
    BORDE_SFA = _darken(COLOR_COL_SFA, 0.80)
    BORDE_GRUPO_SFA = _darken(COLOR_GRUPO_SFA, 0.78)
    BORDE_DIF = _darken(COLOR_COL_DIF, 0.80)

    # Estado por semana (persistido en app_state para Guardar/Cargar bundle completo)
    seed = _normalizar_payload_anticipos_topes(getattr(app_state, "planilla_anticipos_topes_data", None))
    app_state.planilla_anticipos_topes_data = seed
    datos_por_semana: dict[str, list[dict[str, str]]] = seed["semanas"]

    # -------------------------
    # Contenedor
    # -------------------------
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
        "Concepto",
        "Lunes",
        "Martes",
        "Miércoles",
        "Jueves",
        "Viernes",
        "Sábado",
        "Total",
        "Totales",
        "",
        "Totales",
        "",
    ]

    widths = [340, 132, 132, 132, 132, 132, 132, 150, 150, 150, 150, 150]
    total_width = sum(widths)

    # -------------------------
    # Headers (Canvas)
    # -------------------------
    header_h1 = 28
    header1 = tk.Canvas(cont, height=header_h1, bg=COLOR_FONDO, highlightthickness=0, bd=0)
    header1.grid(row=0, column=0, sticky="ew")

    header_h2 = 26
    header2 = tk.Canvas(cont, height=header_h2, bg=COLOR_FONDO, highlightthickness=0, bd=0)
    header2.grid(row=1, column=0, sticky="ew")

    def _ajustar_ancho_combo_semana(valores: list[str] | None = None, texto_actual: str = ""):
        try:
            combo_semana.configure(width=toolbar_width)
        except Exception:
            pass

    # -------------------------
    # Tree
    # -------------------------
    tree = ttk.Treeview(cont, columns=cols, show="tree", height=18, style=tree_style)
    tree.grid(row=3, column=0, sticky="nsew")
    tree.column("#0", width=0, stretch=False)
    tree.heading("#0", text="")

    for c, w in zip(cols, widths):
        anchor = "w" if c == "concepto" else "e"
        tree.column(c, width=w, anchor=anchor, stretch=False)

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
                if _clasificar_diferencia(valor) != "alert":
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
    tree.configure(yscrollcommand=vs.set)

    hs = ttk.Scrollbar(cont, orient="horizontal", command=tree.xview)
    hs.grid(row=4, column=0, sticky="ew")

    def _sync_headers(*_):
        x0, _ = tree.xview()
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

    # -------------------------
    # Encabezados
    # -------------------------
    _day_buttons: dict[str, tk.Button] = {}
    _prescripto_button: tk.Button | None = None

    def _draw_group_headers():
        header1.delete("all")

        # “No cajón” arriba de Concepto + filtro Semana
        x0 = 0
        w0 = widths[0]
        header1.create_rectangle(x0, 0, x0 + w0, header_h1, fill=COLOR_FONDO, outline=COLOR_FONDO, width=0)

        x = w0

        # Anticipos y Topes
        w_ant = sum(widths[1:8])
        header1.create_rectangle(x, 0, x + w_ant, header_h1, fill=COLOR_GRUPO_ANTICIPOS, outline=BORDE_GRUPO_ANT, width=1)
        header1.create_text(x + w_ant / 2, header_h1 / 2, text="Anticipos y Topes", font=("Segoe UI Semibold", 9))
        x += w_ant

        # Control reporte
        w_rep = widths[8]
        header1.create_rectangle(x, 0, x + w_rep, header_h1, fill=COLOR_GRUPO_REPORTE, outline=BORDE_GRUPO_REPORTE, width=1)
        header1.create_text(x + w_rep / 2, header_h1 / 2, text="Control Reporte", font=("Segoe UI Semibold", 9))
        x += w_rep

        # Diferencia (intermedia)
        w_diff_rep = widths[9]
        header1.create_rectangle(x, 0, x + w_diff_rep, header_h1, fill=COLOR_GRUPO_DIF, outline=BORDE_DIF, width=1)
        header1.create_text(x + w_diff_rep / 2, header_h1 / 2, text="Diferencia", font=("Segoe UI Semibold", 9))
        x += w_diff_rep

        # Control SFA
        w_sfa = widths[10]
        header1.create_rectangle(x, 0, x + w_sfa, header_h1, fill=COLOR_GRUPO_SFA, outline=BORDE_GRUPO_SFA, width=1)
        header1.create_text(x + w_sfa / 2, header_h1 / 2, text="Control SFA", font=("Segoe UI Semibold", 9))
        x += w_sfa

        # Diferencia final
        w_dif = widths[11]
        header1.create_rectangle(x, 0, x + w_dif, header_h1, fill=COLOR_GRUPO_DIF, outline=BORDE_DIF, width=1)
        header1.create_text(x + w_dif / 2, header_h1 / 2, text="Diferencia", font=("Segoe UI Semibold", 9))

        header1.configure(scrollregion=(0, 0, total_width, header_h1))

    def _draw_column_headers():
        nonlocal _prescripto_button

        header2.delete("all")

        for b in list(_day_buttons.values()):
            try:
                b.destroy()
            except Exception:
                pass
        _day_buttons.clear()

        if _prescripto_button is not None:
            try:
                _prescripto_button.destroy()
            except Exception:
                pass
            _prescripto_button = None

        x = 0
        for i, (label, wpx) in enumerate(zip(col_labels, widths)):
            if i == 0:
                fill, outline = COLOR_COL_ANTICIPOS, BORDE_ANT
            elif 1 <= i <= 7:
                fill, outline = COLOR_COL_ANTICIPOS, BORDE_ANT
            elif i == 8:
                fill, outline = COLOR_COL_REPORTE, BORDE_REPORTE
            elif i == 10:
                fill, outline = COLOR_COL_SFA, BORDE_SFA
            else:
                fill, outline = COLOR_COL_DIF, BORDE_DIF

            header2.create_rectangle(x, 0, x + wpx, header_h2, fill=fill, outline=outline, width=1)

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
                header2.create_window(x + wpx / 2, header_h2 / 2, window=btn, width=wpx - 2, height=header_h2 - 2)

            elif i == 8:
                header2.create_text(x + wpx / 2, header_h2 / 2, text=label, font=("Segoe UI Semibold", 9))

            elif i == 10:
                _prescripto_button = tk.Button(
                    header2,
                    text=label,
                    command=_importar_prescripto_sfa,
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
                _aplicar_hover_boton(_prescripto_button, fill)
                header2.create_window(x + wpx / 2, header_h2 / 2, window=_prescripto_button, width=wpx - 2, height=header_h2 - 2)

            else:
                header2.create_text(x + wpx / 2, header_h2 / 2, text=label, font=("Segoe UI Semibold", 9))

            x += wpx

        header2.configure(scrollregion=(0, 0, total_width, header_h2))

    # -------------------------
    # Semana
    # -------------------------
    current_semana = seed.get("current_semana", SEMANAS[0])

    def _persistir_estado():
        app_state.planilla_anticipos_topes_data = _normalizar_payload_anticipos_topes(
            {
                "version": 1,
                "current_semana": current_semana,
                "semanas": datos_por_semana,
            }
        )

    def _sync_totales_txt_anticipos_topes(sem: str):
        guardar = getattr(app_state, "guardar_totales_importados", None)
        obtener = getattr(app_state, "obtener_totales_importados_semana", None)
        if not callable(guardar):
            return
        sem_txt = str(sem or "").strip() or SEMANAS[0]
        payload = obtener(sem_txt) if callable(obtener) else {}
        valores_txt = {}
        if isinstance(payload, dict):
            for etiqueta in ("Total ventas", "Total comisiones", "Total premios", "Total prescripcion", "Total facuni", "Total comision agencia amiga"):
                fila = payload.get(etiqueta, {})
                if isinstance(fila, dict):
                    valores_txt[etiqueta] = float(fila.get("txt", 0.0) or 0.0)
        total_anticipos = 0.0
        total_topes = 0.0
        for row in datos_por_semana.get(sem_txt, []):
            if not isinstance(row, dict):
                continue
            concepto = str(row.get("concepto", "") or "").strip().lower()
            try:
                total = _parse_valor(row.get("total", ""))
            except Exception:
                total = 0.0
            if concepto == "anticipos":
                total_anticipos = total
            elif concepto == FILA_TOTALES_ANTI_REC.lower():
                total_topes = total
        valores_txt["Total anticipos"] = float(total_anticipos or 0.0)
        valores_txt["Total topes"] = float(total_topes or 0.0)
        guardar(sem_txt, "txt", valores_txt)

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

        publicar = getattr(app_state, "publicar_filtro_area_recaudacion", None)
        if callable(publicar):
            try:
                publicar("", sem_n, desde, hasta)
            except Exception:
                pass

    def _actualizar_label_rango_del_al(sem: str):
        return

    def _guardar_semana(sem: str):
        rows = []
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            rows.append({k: str(v) for k, v in zip(cols, vals)})
        datos_por_semana[sem] = rows
        _persistir_estado()
        _sync_totales_txt_anticipos_topes(sem)

    def _cargar_semana(sem: str):
        _actualizar_label_rango_del_al(sem)
        for iid in tree.get_children():
            tree.delete(iid)
        for row in datos_por_semana.get(sem, []):
            tree.insert("", "end", values=[row.get(c, "") for c in cols])
        _actualizar_fila_totales_anti_rec()
        _aplicar_zebra_anticipos(tree)
        _request_diff_render()

    def _on_week_change(_evt=None):
        nonlocal current_semana
        nueva = _semana_interna(combo_semana.get() or SEMANAS[0])
        _guardar_semana(current_semana)
        _cargar_semana(nueva)
        current_semana = nueva
        _persistir_estado()
        if not _week_sync_guard["active"]:
            _publicar_semana_global_desde_combo(current_semana)
        if estado_var is not None:
            estado_var.set(f"Semana activa: {current_semana}")

    # -------------------------
    # Recalculo
    # -------------------------
    def _es_fila_totales_anti_rec(valores: list[str]) -> bool:
        return bool(valores) and str(valores[0]).strip().lower() == FILA_TOTALES_ANTI_REC.lower()

    def _actualizar_fila_totales_anti_rec():
        iid_totales = _buscar_iid_por_concepto(FILA_TOTALES_ANTI_REC)
        if iid_totales is None:
            return

        suma_total = 0.0
        suma_reporte = 0.0
        suma_dif_reporte = 0.0
        suma_sfa = 0.0
        suma_dif_sfa = 0.0

        for iid in tree.get_children():
            if iid == iid_totales:
                continue
            vals = list(tree.item(iid, "values"))
            concepto = str(vals[0]).strip().lower() if vals else ""
            if not concepto.startswith("anti rec"):
                continue
            suma_total += _parse_valor(vals[7])
            suma_reporte += _parse_valor(vals[8])
            suma_dif_reporte += _parse_valor(vals[9])
            suma_sfa += _parse_valor(vals[10])
            suma_dif_sfa += _parse_valor(vals[11])

        vals_totales = list(tree.item(iid_totales, "values"))
        for idx in range(1, 7):
            vals_totales[idx] = ""
        vals_totales[7] = _fmt_pesos(suma_total)
        vals_totales[8] = _fmt_pesos(suma_reporte)
        vals_totales[9] = _fmt_pesos(suma_dif_reporte)
        vals_totales[10] = _fmt_pesos(suma_sfa)
        vals_totales[11] = _fmt_pesos(suma_dif_sfa)
        tree.item(iid_totales, values=vals_totales)

    def _recalc_row(iid: str):
        vals = list(tree.item(iid, "values"))

        if _es_fila_totales_anti_rec(vals):
            _actualizar_fila_totales_anti_rec()
            _aplicar_zebra_anticipos(tree)
            _request_diff_render()
            return

        total = sum(_parse_valor(vals[i]) for i in range(1, 7))
        vals[7] = _fmt_pesos(total)

        control_reporte = _parse_valor(vals[8])
        diff_reporte = _fmt_pesos(control_reporte - total) if str(vals[8]).strip() else ""
        vals[9] = diff_reporte

        control_sfa = _parse_valor(vals[10])
        diff_sfa = _fmt_pesos(control_sfa - total) if str(vals[10]).strip() else ""
        vals[11] = diff_sfa

        tree.item(iid, values=vals)
        _actualizar_fila_totales_anti_rec()
        _aplicar_zebra_anticipos(tree)
        _request_diff_render()

    def _buscar_iid_por_concepto(concepto: str) -> str | None:
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

    # -------------------------
    # Import días: Z061->Anticipos, Z062->por juego
    # -------------------------
    def _importar_resumen_a_columna(ruta: str, target_col_idx: int, semana_objetivo: str):
        """Carga Z061 en Anticipos y Z062 en filas por código, en la columna target_col_idx."""
        try:
            resumen = leer_json_sfa(ruta)
        except Exception as e:
            messagebox.showerror("Importación", f"No pude leer el JSON/TXT SFA:\n{e}")
            return False

        # limpiar columna objetivo
        for iid in tree.get_children():
            _set_valor(iid, target_col_idx, None)

        total_z061 = 0.0
        z062_por_codigo: dict[str, float] = {}

        for (juego, _sorteo, concepto), importe in (resumen or {}).items():
            c = str(concepto).strip().upper()
            j = str(juego).strip().lstrip("0") or "0"
            if c == "Z061":
                total_z061 += float(importe or 0.0)
            elif c == "Z062":
                z062_por_codigo[j] = z062_por_codigo.get(j, 0.0) + float(importe or 0.0)

        iid_anticipos = _buscar_iid_por_concepto("Anticipos")
        if iid_anticipos is not None:
            _set_valor(iid_anticipos, target_col_idx, total_z061)

        for codigo, fila in CODIGO_A_FILA_Z062.items():
            iid = _buscar_iid_por_concepto(fila)
            if iid is None:
                continue
            _set_valor(iid, target_col_idx, z062_por_codigo.get(codigo) if codigo in z062_por_codigo else None)

        _guardar_semana(semana_objetivo)
        return True

    def _importar_resumen_dia(nombre_dia: str):
        nonlocal current_semana
        semana_objetivo = current_semana

        ruta = filedialog.askopenfilename(
            title=f"Importar resumen diario SFA ({nombre_dia.capitalize()})",
            filetypes=[("JSON o TXT", "*.json;*.JSON;*.txt;*.TXT"), ("Todos", "*.*")],
        )
        if not ruta:
            return

        col_idx = cols.index(nombre_dia)
        ok = _importar_resumen_a_columna(ruta, col_idx, semana_objetivo)
        if ok and estado_var is not None:
            base = ruta.replace("\\", "/").split("/")[-1]
            estado_var.set(f"{semana_objetivo}: importado {nombre_dia.capitalize()} ({base}).")

    # -------------------------
    # Prescripto SFA (botón dedicado)
    # -------------------------
    def _importar_prescripto_sfa():
        nonlocal current_semana
        semana_objetivo = current_semana

        ruta = filedialog.askopenfilename(
            title="Importar Prescripto",
            filetypes=[("JSON o TXT", "*.json;*.JSON;*.txt;*.TXT"), ("Todos", "*.*")],
        )
        if not ruta:
            return

        col_idx = cols.index("sfa_prescripto")
        ok = _importar_resumen_a_columna(ruta, col_idx, semana_objetivo)
        if ok and estado_var is not None:
            base = ruta.replace("\\", "/").split("/")[-1]
            estado_var.set(f"{semana_objetivo}: importado Prescripto ({base}).")


    def _cargar_control_reporte_desde_tobill():
        nonlocal current_semana
        semana_objetivo = current_semana
        try:
            payload_semana = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            sem_n = int(payload_semana.get("semana", 0) or 0)
            if 1 <= sem_n <= 5:
                semana_objetivo = f"Semana {sem_n}"
        except Exception:
            pass
        if semana_objetivo in SEMANAS and semana_objetivo != current_semana:
            _guardar_semana(current_semana)
            current_semana = semana_objetivo
            try:
                combo_semana.set(semana_objetivo)
            except Exception:
                pass
            _cargar_semana(semana_objetivo)

        col_rep_idx = cols.index("reporte_prescripto")
        col_diff_rep_idx = cols.index("diferencia_reporte")

        # IMPORTANTE: NO tocar sfa_prescripto acá.
        # Esa columna se completa con el botón "Prescripto" (archivo SFA Alta).
        for iid in tree.get_children():
            vals = list(tree.item(iid, "values"))
            vals[col_rep_idx] = ""
            vals[col_diff_rep_idx] = ""
            tree.item(iid, values=vals)
            _recalc_row(iid)

        sale_limit_por_semana = getattr(app_state, "reporte_tobill_sale_limit_por_juego_por_semana", {}) or {}
        if isinstance(sale_limit_por_semana, dict):
            sale_limit = sale_limit_por_semana.get(semana_objetivo, {})
        else:
            sale_limit = {}
        if not isinstance(sale_limit, dict):
            sale_limit = {}
        if not sale_limit:
            sale_limit = getattr(app_state, "reporte_tobill_sale_limit_por_juego", {}) or {}

        advance_por_semana = getattr(app_state, "reporte_tobill_advance_importe_por_semana", {}) or {}
        if isinstance(advance_por_semana, dict):
            advance = float(advance_por_semana.get(semana_objetivo, 0.0) or 0.0)
        else:
            advance = 0.0
        if advance == 0.0:
            advance = float(getattr(app_state, "reporte_tobill_advance_importe", 0.0) or 0.0)

        iid_anticipos = _buscar_iid_por_concepto("Anticipos")
        if iid_anticipos is not None and advance != 0.0:
            _set_valor(iid_anticipos, col_rep_idx, advance)

        for juego_tab, fila in JUEGO_PLANILLA_A_FILA_ANTI_REC.items():
            iid = _buscar_iid_por_concepto(fila)
            if iid is None:
                continue
            importe = float(sale_limit.get(juego_tab, 0.0) or 0.0)
            if importe != 0.0:
                _set_valor(iid, col_rep_idx, importe)
            else:
                _set_valor(iid, col_rep_idx, None)

        _guardar_semana(semana_objetivo)
        if estado_var is not None:
            estado_var.set(f"{semana_objetivo}: actualizado Control Reporte (Tobill: SALE LIMIT + ADVANCE).")

    # -------------------------
    # Editor inline
    # -------------------------
    edit_entry = None

    editable_col_indices = {1, 2, 3, 4, 5, 6, 8, 10}
    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)

    def _close_editor(save=True):
        nonlocal edit_entry, current_semana
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
                    _aplicar_zebra_anticipos(tree)
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
            if _es_fila_totales_anti_rec(vals):
                continue
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
        row_ids = [iid for iid in tree.get_children() if not _es_fila_totales_anti_rec(list(tree.item(iid, "values")))]
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
            _aplicar_zebra_anticipos(tree)
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
        restore_undo_snapshot(tree, snapshot)
        for iid, _vals in (snapshot.get("rows", []) or []):
            _recalc_row(str(iid))
        _aplicar_zebra_anticipos(tree)
        _request_diff_render()
        _guardar_semana(current_semana)
        return "break"

    def _on_double_click(evt):
        nonlocal edit_entry
        _close_editor(save=True)

        region = tree.identify("region", evt.x, evt.y)
        if region != "cell":
            return

        col = tree.identify_column(evt.x)
        iid = tree.identify_row(evt.y)
        if not iid:
            return

        col_idx = int(col.replace("#", "")) - 1
        if col_idx not in editable_col_indices:
            return

        clipboard_state["cell"] = (iid, col_idx)

        if _es_fila_totales_anti_rec(list(tree.item(iid, "values"))):
            return

        x, y, w, h = tree.bbox(iid, col)
        if w <= 0 or h <= 0:
            return

        val = tree.item(iid, "values")[col_idx]
        edit_entry = ttk.Entry(tree)
        edit_entry.place(x=x, y=y, width=w, height=h)
        edit_entry.insert(0, val)
        edit_entry.focus_set()
        edit_entry._cell = (iid, col)
        edit_entry.bind("<Return>", lambda _e: _close_editor(save=True))
        edit_entry.bind("<Escape>", lambda _e: _close_editor(save=False))
        edit_entry.bind("<FocusOut>", lambda _e: _close_editor(save=True))

    tree.bind("<Double-1>", _on_double_click)
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")

    # Init
    _draw_group_headers()
    _draw_column_headers()

    combo_semana.bind("<<ComboboxSelected>>", _on_week_change)
    combo_semana.set("")
    _cargar_semana(current_semana)
    _cargar_control_reporte_desde_tobill()
    def _sync_combo_semanas(reset_selection: bool = False, prefer_semana: int = 0):
        values = list(_combo_values_semanas() or [])
        try:
            combo_semana.configure(values=values)
            _ajustar_ancho_combo_semana(values, str(combo_semana.get() or "").strip())
        except Exception:
            pass
        if reset_selection or not values:
            try:
                combo_semana.set("")
                combo_semana.configure(state="disabled")
                _ajustar_ancho_combo_semana([], "")
            except Exception:
                pass
            return
        try:
            combo_semana.configure(state="readonly")
        except Exception:
            pass
        visible_pref = _semana_visible(f"Semana {prefer_semana}") if 1 <= int(prefer_semana or 0) <= 5 else ""
        actual = str(combo_semana.get() or "").strip()
        target = visible_pref if visible_pref in values else (actual if actual in values else values[0])
        _ajustar_ancho_combo_semana(values, target)
        try:
            combo_semana.set(target)
        except Exception:
            pass

    _aplicar_zebra_anticipos(tree)
    _request_diff_render()
    _sync_headers()

    if not hasattr(app_state, "planilla_semana_filtro_hooks"):
        app_state.planilla_semana_filtro_hooks = {}

    def _aplicar_filtro_area_recaudacion(payload: dict):
        sem_n = 0
        try:
            sem_n = int((payload or {}).get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        if sem_n < 1:
            _sync_combo_semanas(reset_selection=True)
            return
        _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
        _week_sync_guard["active"] = True
        try:
            _on_week_change()
        finally:
            _week_sync_guard["active"] = False

    app_state.planilla_semana_filtro_hooks["anticipos_topes"] = _aplicar_filtro_area_recaudacion

    payload_inicial = getattr(app_state, "planilla_semana_filtro_actual", {})
    if isinstance(payload_inicial, dict):
        _aplicar_filtro_area_recaudacion(dict(payload_inicial))

    def _reset_anticipos_topes():
        nonlocal current_semana

        nuevo = _estado_inicial_anticipos_topes()
        datos_por_semana.clear()
        datos_por_semana.update(nuevo["semanas"])

        combo_semana.set("")
        try:
            combo_semana.configure(values=[])
            combo_semana.configure(state="disabled")
        except Exception:
            pass
        current_semana = SEMANAS[0]
        _cargar_semana(current_semana)
        _persistir_estado()
        _sync_headers()

    if not hasattr(app_state, "planilla_anticipos_reset_hooks"):
        app_state.planilla_anticipos_reset_hooks = {}
    app_state.planilla_anticipos_reset_hooks["anticipos_topes"] = _reset_anticipos_topes

    if not hasattr(app_state, "planilla_anticipos_refresh_hooks"):
        app_state.planilla_anticipos_refresh_hooks = {}
    app_state.planilla_anticipos_refresh_hooks["anticipos_topes"] = _cargar_control_reporte_desde_tobill

    if not hasattr(app_state, "planilla_bundle_snapshot_hooks"):
        app_state.planilla_bundle_snapshot_hooks = {}
    if not hasattr(app_state, "planilla_bundle_load_hooks"):
        app_state.planilla_bundle_load_hooks = {}

    def _snapshot_anticipos_topes() -> dict:
        _guardar_semana(current_semana)
        return _normalizar_payload_anticipos_topes(app_state.planilla_anticipos_topes_data)

    def _cargar_snapshot_anticipos_topes(payload: dict):
        nonlocal current_semana
        normalizado = _normalizar_payload_anticipos_topes(payload)
        app_state.planilla_anticipos_topes_data = normalizado

        datos_por_semana.clear()
        datos_por_semana.update(normalizado["semanas"])

        current_semana = normalizado.get("current_semana", SEMANAS[0])
        if current_semana not in SEMANAS:
            current_semana = SEMANAS[0]
        combo_semana.set(_semana_visible(current_semana))
        _cargar_semana(current_semana)
        _sync_totales_txt_anticipos_topes(current_semana)
        _sync_headers()

    app_state.planilla_bundle_snapshot_hooks["anticipos_topes"] = _snapshot_anticipos_topes
    app_state.planilla_bundle_load_hooks["anticipos_topes"] = _cargar_snapshot_anticipos_topes

    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    _refresh_visual_job = {"id": None}

    def _refresh_visual_anticipos_topes():
        try:
            payload = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            sem_n = int(payload.get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        if sem_n < 1:
            _sync_combo_semanas(reset_selection=True)
        else:
            _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
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

    app_state.planilla_visual_refresh_hooks["Anticipos y Topes"] = _refresh_visual_anticipos_topes

    if estado_var is not None:
        estado_var.set("Anticipos y Topes listo.")
