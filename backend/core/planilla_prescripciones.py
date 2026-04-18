# tabs/planilla_prescripciones.py
from __future__ import annotations

import re
from dataclasses import dataclass
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from xml.etree import ElementTree as ET
import zipfile

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
from utils_excel import leer_excel_rows


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




def _publicar_semana_global_desde_combo(semana_txt: str, juego_txt: str = ""):
        if not hasattr(app_state, "publicar_filtro_area_recaudacion"):
            return
        semana_interna = str(semana_txt or "").strip()
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
        try:
            app_state.publicar_filtro_area_recaudacion(juego_final, semana_n, desde, hasta)
        except Exception:
            pass

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
    s = s.replace("$", "").strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def _norm_sorteo(x):
    txt = str(x).strip()
    if not txt:
        return "0"
    txt = txt.replace(",", ".")
    try:
        num = float(txt)
        if num.is_integer():
            return str(int(num))
    except Exception:
        pass
    return txt.lstrip("0") or "0"


def _map_juego_a_tab_planilla(juego_raw: str) -> str:
    j = (juego_raw or "").strip().lower()
    j = j.replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").replace("ü", "u")

    if "quiniela" in j and "ya" in j:
        return "Quiniela Ya"
    if "quiniela" in j and "poceada" in j:
        return "Poceada"
    if "quiniela" in j:
        return "Quiniela"
    if "loto" in j and "5" in j:
        return "Loto 5"
    if "loto" in j:
        return "Loto"
    if "quini" in j and "6" in j:
        return "Quini 6"
    if "brinco" in j:
        return "Brinco"
    if "tombolina" in j:
        return "Tombolina"
    if "poceada" in j:
        return "Poceada"
    if j.strip() == "lt" or "loteria tradicional" in j:
        return "LT"
    return ""


def _col_ref_a_idx(cell_ref: str) -> int:
    letras = "".join(ch for ch in str(cell_ref or "") if ch.isalpha()).upper()
    idx = 0
    for ch in letras:
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return max(idx - 1, 0)


def _parse_xlsx_shared_strings(zf: zipfile.ZipFile, ns: dict[str, str]) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    out: list[str] = []
    for si in root.findall("main:si", ns):
        textos = [node.text or "" for node in si.findall(".//main:t", ns)]
        out.append("".join(textos))
    return out


def _parse_xlsx_sheet_rows(path: str) -> list[list[str]]:
    ns = {
        "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        "rel": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        "pkgrel": "http://schemas.openxmlformats.org/package/2006/relationships",
    }

    with zipfile.ZipFile(path) as zf:
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        shared_strings = _parse_xlsx_shared_strings(zf, ns)

        hoja = workbook.find("main:sheets/main:sheet", ns)
        if hoja is None:
            return []

        rel_id = hoja.attrib.get(f"{{{ns['rel']}}}id", "")
        target_by_id = {
            rel.attrib.get("Id", ""): rel.attrib.get("Target", "")
            for rel in rels.findall("pkgrel:Relationship", ns)
        }
        target = target_by_id.get(rel_id, "worksheets/sheet1.xml").lstrip("/")
        sheet_path = target if target.startswith("xl/") else f"xl/{target}"

        sheet_root = ET.fromstring(zf.read(sheet_path))
        rows: list[list[str]] = []
        for row_node in sheet_root.findall("main:sheetData/main:row", ns):
            fila: list[str] = []
            for cell in row_node.findall("main:c", ns):
                cell_ref = cell.attrib.get("r", "")
                idx = _col_ref_a_idx(cell_ref)
                while len(fila) <= idx:
                    fila.append("")

                cell_type = cell.attrib.get("t", "")
                if cell_type == "inlineStr":
                    valor = "".join(node.text or "" for node in cell.findall(".//main:t", ns))
                else:
                    v_node = cell.find("main:v", ns)
                    valor = v_node.text if v_node is not None and v_node.text is not None else ""
                    if cell_type == "s":
                        try:
                            valor = shared_strings[int(valor)]
                        except Exception:
                            valor = ""
                fila[idx] = str(valor or "").strip()
            rows.append(fila)
        return rows


def _parse_consulta_prescripciones_rows(rows: list[list[str]]) -> dict[str, list[int]]:
    if not rows:
        return {}

    headers = rows[0]
    headers_norm = [str(h or "").strip().lower() for h in headers]

    def idx(cands: list[str]) -> int:
        for i, h in enumerate(headers_norm):
            for c in cands:
                if c in h:
                    return i
        return -1

    col_juego = idx(["juego"])
    col_sorteo = idx(["n° sorteo", "nº sorteo", "nro sorteo", "numero sorteo", "número sorteo", "sorteo"])
    if col_juego < 0 or col_sorteo < 0:
        raise ValueError("No hay columnas 'Juego' y 'N° Sorteo' en el Excel.")

    out: dict[str, set[int]] = {}
    for row in rows[1:]:
        juego_raw = str((row[col_juego] if col_juego < len(row) else "") or "").strip()
        sorteo_raw = str((row[col_sorteo] if col_sorteo < len(row) else "") or "").strip()
        if not juego_raw or not sorteo_raw:
            continue

        juego = _map_juego_a_tab_planilla(juego_raw)
        if not juego:
            continue

        digitos = "".join(ch for ch in sorteo_raw if ch.isdigit())
        if not digitos:
            continue
        sorteo = int(digitos)
        out.setdefault(juego, set()).add(sorteo)

    return {j: sorted(list(vals)) for j, vals in out.items()}


def _parse_consulta_prescripciones_excel(path: str) -> dict[str, list[int]]:
    try:
        rows = leer_excel_rows(path, read_only=True)
    except Exception as exc:
        raise ValueError(f"No se pudo leer el Excel de consulta de prescripciones. {exc}") from exc

    rows_normalized = [[str(cell or "").strip() for cell in row] for row in rows]
    return _parse_consulta_prescripciones_rows(rows_normalized)


@dataclass
class PrescWidgets:
    juego: str
    header_canvas: tk.Canvas
    header2_canvas: tk.Canvas
    filter_canvas: tk.Canvas
    tree: ttk.Treeview
    cols: list[str]
    filter_vars: dict[str, tk.StringVar]
    all_item_ids: list[str]


def _texto_filtro_normalizado(valor) -> str:
    return str(valor or "").strip().lower()


def _fila_prescripcion_coincide_filtros(pw: PrescWidgets, vals: list[str]) -> bool:
    for idx, col in enumerate(pw.cols):
        filtro = _texto_filtro_normalizado(pw.filter_vars.get(col).get() if col in pw.filter_vars else "")
        if not filtro:
            continue
        valor = _texto_filtro_normalizado(vals[idx] if idx < len(vals) else "")
        if filtro not in valor:
            return False
    return True


def _restaurar_items_prescripciones(pw: PrescWidgets):
    for idx, iid in enumerate(pw.all_item_ids):
        try:
            pw.tree.reattach(iid, "", idx)
        except Exception:
            pass


def _diff_fmt(a_val: float | None, b_val: float | None) -> str:
    if a_val is None or b_val is None:
        return ""
    return fmt_pesos(b_val - a_val)


_DIFF_COLOR_WARN_MIN = 20.0
_DIFF_COLOR_DANGER_MIN = 40.0
_DIFF_ROW_WARN_BG = "#FFF3D6"
_DIFF_ROW_DANGER_BG = "#FDE2E2"


def _clasificar_diferencia(valor: str) -> str | None:
    numero = parse_pesos(valor)
    if numero is None:
        return None
    numero = abs(float(numero))
    if numero > _DIFF_COLOR_DANGER_MIN:
        return "danger"
    if numero > _DIFF_COLOR_WARN_MIN:
        return "warn"
    return None


def _diff_foreground(valor: str) -> str:
    return "#0F172A"


def _aplicar_filtros_prescripciones(pw: PrescWidgets):
    _restaurar_items_prescripciones(pw)

    for iid in list(pw.tree.get_children()):
        vals = [str(v) for v in pw.tree.item(iid, "values")]
        sorteo_txt = (vals[0] if vals else "").strip().lower()
        if sorteo_txt == "totales":
            continue
        if not any(str(v).strip() for v in vals):
            continue
        if not _fila_prescripcion_coincide_filtros(pw, vals):
            pw.tree.detach(iid)

    _actualizar_fila_totales_prescripciones(pw)
    _aplicar_zebra_prescripciones(pw.tree)
    pw.tree.event_generate("<Configure>")


def _guardar_edicion_prescripcion(juego: str, semana_actual: int, old_vals: list[str], new_vals: list[str]):
    """
    Persiste la edición de la grilla en app_state, para que Guardar como...
    realmente tenga algo que serializar.
    """
    old_sorteo = _norm_sorteo(old_vals[0]) if old_vals and str(old_vals[0]).strip() else ""
    new_sorteo = _norm_sorteo(new_vals[0]) if new_vals and str(new_vals[0]).strip() else ""

    if not juego or not new_sorteo:
        return

    new_t = parse_pesos(new_vals[1]) if len(new_vals) > 1 else None
    new_r = parse_pesos(new_vals[2]) if len(new_vals) > 2 else None
    new_s = parse_pesos(new_vals[4]) if len(new_vals) > 4 else None

    if not hasattr(app_state, "tickets_prescripciones_por_juego") or not isinstance(app_state.tickets_prescripciones_por_juego, dict):
        app_state.tickets_prescripciones_por_juego = {}
    if not hasattr(app_state, "reporte_prescripciones_por_juego") or not isinstance(app_state.reporte_prescripciones_por_juego, dict):
        app_state.reporte_prescripciones_por_juego = {}
    if not hasattr(app_state, "sfa_prescripciones_por_juego") or not isinstance(app_state.sfa_prescripciones_por_juego, dict):
        app_state.sfa_prescripciones_por_juego = {}
    if not hasattr(app_state, "prescripciones_sorteos_por_semana_por_juego") or not isinstance(app_state.prescripciones_sorteos_por_semana_por_juego, dict):
        app_state.prescripciones_sorteos_por_semana_por_juego = {}

    tickets = app_state.tickets_prescripciones_por_juego.setdefault(juego, {})
    reporte = app_state.reporte_prescripciones_por_juego.setdefault(juego, {})
    sfa = app_state.sfa_prescripciones_por_juego.setdefault(juego, {})

    if old_sorteo and old_sorteo != new_sorteo:
        tickets.pop(old_sorteo, None)
        reporte.pop(old_sorteo, None)
        sfa.pop(old_sorteo, None)

    if new_t is None:
        tickets.pop(new_sorteo, None)
    else:
        tickets[new_sorteo] = float(new_t)

    if new_r is None:
        reporte.pop(new_sorteo, None)
    else:
        reporte[new_sorteo] = float(new_r)

    if new_s is None:
        sfa.pop(new_sorteo, None)
    else:
        sfa[new_sorteo] = float(new_s)

    base = app_state.prescripciones_sorteos_por_semana_por_juego.setdefault(juego, {})
    existentes_raw = base.get(semana_actual, [])
    existentes: list[int] = []
    for s in existentes_raw if isinstance(existentes_raw, list) else []:
        try:
            existentes.append(int(s))
        except Exception:
            continue

    if old_sorteo and old_sorteo != new_sorteo:
        try:
            old_int = int(old_sorteo)
            existentes = [x for x in existentes if x != old_int]
        except Exception:
            pass

    try:
        new_int = int(new_sorteo)
        if new_int not in existentes:
            existentes.append(new_int)
    except Exception:
        pass

    base[semana_actual] = sorted(set(existentes))

    for hook in getattr(app_state, "planilla_presc_refresh_hooks", {}).values():
        if callable(hook):
            try:
                hook()
            except Exception:
                pass

    for hook in getattr(app_state, "planilla_totales_refresh_hooks", {}).values():
        if callable(hook):
            try:
                hook()
            except Exception:
                pass


def _actualizar_fila_totales_prescripciones(pw: PrescWidgets):
    ids = list(pw.tree.get_children())
    if not ids:
        return

    last_data_idx = -1
    sumas = [0.0] * (len(pw.cols) - 1)

    for idx, iid in enumerate(ids):
        vals = [str(v) for v in pw.tree.item(iid, "values")]
        if vals and str(vals[0]).strip().lower() == "totales":
            continue
        try:
            int(str(vals[0]).strip())
            es_sorteo = True
        except Exception:
            es_sorteo = False
        if not es_sorteo:
            continue

        last_data_idx = idx
        for col in range(1, len(pw.cols)):
            n = parse_pesos(vals[col])
            if n is not None:
                sumas[col - 1] += float(n)

    for iid in ids:
        vals = [str(v) for v in pw.tree.item(iid, "values")]
        if vals and str(vals[0]).strip().lower() == "totales":
            pw.tree.item(iid, values=[""] * len(pw.cols))

    if last_data_idx < 0:
        return
    idx_total = last_data_idx + 1
    if idx_total >= len(ids):
        return

    row = [""] * len(pw.cols)
    row[0] = "Totales"
    for col in range(1, len(pw.cols)):
        row[col] = fmt_pesos(sumas[col - 1])
    pw.tree.item(ids[idx_total], values=row)


def _aplicar_zebra_prescripciones(tree: ttk.Treeview):
    tree.tag_configure("even", background="#FFFFFF", foreground="#0F172A")
    tree.tag_configure("odd", background="#F8FAFC", foreground="#0F172A")
    tree.tag_configure("empty", foreground="#9CA3AF")
    tree.tag_configure("total", background="#FFF6BF", foreground="#1F2937", font=("Segoe UI Semibold", 9))
    tree.tag_configure("diff_warn", background=_DIFF_ROW_WARN_BG, foreground="#0F172A")
    tree.tag_configure("diff_danger", background=_DIFF_ROW_DANGER_BG, foreground="#0F172A")

    for idx, iid in enumerate(tree.get_children()):
        vals = [str(v) for v in tree.item(iid, "values")]
        has_data = any(str(v).strip() for v in vals)
        es_total = bool(vals and str(vals[0]).strip().lower() == "totales")
        base_tag = "even" if idx % 2 == 0 else "odd"

        if es_total:
            tree.item(iid, tags=("total",))
            continue

        if not has_data:
            tree.item(iid, tags=(base_tag, "empty"))
            continue

        diffs = []
        if len(vals) > 3:
            diffs.append(_clasificar_diferencia(vals[3]))
        if len(vals) > 5:
            diffs.append(_clasificar_diferencia(vals[5]))

        if "danger" in diffs:
            tree.item(iid, tags=("diff_danger",))
        elif "warn" in diffs:
            tree.item(iid, tags=("diff_warn",))
        else:
            tree.item(iid, tags=(base_tag,))


def _crear_prescripciones_tab(
    parent: ttk.Frame,
    juego: str,
    get_semana_actual,
    estado_var=None,
) -> PrescWidgets:
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(0, weight=1)

    style = ttk.Style(parent)
    tree_style = f"PlanillaPresc.{juego}.Treeview"
    style.configure(
        tree_style,
        rowheight=27,
        font=("Segoe UI", 9),
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        bordercolor="#D5DEE8",
        lightcolor="#D5DEE8",
        darkcolor="#D5DEE8",
        foreground="#0F172A",
    )
    style.map(
        tree_style,
        background=[("selected", "#DCEBFF")],
        foreground=[("selected", "#0F172A")],
    )

    cols = ["sorteo", "t_presc", "r_presc", "d_tr", "s_presc", "d_sr"]

    cont = ttk.Frame(parent)
    cont.grid(row=0, column=0, sticky="nsew")
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

    def _darken(hex_color: str, factor: float = 0.82) -> str:
        h = hex_color.lstrip("#")
        r = int(h[0:2], 16)
        g = int(h[2:4], 16)
        b = int(h[4:6], 16)
        r = max(0, min(255, int(r * factor)))
        g = max(0, min(255, int(g * factor)))
        b = max(0, min(255, int(b * factor)))
        return f"#{r:02X}{g:02X}{b:02X}"

    BORDER_GRIS = _darken(GRIS_BASE, 0.78)
    BORDER_TICKETS = _darken(COLOR_TICKETS, 0.78)
    BORDER_REPORTE = _darken(COLOR_REPORTE, 0.78)
    BORDER_DIF = _darken(COLOR_DIF, 0.78)
    BORDER_SFA = _darken(COLOR_SFA, 0.78)

    W_SORTEO = 130
    W_COL = 145
    widths_base = [W_SORTEO, W_COL, W_COL, W_COL, W_COL, W_COL]

    H1 = 28
    H2 = 26
    header_canvas = tk.Canvas(cont, height=H1, highlightthickness=0, bg=GRIS_BASE)
    header_canvas.grid(row=0, column=0, sticky="ew")

    header2_canvas = tk.Canvas(cont, height=H2, highlightthickness=0, bg=GRIS_BASE)
    header2_canvas.grid(row=1, column=0, sticky="ew")

    H3 = 0
    filter_canvas = tk.Canvas(cont, height=H3, highlightthickness=0, bg="#F5F7FA")

    tree = ttk.Treeview(cont, columns=cols, show="tree", height=18, style=tree_style)
    tree.grid(row=3, column=0, sticky="nsew")
    tree.column("#0", width=0, stretch=False)
    tree.heading("#0", text="")

    vs = ttk.Scrollbar(cont, orient="vertical", command=tree.yview)
    vs.grid(row=3, column=1, sticky="ns")
    tree.configure(yscrollcommand=vs.set)

    hs = ttk.Scrollbar(cont, orient="horizontal", command=tree.xview)
    hs.grid(row=4, column=0, sticky="ew")

    filter_vars = {col: tk.StringVar() for col in cols}
    header_labels = {
        "sorteo": "Sorteo",
        "t_presc": "Consulta de Tickets / Prescripto",
        "r_presc": "Control de Reporte / Prescripto",
        "d_tr": "Diferencias / Tickets-Reporte",
        "s_presc": "Control SFA / Prescripto",
        "d_sr": "Diferencias / SFA-Reporte",
    }

    for _ in range(120):
        tree.insert("", "end", text="", values=[""] * len(cols))

    _aplicar_zebra_prescripciones(tree)

    resize_job = None
    last_x0 = None

    def _draw_headers(widths: list[int]):
        header_canvas.delete("all")
        header2_canvas.delete("all")
        filter_canvas.delete("all")

        x = 0
        w0 = widths[0]
        header_canvas.create_rectangle(x, 0, x + w0, H1, fill=GRIS_BASE, outline=BORDER_GRIS, width=2)
        x += w0

        w = widths[1]
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_TICKETS, outline=BORDER_TICKETS, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Consulta de Tickets", font=("Segoe UI Semibold", 9))
        x += w

        w = widths[2]
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_REPORTE, outline=BORDER_REPORTE, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Control de Reporte", font=("Segoe UI Semibold", 9))
        x += w

        w = widths[3]
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_DIF, outline=BORDER_DIF, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Diferencias", font=("Segoe UI Semibold", 9))
        x += w

        w = widths[4]
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_SFA, outline=BORDER_SFA, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Control SFA", font=("Segoe UI Semibold", 9))
        x += w

        w = widths[5]
        header_canvas.create_rectangle(x, 0, x + w, H1, fill=COLOR_DIF, outline=BORDER_DIF, width=2)
        header_canvas.create_text(x + w / 2, H1 / 2, text="Diferencias", font=("Segoe UI Semibold", 9))

        texts = ["Sorteo", "Prescripto", "Prescripto", "Tickets-Reporte", "Prescripto", "SFA-Reporte"]
        bgs = [GRIS_BASE, COLOR_TICKETS, COLOR_REPORTE, COLOR_DIF, COLOR_SFA, COLOR_DIF]
        borders = [BORDER_GRIS, BORDER_TICKETS, BORDER_REPORTE, BORDER_DIF, BORDER_SFA, BORDER_DIF]

        x = 0
        for tx, wpx, bg, br in zip(texts, widths, bgs, borders):
            header2_canvas.create_rectangle(x, 0, x + wpx, H2, fill=bg, outline=br, width=1)
            header2_canvas.create_text(x + wpx / 2, H2 / 2, text=tx, font=("Segoe UI Semibold", 9))
            x += wpx

        total = sum(widths)
        header_canvas.configure(scrollregion=(0, 0, total, H1))
        header2_canvas.configure(scrollregion=(0, 0, total, H2))
        filter_canvas.configure(scrollregion=(0, 0, total, H3))

    def _set_column_widths() -> list[int]:
        widths = list(widths_base)
        available = max(0, cont.winfo_width() - 24)
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

        tree.column("sorteo", width=widths[0], anchor="center", stretch=False)
        for i, c in enumerate(cols[1:], start=1):
            tree.column(c, width=widths[i], anchor="e", stretch=False)

        return widths

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
        entry_var = tk.StringVar(value=filter_vars[col].get())
        entry = ttk.Entry(win, width=34, textvariable=entry_var, justify=justify)
        entry.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="ew")
        entry.focus_set()
        entry.icursor("end")

        def _aplicar():
            filter_vars[col].set(str(entry_var.get() or "").strip())
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
        widths = _set_column_widths()
        _draw_headers(widths)
        _sync_headers(force=True)

    def _request_layout_sync(_evt=None):
        nonlocal resize_job
        if resize_job is not None:
            return
        resize_job = cont.after_idle(_apply_layout_sync)

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
        if "total" in tags:
            return "#FFF6BF"
        if "odd" in tags:
            return "#F8FAFC"
        return "#FFFFFF"

    def _render_diff_labels():
        _clear_diff_labels(destroy=False)
        return

    def _render_diff_labels_idle():
        nonlocal diff_render_job, diff_render_after_job
        diff_render_after_job = None
        diff_render_job = None
        _render_diff_labels()

    def _request_diff_render(_evt=None):
        nonlocal diff_render_job, diff_render_after_job
        if diff_render_job is not None or diff_render_after_job is not None:
            return
        diff_render_after_job = cont.after(32, lambda: cont.after_idle(_render_diff_labels_idle))

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
        targets=(tree, cont, header_canvas, header2_canvas, filter_canvas),
        on_scroll=_request_diff_render,
    )
    tree.bind("<Configure>", _request_layout_sync)
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
    cont.bind("<Configure>", _request_layout_sync)

    editable_cols = {"sorteo", "t_presc", "r_presc", "s_presc"}
    editable_col_indices = {idx for idx, name in enumerate(cols) if name in editable_cols}
    clipboard_state = bind_active_cell_tracking(tree)
    undo_state = create_undo_state(limit=100)

    def _normalizar_valor_celda(col_idx: int, valor_raw: str) -> str:
        nuevo = str(valor_raw or "").strip()
        if cols[col_idx] == "sorteo":
            digitos = "".join(ch for ch in nuevo if ch.isdigit())
            return str(int(digitos)) if digitos else ""
        parsed = parse_pesos(nuevo)
        return fmt_pesos(parsed) if parsed is not None else (nuevo if nuevo else "")

    def _recalcular_prescripciones_ui():
        _actualizar_fila_totales_prescripciones(pw)
        _aplicar_zebra_prescripciones(tree)
        _request_diff_render()

    def _copiar_celdas(_evt=None):
        row_iid, col_idx = get_anchor_cell(tree, clipboard_state, default_col=0)
        if not row_iid or col_idx not in editable_col_indices:
            return "break"
        seleccion = ordered_selected_rows(tree)
        row_ids = seleccion if len(seleccion) > 1 and row_iid in seleccion else [row_iid]
        matrix = []
        for iid in row_ids:
            vals = [str(v) for v in tree.item(iid, "values")]
            if vals and str(vals[0]).strip().lower() == "totales":
                continue
            while len(vals) < len(cols):
                vals.append("")
            matrix.append([vals[col_idx]])
        set_clipboard_matrix(tree, matrix)
        return "break"

    def _pegar_celdas(_evt=None):
        matrix = get_clipboard_matrix(tree)
        if not matrix:
            return "break"
        anchor_iid, anchor_col = get_anchor_cell(tree, clipboard_state, default_col=0)
        row_ids = [iid for iid in tree.get_children() if str((tree.item(iid, "values") or [""])[0]).strip().lower() != "totales"]
        if not anchor_iid or anchor_iid not in row_ids:
            return "break"
        try:
            semana_actual = int(get_semana_actual())
        except Exception:
            semana_actual = 1
        start = row_ids.index(anchor_iid)
        hubo_cambios = False
        undo_rows = []
        for r_off, row_data in enumerate(matrix):
            row_pos = start + r_off
            if row_pos >= len(row_ids):
                break
            target_iid = row_ids[row_pos]
            old_vals = [str(v) for v in tree.item(target_iid, "values")]
            vals = list(old_vals)
            while len(vals) < len(cols):
                vals.append("")
            row_changed = False
            for c_off, cell_raw in enumerate(row_data):
                col_idx = anchor_col + c_off
                if col_idx not in editable_col_indices:
                    continue
                vals[col_idx] = _normalizar_valor_celda(col_idx, cell_raw)
                row_changed = True
            if not row_changed:
                continue
            if vals == old_vals:
                continue
            t_val = parse_pesos(vals[1])
            r_val = parse_pesos(vals[2])
            s_val = parse_pesos(vals[4])
            vals[3] = _diff_fmt(t_val, r_val)
            vals[5] = _diff_fmt(s_val, r_val)
            undo_rows.append((target_iid, list(old_vals)))
            tree.item(target_iid, values=vals)
            _guardar_edicion_prescripcion(juego, semana_actual, old_vals, vals)
            hubo_cambios = True
        if hubo_cambios:
            push_undo_rows(undo_state, tree, undo_rows, meta={"semana": semana_actual, "juego": juego})
            _recalcular_prescripciones_ui()
        return "break"

    def _deshacer_celdas(_evt=None):
        snapshot = pop_undo_snapshot(undo_state)
        if not snapshot:
            return "break"
        meta = snapshot.get("meta", {}) if isinstance(snapshot, dict) else {}
        try:
            semana_undo = int(meta.get("semana", get_semana_actual()) or get_semana_actual())
        except Exception:
            semana_undo = 1
        filas_antes = {str(iid): [str(v) for v in tree.item(iid, "values")] for iid, _vals in (snapshot.get("rows", []) or [])}
        restore_undo_snapshot(tree, snapshot)
        for iid, vals_restaurados in (snapshot.get("rows", []) or []):
            vals_actuales = filas_antes.get(str(iid), [])
            _guardar_edicion_prescripcion(juego, semana_undo, vals_actuales, [str(v) for v in (vals_restaurados or [])])
        _recalcular_prescripciones_ui()
        return "break"

    def _on_double_click(evt):
        row_iid = tree.identify_row(evt.y)
        col_id = tree.identify_column(evt.x)
        if not row_iid or not col_id.startswith("#"):
            return

        col_idx = int(col_id[1:]) - 1
        if col_idx < 0 or col_idx >= len(cols):
            return
        if cols[col_idx] not in editable_cols:
            return

        clipboard_state["cell"] = (row_iid, col_idx)

        vals = [str(v) for v in tree.item(row_iid, "values")]
        if vals and str(vals[0]).strip().lower() == "totales":
            return

        bbox = tree.bbox(row_iid, col_id)
        if not bbox:
            return
        x, y, wpx, hpx = bbox
        if wpx <= 0 or hpx <= 0:
            return

        current = vals[col_idx] if col_idx < len(vals) else ""
        old_vals = list(vals)

        justify = "center" if cols[col_idx] == "sorteo" else "right"
        editor = ttk.Entry(tree, justify=justify)
        editor.insert(0, current)
        editor.select_range(0, tk.END)
        editor.focus_set()
        editor.place(x=x, y=y, width=wpx, height=hpx)

        def _commit(_evt=None):
            vals[col_idx] = _normalizar_valor_celda(col_idx, editor.get())
            t_val = parse_pesos(vals[1])
            r_val = parse_pesos(vals[2])
            s_val = parse_pesos(vals[4])
            vals[3] = _diff_fmt(t_val, r_val)
            vals[5] = _diff_fmt(s_val, r_val)

            try:
                semana_actual = int(get_semana_actual())
            except Exception:
                semana_actual = 1

            if vals != old_vals:
                push_undo_rows(undo_state, tree, [(row_iid, list(old_vals))], meta={"semana": semana_actual, "juego": juego})
                tree.item(row_iid, values=vals)
                _recalcular_prescripciones_ui()
                _guardar_edicion_prescripcion(juego, semana_actual, old_vals, vals)
            editor.destroy()

        def _cancel(_evt=None):
            editor.destroy()

        editor.bind("<Return>", _commit)
        editor.bind("<KP_Enter>", _commit)
        editor.bind("<Escape>", _cancel)
        editor.bind("<FocusOut>", _commit)

    tree.bind("<Double-1>", _on_double_click, add="+")
    tree.bind("<Control-c>", _copiar_celdas, add="+")
    tree.bind("<Control-C>", _copiar_celdas, add="+")
    tree.bind("<Control-v>", _pegar_celdas, add="+")
    tree.bind("<Control-V>", _pegar_celdas, add="+")
    tree.bind("<Control-z>", _deshacer_celdas, add="+")
    tree.bind("<Control-Z>", _deshacer_celdas, add="+")
    header2_canvas.bind("<Button-1>", _on_header2_click)
    _apply_layout_sync()
    _aplicar_zebra_prescripciones(tree)
    _request_diff_render()
    pw = PrescWidgets(
        juego=juego,
        header_canvas=header_canvas,
        header2_canvas=header2_canvas,
        filter_canvas=filter_canvas,
        tree=tree,
        cols=cols,
        filter_vars=filter_vars,
        all_item_ids=list(tree.get_children()),
    )

    def _on_filter_change(*_args):
        _aplicar_filtros_prescripciones(pw)
        _actualizar_textos_header_filtros()

    for var in filter_vars.values():
        var.trace_add("write", _on_filter_change)
    _actualizar_textos_header_filtros()

    return pw


def _refresh_prescripciones(pw: PrescWidgets, semana_actual: int, estado_var=None):
    _restaurar_items_prescripciones(pw)
    semana_key = f"Semana {max(1, min(5, int(semana_actual or 1)))}"
    reporte = {}
    if hasattr(app_state, "reporte_prescripciones_por_juego_por_semana"):
        reporte_por_semana = getattr(app_state, "reporte_prescripciones_por_juego_por_semana", {}) or {}
        if isinstance(reporte_por_semana, dict):
            semana_bucket = reporte_por_semana.get(semana_key, {}) or {}
            if isinstance(semana_bucket, dict):
                reporte = semana_bucket.get(pw.juego, {}) or {}
    if not reporte and hasattr(app_state, "reporte_prescripciones_por_juego"):
        reporte = app_state.reporte_prescripciones_por_juego.get(pw.juego, {}) or {}
    tickets = app_state.tickets_prescripciones_por_juego.get(pw.juego, {}) if hasattr(app_state, "tickets_prescripciones_por_juego") else {}
    sfa = app_state.sfa_prescripciones_por_juego.get(pw.juego, {}) if hasattr(app_state, "sfa_prescripciones_por_juego") else {}

    reporte_n = {_norm_sorteo(k): v for k, v in (reporte or {}).items()}
    tickets_n = {_norm_sorteo(k): v for k, v in (tickets or {}).items()}
    sfa_n = {_norm_sorteo(k): v for k, v in (sfa or {}).items()}

    sorteos_por_semana = getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", {}).get(pw.juego, {})
    sorteos_semana = sorteos_por_semana.get(int(semana_actual), []) if isinstance(sorteos_por_semana, dict) else []
    sorteos_base: list[str] = []
    for s in sorteos_semana:
        clave = _norm_sorteo(s)
        if clave:
            sorteos_base.append(clave)

    ids = pw.tree.get_children()
    for iid in ids:
        pw.tree.item(iid, values=[""] * len(pw.cols))

    for idx, sorteo in enumerate(sorteos_base):
        if idx >= len(ids):
            break

        r_val = float(reporte_n.get(sorteo, 0.0) or 0.0)
        t_raw = tickets_n.get(sorteo)
        s_raw = sfa_n.get(sorteo)
        t_val = float(t_raw) if t_raw is not None else None
        s_val = float(s_raw) if s_raw is not None else None

        row = ["", "", "", "", "", ""]
        row[0] = sorteo
        row[1] = fmt_pesos(t_val) if t_val is not None else ""
        row[2] = fmt_pesos(r_val)
        row[3] = _diff_fmt(t_val, r_val)
        row[4] = fmt_pesos(s_val) if s_val is not None else ""
        row[5] = _diff_fmt(s_val, r_val)
        pw.tree.item(ids[idx], values=row)

    _actualizar_fila_totales_prescripciones(pw)
    _aplicar_zebra_prescripciones(pw.tree)
    pw.tree.event_generate("<Configure>")
    _aplicar_filtros_prescripciones(pw)


def _clear_prescripciones(pw: PrescWidgets):
    _restaurar_items_prescripciones(pw)
    ids = pw.tree.get_children()
    for iid in ids:
        pw.tree.item(iid, values=[""] * len(pw.cols))
    _actualizar_fila_totales_prescripciones(pw)
    _aplicar_zebra_prescripciones(pw.tree)
    pw.tree.event_generate("<Configure>")


def _ensure_prescripciones_toolbar_style(parent):
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


def build_prescripciones(fr_seccion: ttk.Frame, estado_var):
    fr_seccion.columnconfigure(0, weight=1)
    fr_seccion.rowconfigure(1, weight=1)

    juegos_nombres = [j for j, _ in JUEGOS]
    presc_widgets_por_juego: dict[str, PrescWidgets] = {}

    top = ttk.Frame(fr_seccion, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=10, pady=(6, 10))
    top.columnconfigure(6, weight=1)

    filtro_style = _ensure_prescripciones_toolbar_style(fr_seccion)

    boton_importar_style = "Marino.TButton"

    ttk.Label(top, text="Juego:", style="PanelLabel.TLabel").grid(row=0, column=0, sticky="w")
    combo = ttk.Combobox(top, state="readonly", values=juegos_nombres, width=30, style=filtro_style)
    combo.grid(row=0, column=1, sticky="w", padx=(8, 10))

    ttk.Label(top, text="Semana:", style="PanelLabel.TLabel").grid(row=0, column=2, sticky="e", padx=(12, 6))
    combo_semana = ttk.Combobox(top, state="disabled", values=_combo_values_semanas(), width=28, style=filtro_style)
    combo_semana.grid(row=0, column=3, sticky="w")
    combo_semana.set("")

    def _semana_actual() -> int:
        visible = str(combo_semana.get() or "").strip()
        if not visible:
            payload = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            try:
                n = int(payload.get("semana", 0) or 0)
                if 1 <= n <= 5:
                    return n
            except Exception:
                pass
            return 1
        interna = _semana_interna(visible)
        txt = str(interna or "").strip().lower().replace("semana", "").strip()
        try:
            n = int(txt)
            if 1 <= n <= 5:
                return n
        except Exception:
            pass
        return 1


    def _ajustar_ancho_combo_semana(valores: list[str] | None = None, texto_actual: str = ""):
        candidatos = [str(v or "").strip() for v in (valores or [])]
        if texto_actual:
            candidatos.append(str(texto_actual).strip())
        largo = max([len(x) for x in candidatos if x] or [24])
        ancho = max(24, min(28, largo + 2))
        try:
            combo_semana.configure(width=ancho)
        except Exception:
            pass

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

    def _actualizar_label_rango_del_al(semana_n: int):
        hook = getattr(app_state, "planilla_actualizar_rango_del_al_seccion", None)
        if callable(hook):
            hook(semana_n)

    def _importar_consulta_prescripciones():
        path = filedialog.askopenfilename(
            title="Importar consulta de prescripciones",
            filetypes=[("Excel", "*.xlsx;*.XLSX;*.xlsm;*.XLSM"), ("Todos", "*.*")],
        )
        if not path:
            return

        try:
            sorteos_por_juego = _parse_consulta_prescripciones_excel(path)
        except Exception as e:
            messagebox.showerror("Prescripciones", f"No se pudo importar el Excel:\n{e}")
            return

        semana = _semana_actual()
        base = getattr(app_state, "prescripciones_sorteos_por_semana_por_juego", None)
        if not isinstance(base, dict):
            app_state.prescripciones_sorteos_por_semana_por_juego = {}
            base = app_state.prescripciones_sorteos_por_semana_por_juego

        nuevos = 0
        total_semana = 0
        for juego, sorteos in sorteos_por_juego.items():
            base.setdefault(juego, {})

            existentes_raw = base[juego].get(semana, [])
            existentes: list[int] = []
            for s in existentes_raw:
                try:
                    existentes.append(int(s))
                except Exception:
                    continue

            vistos = set(existentes)
            merged = list(existentes)
            for s in sorteos:
                try:
                    n = int(s)
                except Exception:
                    continue
                if n in vistos:
                    continue
                merged.append(n)
                vistos.add(n)
                nuevos += 1

            base[juego][semana] = merged
            total_semana += len(merged)

        if estado_var is not None:
            estado_var.set(
                f"Consulta de prescripciones importada en Semana {semana}: +{nuevos} sorteos nuevos "
                f"(total semana: {total_semana})."
            )

        juego_actual = combo.get().strip()
        if juego_actual in presc_widgets_por_juego:
            _refresh_prescripciones(presc_widgets_por_juego[juego_actual], semana, estado_var)

        for hook in getattr(app_state, "planilla_totales_refresh_hooks", {}).values():
            if callable(hook):
                try:
                    hook()
                except Exception:
                    pass

    ttk.Button(
        top,
        text="Importar consulta de prescripciones",
        command=_importar_consulta_prescripciones,
        style=boton_importar_style,
    ).grid(row=0, column=4, sticky="w", padx=(12, 0))

    stack = ttk.Frame(fr_seccion)
    stack.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
    stack.columnconfigure(0, weight=1)
    stack.rowconfigure(0, weight=1)

    frames = {}
    built: set[str] = set()

    if not hasattr(app_state, "planilla_presc_refresh_hooks"):
        app_state.planilla_presc_refresh_hooks = {}

    for juego in juegos_nombres:
        fr = ttk.Frame(stack)
        fr.grid(row=0, column=0, sticky="nsew")
        fr.columnconfigure(0, weight=1)
        fr.rowconfigure(0, weight=1)
        frames[juego] = fr

    _show_job = {"id": None, "juego": ""}

    def _refresh_visual_prescripciones(juego_actual: str | None = None):
        juego_target = (juego_actual or combo.get() or "").strip()
        if not juego_target:
            return
        semana = _semana_actual()
        _actualizar_label_rango_del_al(semana)
        pw = presc_widgets_por_juego.get(juego_target)
        if pw is not None:
            _refresh_prescripciones(pw, semana, estado_var)

    def mostrar(j):
        if _show_job["id"] is not None:
            try:
                fr_seccion.after_cancel(_show_job["id"])
            except Exception:
                pass
            _show_job["id"] = None

        def _run():
            _show_job["id"] = None
            semana = _semana_actual()
            _actualizar_label_rango_del_al(semana)
            if j not in built:
                pw = _crear_prescripciones_tab(frames[j], j, _semana_actual, estado_var)
                presc_widgets_por_juego[j] = pw
                app_state.planilla_presc_refresh_hooks[j] = (lambda pp=pw: lambda: _refresh_prescripciones(pp, _semana_actual(), estado_var))()
                if not hasattr(app_state, "planilla_presc_clear_hooks"):
                    app_state.planilla_presc_clear_hooks = {}
                app_state.planilla_presc_clear_hooks[j] = (lambda pp=pw: lambda: _clear_prescripciones(pp))()
                built.add(j)
            frames[j].tkraise()
            _refresh_visual_prescripciones(j)

        try:
            _show_job["id"] = fr_seccion.after_idle(_run)
        except Exception:
            _run()

    combo.bind("<<ComboboxSelected>>", lambda _e: mostrar(combo.get()))

    def _on_semana_combo_selected(_e=None):
        semana_sel = _semana_actual()
        _publicar_semana_global_desde_combo(f"Semana {semana_sel}", combo.get().strip())
        mostrar(combo.get())

    combo_semana.bind("<<ComboboxSelected>>", _on_semana_combo_selected)

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
        else:
            _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)

        juego_actual = combo.get().strip()
        if juego_actual:
            mostrar(juego_actual)

    app_state.planilla_semana_filtro_hooks["prescripciones"] = _aplicar_filtro_area_recaudacion

    if not hasattr(app_state, "planilla_visual_refresh_hooks"):
        app_state.planilla_visual_refresh_hooks = {}

    def _refresh_visual_prescripciones_hook():
        try:
            payload = getattr(app_state, "planilla_semana_filtro_actual", {}) or {}
            sem_n = int(payload.get("semana", 0) or 0)
        except Exception:
            sem_n = 0
        try:
            if sem_n < 1:
                _sync_combo_semanas(reset_selection=True)
            else:
                _sync_combo_semanas(reset_selection=False, prefer_semana=sem_n)
        except Exception:
            pass
        try:
            _refresh_visual_prescripciones()
        except Exception:
            pass

    app_state.planilla_visual_refresh_hooks["Prescripciones"] = _refresh_visual_prescripciones_hook

    payload_inicial = getattr(app_state, "planilla_semana_filtro_actual", {})
    combo.set(juegos_nombres[0])
    if isinstance(payload_inicial, dict):
        _aplicar_filtro_area_recaudacion(dict(payload_inicial))
    else:
        _sync_combo_semanas(reset_selection=True)

    _actualizar_label_rango_del_al(_semana_actual())
    mostrar(juegos_nombres[0])

    return presc_widgets_por_juego
