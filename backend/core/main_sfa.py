# main_sfa.py
# Estética unificada con fondo de panel más oscuro como Planilla Facturación

import os
import sys
import re
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


def _quitar_elemento_focus(layout):
    limpio = []
    for element, options in layout:
        if "focus" in element.lower():
            continue
        hijos = options.get("children")
        if hijos:
            options = dict(options)
            options["children"] = _quitar_elemento_focus(hijos)
        limpio.append((element, options))
    return limpio


def configurar_estilo_app(root):
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    # =========================
    # PALETA
    # =========================
    BG_PANEL = "#B9D2EB"      # <- azul más vivo, profesional y unificado
    BG_APP = BG_PANEL
    BG_CARD = "#F3F7FC"
    FG_MAIN = "#0F172A"
    FG_MUTED = "#475569"
    FG_HINT = "#94A3B8"

    PRIMARY = "#163E72"
    PRIMARY_HOVER = "#1F528F"
    PRIMARY_PRESS = "#12365F"

    BORDER = "#8FAECC"
    BORDER_LIGHT = "#B3C8DE"
    NOTEBOOK_LINE = "#A7C3DE"

    TAB_IDLE = "#BED4EC"
    TAB_HOVER = "#AEC9E7"
    TAB_SELECTED = "#E5EEF8"
    TAB_TEXT = "#1A3A6B"

    INNER_IDLE = "#D9E7F5"
    INNER_HOVER = "#C8DCF0"
    INNER_SELECTED = "#EAF2FA"

    FIELD_BG = "#F6FAFE"
    FIELD_DISABLED = "#E5EDF6"

    root.configure(bg=BG_APP)
    root.option_add("*Font", "{Segoe UI} 9")
    root.option_add("*tearOff", False)
    root.option_add("*TButton.takeFocus", 0)
    root.option_add("*TEntry.takeFocus", 0)
    root.option_add("*TCombobox.takeFocus", 0)

    # =========================
    # FRAMES
    # =========================
    style.configure("TFrame", background=BG_PANEL)
    style.configure("Panel.TFrame", background=BG_PANEL)
    style.configure("Card.TFrame", background=BG_CARD)
    style.configure("Surface.TFrame", background=BG_PANEL)
    style.configure("TLabelframe", background=BG_PANEL, bordercolor=BORDER_LIGHT)
    style.configure("TLabelframe.Label", background=BG_PANEL, foreground=FG_MAIN)

    # =========================
    # LABELS
    # =========================
    style.configure(
        "TLabel",
        background=BG_PANEL,
        foreground=FG_MUTED,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Titulo.TLabel",
        background=BG_PANEL,
        foreground=FG_MAIN,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Section.TLabel",
        background=BG_PANEL,
        foreground=FG_MAIN,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "Hint.TLabel",
        background=BG_PANEL,
        foreground=FG_HINT,
        font=("Segoe UI", 8),
    )
    style.configure(
        "PanelLabel.TLabel",
        background=BG_PANEL,
        foreground=FG_MUTED,
        font=("Segoe UI", 9),
    )
    style.configure(
        "PanelTitle.TLabel",
        background=BG_PANEL,
        foreground=FG_MAIN,
        font=("Segoe UI Semibold", 10),
    )
    style.configure(
        "CardLabel.TLabel",
        background=BG_CARD,
        foreground=FG_MUTED,
        font=("Segoe UI", 9),
    )
    style.configure(
        "CardTitle.TLabel",
        background=BG_CARD,
        foreground=FG_MAIN,
        font=("Segoe UI Semibold", 10),
    )

    # =========================
    # BOTONES
    # =========================
    style.configure(
        "TButton",
        padding=(10, 5),
        relief="flat",
        borderwidth=1,
        focusthickness=0,
        background="#E6EEF8",
        foreground=FG_MAIN,
        bordercolor=BORDER,
        font=("Segoe UI", 9),
        cursor="hand2",
    )
    style.map(
        "TButton",
        background=[("active", "#D7E4F3"), ("pressed", "#C9D9EC")],
        bordercolor=[("active", "#7F9DBD"), ("pressed", "#718FB0")],
        foreground=[("disabled", FG_HINT), ("!disabled", FG_MAIN)],
    )

    style.configure(
        "Marino.TButton",
        padding=(10, 5),
        relief="flat",
        borderwidth=0,
        focusthickness=0,
        font=("Segoe UI Semibold", 9),
        background=PRIMARY,
        foreground="#FFFFFF",
        bordercolor=PRIMARY,
        cursor="hand2",
    )
    style.map(
        "Marino.TButton",
        background=[("active", PRIMARY_HOVER), ("pressed", PRIMARY_PRESS)],
        bordercolor=[("active", PRIMARY_HOVER), ("pressed", PRIMARY_PRESS)],
        foreground=[("disabled", "#CBD5E1"), ("!disabled", "#FFFFFF")],
    )

    style.configure(
        "SecondaryMarino.TButton",
        padding=(10, 5),
        relief="flat",
        borderwidth=1,
        focusthickness=0,
        font=("Segoe UI", 9),
        background=BG_CARD,
        foreground=PRIMARY,
        bordercolor=BORDER,
        cursor="hand2",
    )
    style.map(
        "SecondaryMarino.TButton",
        background=[("active", "#E4ECF7"), ("pressed", "#D5E1F1")],
        foreground=[("disabled", FG_HINT), ("!disabled", PRIMARY)],
        bordercolor=[("active", "#6F94BB"), ("pressed", "#5F84AA")],
    )

    # =========================
    # NOTEBOOKS
    # =========================
    style.configure(
        "TNotebook",
        background=BG_PANEL,
        borderwidth=0,
        tabmargins=(10, 8, 10, 0),
        bordercolor=NOTEBOOK_LINE,
        lightcolor=NOTEBOOK_LINE,
        darkcolor=NOTEBOOK_LINE,
        relief="flat",
    )
    style.configure(
        "TNotebook.Tab",
        padding=(14, 7),
        borderwidth=1,
        bordercolor=NOTEBOOK_LINE,
        lightcolor=NOTEBOOK_LINE,
        darkcolor=NOTEBOOK_LINE,
        focusthickness=0,
        focuscolor="",
        font=("Segoe UI Semibold", 9),
        background=TAB_IDLE,
        foreground=TAB_TEXT,
    )
    style.map(
        "TNotebook.Tab",
        background=[
            ("selected", TAB_SELECTED),
            ("active", TAB_HOVER),
            ("!selected", TAB_IDLE),
        ],
        foreground=[
            ("selected", PRIMARY),
            ("active", PRIMARY),
            ("!selected", TAB_TEXT),
        ],
        bordercolor=[
            ("selected", NOTEBOOK_LINE),
            ("active", NOTEBOOK_LINE),
            ("!selected", NOTEBOOK_LINE),
        ],
        lightcolor=[
            ("selected", NOTEBOOK_LINE),
            ("active", NOTEBOOK_LINE),
            ("!selected", NOTEBOOK_LINE),
        ],
        darkcolor=[
            ("selected", NOTEBOOK_LINE),
            ("active", NOTEBOOK_LINE),
            ("!selected", NOTEBOOK_LINE),
        ],
    )

    style.configure(
        "Inner.TNotebook",
        background=BG_PANEL,
        borderwidth=0,
        tabmargins=(10, 8, 10, 0),
        bordercolor=NOTEBOOK_LINE,
        lightcolor=NOTEBOOK_LINE,
        darkcolor=NOTEBOOK_LINE,
        relief="flat",
    )
    style.configure(
        "Inner.TNotebook.Tab",
        padding=(12, 6),
        borderwidth=1,
        background=INNER_IDLE,
        foreground=TAB_TEXT,
        font=("Segoe UI Semibold", 9),
        bordercolor=BORDER_LIGHT,
    )
    style.map(
        "Inner.TNotebook.Tab",
        background=[
            ("selected", INNER_SELECTED),
            ("active", INNER_HOVER),
            ("!selected", INNER_IDLE),
        ],
        foreground=[
            ("selected", PRIMARY),
            ("active", PRIMARY),
            ("!selected", TAB_TEXT),
        ],
        bordercolor=[
            ("selected", BORDER),
            ("active", BORDER_LIGHT),
            ("!selected", BORDER_LIGHT),
        ],
    )

    # =========================
    # ENTRADAS
    # =========================
    style.configure(
        "TEntry",
        fieldbackground=FIELD_BG,
        foreground=FG_MAIN,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        padding=(7, 4),
    )
    style.configure(
        "TCombobox",
        fieldbackground=FIELD_BG,
        background=FIELD_BG,
        foreground=FG_MAIN,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        arrowcolor=PRIMARY,
        padding=(7, 4),
    )
    style.map(
        "TEntry",
        fieldbackground=[("focus", "#FFFFFF"), ("!focus", FIELD_BG)],
        bordercolor=[("focus", "#3F74A8"), ("!focus", BORDER)],
        lightcolor=[("focus", "#5A86BC"), ("!focus", BORDER)],
        darkcolor=[("focus", "#3F74A8"), ("!focus", BORDER)],
    )
    style.map(
        "TCombobox",
        background=[("active", "#E9F2FB"), ("!active", FIELD_BG)],
        fieldbackground=[("readonly", "#FFFFFF"), ("disabled", FIELD_DISABLED)],
        foreground=[("readonly", FG_MAIN), ("disabled", FG_HINT)],
        bordercolor=[("focus", "#3F74A8"), ("!focus", BORDER)],
    )
    style.configure(
        "DateEntry",
        fieldbackground="#FFFFFF",
        background="#FFFFFF",
        foreground=FG_MAIN,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        arrowcolor=PRIMARY,
    )
    style.map(
        "DateEntry",
        fieldbackground=[("focus", "#FFFFFF"), ("!focus", "#FFFFFF")],
        bordercolor=[("focus", "#3F74A8"), ("!focus", BORDER)],
    )

    # =========================
    # SCROLLBARS
    # =========================
    for orient in ("Vertical", "Horizontal"):
        style.configure(
            f"{orient}.TScrollbar",
            gripcount=0,
            background="#AFC5DB",
            troughcolor="#E3EDF7",
            bordercolor="#AFC5DB",
            arrowcolor="#3F5F82",
            relief="flat",
        )

    # =========================
    # TREEVIEW
    # =========================
    style.configure(
        "Treeview",
        rowheight=28,
        font=("Segoe UI", 9),
        borderwidth=0,
        fieldbackground=BG_CARD,
        background=BG_CARD,
        foreground=FG_MAIN,
    )
    style.map(
        "Treeview",
        background=[("selected", "#C4DAF3")],
        foreground=[("selected", FG_MAIN)],
    )
    style.configure(
        "Treeview.Heading",
        font=("Segoe UI Semibold", 9),
        background="#DCE8F5",
        foreground=FG_MAIN,
        borderwidth=0,
        padding=(10, 7),
        relief="flat",
    )
    style.map(
        "Treeview.Heading",
        background=[("active", "#CFDEF0"), ("!active", "#DCE8F5")],
    )

    # =========================
    # STATUS BAR
    # =========================
    style.configure(
        "Status.TLabel",
        background="#DCE8F5",
        foreground="#1A3A6B",
        borderwidth=0,
        relief="flat",
        bordercolor=BG_PANEL,
        padding=(12, 6),
        font=("Segoe UI", 9),
    )

    style.configure("TSeparator", background=BG_PANEL)

    for sn in (
        "TNotebook.Tab",
        "Inner.TNotebook.Tab",
        "TButton",
        "Marino.TButton",
        "SecondaryMarino.TButton",
    ):
        try:
            layout = style.layout(sn)
            style.layout(sn, _quitar_elemento_focus(layout))
        except tk.TclError:
            pass


def _detectar_conflictos_merge(base_dir: str):
    """
    Detecta conflictos de merge reales solo en archivos del proyecto.
    Evita falsos positivos por líneas '=======' de docstrings/liberías.
    """
    conflictos = []

    archivos = []
    try:
        salida = subprocess.run(
            ["git", "ls-files", "*.py"],
            cwd=base_dir,
            capture_output=True,
            text=True,
            check=True,
        )
        archivos = [Path(base_dir) / ruta for ruta in salida.stdout.splitlines() if ruta]
    except (subprocess.SubprocessError, OSError):
        for archivo in Path(base_dir).rglob("*.py"):
            if any(
                ignorado in archivo.parts
                for ignorado in {
                    ".git",
                    ".venv",
                    "venv",
                    "__pycache__",
                    "build",
                    "dist",
                    "site-packages",
                    "Lib",
                    "Scripts",
                    "IPython",
                }
            ):
                continue
            archivos.append(archivo)

    for archivo in archivos:
        try:
            lineas = archivo.read_text(encoding="utf-8", errors="replace").splitlines()
        except OSError:
            continue

        conflicto_activo = False
        vio_medio = False
        linea_inicio = None

        for numero_linea, linea in enumerate(lineas, start=1):
            s = linea.rstrip()

            if s.startswith("<<<<<<< "):
                conflicto_activo = True
                vio_medio = False
                linea_inicio = numero_linea
            elif conflicto_activo and s == "=======":
                vio_medio = True
            elif conflicto_activo and s.startswith(">>>>>>> "):
                if vio_medio and linea_inicio is not None:
                    conflictos.append(f"- {archivo.relative_to(base_dir)}:{linea_inicio}")
                    break
                conflicto_activo = False
                vio_medio = False
                linea_inicio = None

    if conflictos:
        detalle = "\n".join(conflictos)
        raise RuntimeError(
            "Se detectaron conflictos de merge sin resolver en código Python. "
            "Corregí estas líneas y volvé a ejecutar:\n"
            f"{detalle}"
        )


def _resolver_fabrica_tab_txt_json(modulo_tab_txt_json, nombre_tab: str):
    if nombre_tab == "txt":
        candidatos = ("crear_tab_txt", "crear_tab_txt_json", "crear_tab_txt_sfa")
    else:
        candidatos = ("crear_tab_json", "crear_tab_json_sfa")

    for nombre in candidatos:
        fabrica = getattr(modulo_tab_txt_json, nombre, None)
        if callable(fabrica):
            return fabrica
    return None


def _crear_tab_segura(creador, frame, root, estado_var, nombre_tab: str):
    try:
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        surface = ttk.Frame(frame, style="Panel.TFrame")
        surface.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)
        surface.columnconfigure(0, weight=1)
        surface.rowconfigure(0, weight=1)

        creador(surface, root, estado_var)

    except Exception as e:
        for child in frame.winfo_children():
            child.destroy()

        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        cont = ttk.Frame(frame, style="Card.TFrame")
        cont.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)
        cont.columnconfigure(0, weight=1)

        ttk.Label(
            cont,
            text=f"No se pudo inicializar la sección: {nombre_tab}",
            style="CardTitle.TLabel",
        ).grid(row=0, column=0, sticky="w")

        ttk.Label(
            cont,
            text=str(e),
            wraplength=900,
            justify="left",
            style="CardLabel.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(6, 0))

        if estado_var is not None:
            estado_var.set(f"Error al inicializar '{nombre_tab}': {e}")


def _crear_tab_fallback(frame, nombre_tab: str, detalle: str):
    for child in frame.winfo_children():
        child.destroy()

    frame.columnconfigure(0, weight=1)
    frame.rowconfigure(0, weight=1)

    cont = ttk.Frame(frame, style="Card.TFrame")
    cont.grid(row=0, column=0, sticky="nsew", padx=16, pady=14)
    cont.columnconfigure(0, weight=1)

    ttk.Label(
        cont,
        text=f"No se pudo inicializar la sección: {nombre_tab}",
        style="CardTitle.TLabel",
    ).grid(row=0, column=0, sticky="w")

    ttk.Label(
        cont,
        text=detalle,
        wraplength=900,
        justify="left",
        style="CardLabel.TLabel",
    ).grid(row=1, column=0, sticky="w", pady=(6, 0))


def _crear_vista_filtrable(parent, titulo: str, opciones: list[str], on_select):
    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(1, weight=1)

    top = ttk.Frame(parent, style="Panel.TFrame")
    top.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 8))
    ttk.Label(top, text=titulo, style="PanelLabel.TLabel").pack(
        side="left", padx=(0, 6)
    )

    combo = ttk.Combobox(top, state="readonly", values=opciones, width=20)
    combo.pack(side="left")

    stack = ttk.Frame(parent, style="Panel.TFrame")
    stack.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 12))
    stack.columnconfigure(0, weight=1)
    stack.rowconfigure(0, weight=1)

    frames = {}
    for opt in opciones:
        fr = ttk.Frame(stack, style="Panel.TFrame")
        fr.grid(row=0, column=0, sticky="nsew")
        fr.columnconfigure(0, weight=1)
        fr.rowconfigure(0, weight=1)
        frames[opt] = fr

    def _on_change(_evt=None):
        selected = combo.get()
        if selected in frames:
            frames[selected].tkraise()
            on_select(selected, frames[selected])

    combo.bind("<<ComboboxSelected>>", _on_change)

    if opciones:
        combo.set(opciones[0])
        _on_change()

    return combo, frames


def main():
    _detectar_conflictos_merge(BASE_DIR)

    import tab_txt_json
    from tab_diff import crear_tab_diff
    from tab_tickets import crear_tab_tickets
    from tab_reporte import crear_tab_reporte
    from tab_facuni import crear_tab_facuni
    from tabs.tab_planilla_facturacion import (
        crear_tab_planilla_facturacion,
        guardar_bundle_default,
    )

    root = tk.Tk()
    root.title("SFA / Facturación")
    root.geometry("1250x700")
    root.minsize(1160, 660)
    configurar_estilo_app(root)

    shell = ttk.Frame(root, style="Panel.TFrame")
    shell.pack(fill="both", expand=True, padx=8, pady=(8, 8))
    shell.columnconfigure(0, weight=1)
    shell.rowconfigure(0, weight=1)

    nb_main = ttk.Notebook(shell)
    nb_main.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

    frame_sfa = ttk.Frame(nb_main, style="Panel.TFrame", padding=(6, 6, 6, 6))
    frame_tickets = ttk.Frame(nb_main, style="Panel.TFrame", padding=(6, 6, 6, 6))
    frame_rep_fact = ttk.Frame(nb_main, style="Panel.TFrame", padding=(6, 6, 6, 6))
    frame_facuni = ttk.Frame(nb_main, style="Panel.TFrame", padding=(6, 6, 6, 6))
    frame_planilla = ttk.Frame(nb_main, style="Panel.TFrame", padding=(6, 6, 6, 6))

    nb_main.add(frame_sfa, text="SFA")
    nb_main.add(frame_tickets, text="Consulta de tickets")
    nb_main.add(frame_rep_fact, text="Reporte Facturación")
    nb_main.add(frame_facuni, text="FACUNI")
    nb_main.add(frame_planilla, text="Planilla Facturación")

    estado_var = tk.StringVar(value="Listo.")

    crear_tab_txt = _resolver_fabrica_tab_txt_json(tab_txt_json, "txt")
    crear_tab_json = _resolver_fabrica_tab_txt_json(tab_txt_json, "json")

    sfa_builders = {
        "TXT": (
            crear_tab_txt,
            "SFA / TXT",
            "No se encontró una función de creación compatible en tab_txt_json.",
        ),
        "JSON / SFA": (
            crear_tab_json,
            "SFA / JSON",
            "No se encontró una función de creación compatible en tab_txt_json.",
        ),
        "Diferencias": (
            crear_tab_diff,
            "SFA / Diferencias",
            "No se encontró una función de creación compatible para diferencias.",
        ),
    }
    sfa_built = set()

    def _on_select_sfa(nombre: str, fr: ttk.Frame):
        if nombre in sfa_built:
            return
        creador, nombre_tab, detalle = sfa_builders[nombre]
        if callable(creador):
            _crear_tab_segura(creador, fr, root, estado_var, nombre_tab)
        else:
            _crear_tab_fallback(fr, nombre_tab, detalle)
        sfa_built.add(nombre)

    _crear_vista_filtrable(
        frame_sfa,
        "Sección:",
        ["TXT", "JSON / SFA", "Diferencias"],
        _on_select_sfa,
    )

    _crear_tab_segura(
        crear_tab_tickets, frame_tickets, root, estado_var, "Consulta de tickets"
    )
    _crear_tab_segura(
        crear_tab_reporte, frame_rep_fact, root, estado_var, "Reporte Facturación"
    )
    _crear_tab_segura(crear_tab_facuni, frame_facuni, root, estado_var, "FACUNI")
    _crear_tab_segura(
        crear_tab_planilla_facturacion,
        frame_planilla,
        root,
        estado_var,
        "Planilla Facturación",
    )

    # Arrancar directamente en Planilla Facturación
    nb_main.select(frame_planilla)
    
    lbl_estado = ttk.Label(
        root,
        textvariable=estado_var,    
        anchor="w",
        style="Status.TLabel",
    )
    lbl_estado.pack(fill="x", padx=0, pady=(0, 0))

    def _on_close():
        try:
            guardar_bundle_default(estado_var)
        except Exception:
            pass
        finally:
            root.destroy()

    root.protocol("WM_DELETE_WINDOW", _on_close)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        pass
    except Exception as e:
        import traceback

        traceback.print_exc()
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error inesperado", f"{e}\n\n{traceback.format_exc()}")
        except Exception:
            pass