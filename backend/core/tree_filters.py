from __future__ import annotations

import tkinter as tk
from tkinter import messagebox, ttk


def habilitar_filtros_en_tree(
    tree: ttk.Treeview,
    columnas: list[str] | tuple[str, ...],
    root,
    estado_var=None,
    on_after_filter=None,
):
    """Habilita filtro por click en encabezado para TODAS las columnas del treeview.

    Retorna una función `sync()` que debe llamarse luego de recargar datos
    para reindexar filas y reaplicar filtros activos.
    """

    filtros: dict[str, str] = {}
    all_iids: list[str] = []
    headers_base = {c: str(tree.heading(c, "text") or c) for c in columnas}

    def _update_headers():
        for c in columnas:
            base = headers_base.get(c, c)
            txt = f"{base} 🔍" if c in filtros else base
            tree.heading(c, text=txt)

    def _match(iid: str) -> bool:
        vals = [str(v) for v in tree.item(iid, "values")]
        pos = {c: i for i, c in enumerate(columnas)}
        for col, needle in filtros.items():
            idx = pos.get(col)
            if idx is None:
                return False
            v = vals[idx] if idx < len(vals) else ""
            if needle.lower() not in str(v).lower():
                return False
        return True

    def _apply():
        visibles = set(tree.get_children(""))
        insert_idx = 0
        for iid in all_iids:
            if _match(iid):
                if iid not in visibles:
                    tree.reattach(iid, "", "end")
                tree.move(iid, "", insert_idx)
                insert_idx += 1
            else:
                if iid in visibles:
                    tree.detach(iid)

        _update_headers()
        if callable(on_after_filter):
            try:
                on_after_filter()
            except Exception:
                pass

    def _show_filter_dialog(col: str):
        win = tk.Toplevel(root)
        win.title(f"Filtro: {headers_base.get(col, col)}")
        win.transient(root)
        win.grab_set()

        ttk.Label(win, text=f"Filtrar '{headers_base.get(col, col)}' (contiene):").grid(
            row=0, column=0, columnspan=3, padx=10, pady=(10, 4), sticky="w"
        )

        entry = ttk.Entry(win, width=34)
        entry.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10), sticky="ew")
        entry.insert(0, filtros.get(col, ""))
        entry.focus_set()

        def _aplicar():
            txt = entry.get().strip()
            if txt:
                filtros[col] = txt
            else:
                filtros.pop(col, None)
            _apply()
            if estado_var is not None:
                estado_var.set(f"Filtro aplicado en {headers_base.get(col, col)}.")
            win.destroy()

        def _limpiar_col():
            filtros.pop(col, None)
            _apply()
            if estado_var is not None:
                estado_var.set(f"Filtro limpiado en {headers_base.get(col, col)}.")
            win.destroy()

        def _limpiar_todo():
            filtros.clear()
            _apply()
            if estado_var is not None:
                estado_var.set("Filtros limpiados.")
            win.destroy()

        ttk.Button(win, text="Aplicar", command=_aplicar, style="Marino.TButton").grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")
        ttk.Button(win, text="Limpiar columna", command=_limpiar_col).grid(row=2, column=1, padx=5, pady=(0, 10))
        ttk.Button(win, text="Limpiar todo", command=_limpiar_todo).grid(row=2, column=2, padx=(5, 10), pady=(0, 10), sticky="e")

        entry.bind("<Return>", lambda _e: _aplicar())

    for c in columnas:
        tree.heading(c, command=lambda cc=c: _show_filter_dialog(cc))

    def sync():
        nonlocal all_iids
        all_iids = list(tree.get_children(""))
        _apply()

    _update_headers()
    return sync
