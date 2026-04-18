from __future__ import annotations

import tkinter as tk
from tkinter import ttk

_STYLE_NAME = "Planilla.ProFilter.TCombobox"


def _ensure_style(widget: tk.Widget, style_name: str = _STYLE_NAME) -> str:
    style = ttk.Style(widget)
    style.configure(
        style_name,
        padding=(8, 5),
        arrowsize=15,
        fieldbackground="#FFFFFF",
        background="#FFFFFF",
        foreground="#0F172A",
        bordercolor="#8FAECC",
        lightcolor="#C4D6EA",
        darkcolor="#8FAECC",
        arrowcolor="#1A3A6B",
        relief="flat",
    )
    style.map(
        style_name,
        fieldbackground=[("readonly", "#FFFFFF"), ("focus", "#FFFFFF"), ("disabled", "#E9EFF6")],
        background=[("active", "#F3F8FE"), ("readonly", "#FFFFFF"), ("disabled", "#E9EFF6")],
        foreground=[("disabled", "#94A3B8"), ("readonly", "#0F172A")],
        bordercolor=[("focus", "#4E7FB2"), ("active", "#6E96BF"), ("!focus", "#8FAECC")],
        lightcolor=[("focus", "#7FA7D0"), ("active", "#D3E2F1"), ("!focus", "#C4D6EA")],
        darkcolor=[("focus", "#4E7FB2"), ("active", "#6E96BF"), ("!focus", "#8FAECC")],
        arrowcolor=[("active", "#235291"), ("focus", "#1D4ED8"), ("disabled", "#94A3B8"), ("!disabled", "#1A3A6B")],
    )
    return style_name


def _style_popdown(cb: ttk.Combobox) -> None:
    try:
        popdown = cb.tk.eval(f'ttk::combobox::PopdownWindow {cb}')
        listbox = f"{popdown}.f.l"
        scrollbar = f"{popdown}.f.sb"
        cb.tk.call(
            listbox,
            "configure",
            "-font", "{Segoe UI} 10",
            "-background", "#FFFFFF",
            "-foreground", "#0F172A",
            "-selectbackground", "#1D4ED8",
            "-selectforeground", "#FFFFFF",
            "-highlightthickness", 0,
            "-borderwidth", 0,
            "-relief", "flat",
            "-activestyle", "none",
        )
        try:
            cb.tk.call(
                scrollbar,
                "configure",
                "-background", "#C7D7EA",
                "-activebackground", "#9FB6CD",
                "-troughcolor", "#EEF4FC",
                "-borderwidth", 0,
                "-relief", "flat",
                "-width", 12,
            )
        except Exception:
            pass
    except Exception:
        pass


def apply_filter_combobox_style(cb: ttk.Combobox, *, style_name: str = _STYLE_NAME) -> str:
    style_name = _ensure_style(cb, style_name)
    try:
        cb.configure(style=style_name)
    except Exception:
        return style_name

    def _on_enter(_evt=None):
        try:
            cb.state(["active"])
        except Exception:
            pass

    def _on_leave(_evt=None):
        try:
            cb.state(["!active"])
        except Exception:
            pass

    cb.bind("<Enter>", _on_enter, add="+")
    cb.bind("<Leave>", _on_leave, add="+")
    cb.bind("<Button-1>", lambda _evt=None: _style_popdown(cb), add="+")
    cb.bind("<FocusIn>", lambda _evt=None: _style_popdown(cb), add="+")
    try:
        cb.configure(postcommand=lambda cb=cb: _style_popdown(cb))
    except Exception:
        pass
    return style_name
