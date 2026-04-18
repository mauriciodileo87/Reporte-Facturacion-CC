from __future__ import annotations

import sys
import tkinter as tk
from typing import Callable


def bind_smooth_mousewheel(
    *,
    tree: tk.Widget,
    targets: list[tk.Widget] | tuple[tk.Widget, ...],
    on_scroll: Callable[[], None] | None = None,
):
    """
    Vincula rueda del mouse (vertical y horizontal con Shift) para que funcione
    de forma consistente en las distintas plataformas y zonas del contenedor.
    """
    if not targets:
        targets = [tree]

    remainder_y = 0.0
    remainder_x = 0.0

    def _delta_units(event: tk.Event) -> float:
        num = getattr(event, "num", None)
        if num == 4:
            return -1.0
        if num == 5:
            return 1.0

        delta = float(getattr(event, "delta", 0.0) or 0.0)
        if delta == 0.0:
            return 0.0

        factor = 120.0
        if sys.platform == "darwin":
            factor = 3.0
        elif abs(delta) < 120.0:
            # Ruedas/trackpads de alta resolución en Windows/Linux suelen emitir
            # deltas pequeños (p. ej. 15). Con 120 se vuelve "pesado" el scroll.
            factor = 30.0
        return -(delta / factor)

    def _is_horizontal(event: tk.Event) -> bool:
        state = int(getattr(event, "state", 0) or 0)
        num = getattr(event, "num", None)
        return bool(state & 0x0001) or num in (6, 7)

    def _scroll(event: tk.Event):
        nonlocal remainder_y, remainder_x
        units = _delta_units(event)
        if units == 0.0:
            return

        if _is_horizontal(event):
            remainder_x += units
            step_x = int(remainder_x)
            if step_x:
                step_x = max(-6, min(6, step_x))
                tree.xview_scroll(step_x, "units")
                remainder_x -= step_x
        else:
            remainder_y += units
            step_y = int(remainder_y)
            if step_y:
                step_y = max(-6, min(6, step_y))
                tree.yview_scroll(step_y, "units")
                remainder_y -= step_y

        if on_scroll is not None:
            on_scroll()
        return "break"

    for w in targets:
        if not w or not w.winfo_exists():
            continue
        w.bind("<MouseWheel>", _scroll, add="+")
        w.bind("<Shift-MouseWheel>", _scroll, add="+")
        w.bind("<Button-4>", _scroll, add="+")
        w.bind("<Button-5>", _scroll, add="+")
        w.bind("<Shift-Button-4>", _scroll, add="+")
        w.bind("<Shift-Button-5>", _scroll, add="+")
