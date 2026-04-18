from __future__ import annotations

import tkinter as tk
from tkinter import TclError


def bind_active_cell_tracking(tree: tk.Widget, state: dict | None = None) -> dict:
    state = state if isinstance(state, dict) else {}

    def _remember(event=None):
        if event is None:
            return
        try:
            region = tree.identify("region", event.x, event.y)
        except Exception:
            return
        if region != "cell":
            return
        try:
            row = tree.identify_row(event.y)
            col = tree.identify_column(event.x)
        except Exception:
            return
        if not row or not str(col).startswith("#"):
            return
        try:
            state["cell"] = (row, int(str(col)[1:]) - 1)
        except Exception:
            return

    tree.bind("<Button-1>", _remember, add="+")
    return state


def get_anchor_cell(tree: tk.Widget, state: dict | None = None, default_col: int = 0) -> tuple[str | None, int]:
    if isinstance(state, dict):
        cell = state.get("cell")
        if isinstance(cell, (list, tuple)) and len(cell) == 2:
            row, col = cell
            try:
                if row in tree.get_children(""):
                    return str(row), int(col)
            except Exception:
                pass

    try:
        selection = list(tree.selection())
    except Exception:
        selection = []
    row = selection[0] if selection else None
    if not row:
        try:
            row = tree.focus()
        except Exception:
            row = None
    if not row:
        try:
            children = list(tree.get_children(""))
            row = children[0] if children else None
        except Exception:
            row = None
    return (str(row), int(default_col)) if row else (None, int(default_col))


def ordered_selected_rows(tree: tk.Widget) -> list[str]:
    try:
        selected = set(tree.selection())
    except Exception:
        selected = set()
    if not selected:
        return []
    try:
        return [iid for iid in tree.get_children("") if iid in selected]
    except Exception:
        return list(selected)


def get_clipboard_matrix(widget: tk.Widget) -> list[list[str]]:
    try:
        raw = widget.clipboard_get()
    except TclError:
        return []
    except Exception:
        return []

    text = str(raw or "").replace("\r\n", "\n").replace("\r", "\n")
    if text.endswith("\n"):
        text = text[:-1]
    if not text:
        return []
    return [line.split("\t") for line in text.split("\n")]


def set_clipboard_matrix(widget: tk.Widget, matrix: list[list[str]]) -> None:
    rows: list[str] = []
    for row in matrix or []:
        values = ["" if value is None else str(value) for value in (row or [])]
        rows.append("\t".join(values))
    payload = "\n".join(rows)
    if not payload:
        return
    try:
        widget.clipboard_clear()
        widget.clipboard_append(payload)
        widget.update_idletasks()
    except Exception:
        pass



def create_undo_state(limit: int = 50) -> dict:
    try:
        limit_n = int(limit)
    except Exception:
        limit_n = 50
    if limit_n < 1:
        limit_n = 1
    return {"undo_stack": [], "undo_limit": limit_n}


def push_undo_rows(state: dict | None, tree: tk.Widget, rows: list[tuple[str, list[str]]], *, meta: dict | None = None) -> bool:
    if not isinstance(state, dict):
        return False

    cleaned: list[tuple[str, list[str]]] = []
    seen: set[str] = set()
    for iid, values in rows or []:
        iid_txt = str(iid or "")
        if not iid_txt or iid_txt in seen:
            continue
        seen.add(iid_txt)
        cleaned.append((iid_txt, ["" if v is None else str(v) for v in (values or [])]))

    if not cleaned:
        return False

    snapshot = {
        "rows": cleaned,
        "focus": "",
        "selection": [],
        "meta": dict(meta or {}),
    }
    try:
        snapshot["focus"] = str(tree.focus() or "")
    except Exception:
        pass
    try:
        snapshot["selection"] = [str(iid) for iid in tree.selection()]
    except Exception:
        pass

    stack = state.setdefault("undo_stack", [])
    if stack:
        prev = stack[-1]
        if prev.get("rows") == snapshot["rows"] and prev.get("meta") == snapshot["meta"]:
            return False

    stack.append(snapshot)
    try:
        limit = int(state.get("undo_limit", 50) or 50)
    except Exception:
        limit = 50
    if limit < 1:
        limit = 1
    if len(stack) > limit:
        del stack[:-limit]
    return True


def pop_undo_snapshot(state: dict | None) -> dict | None:
    if not isinstance(state, dict):
        return None
    stack = state.get("undo_stack")
    if not isinstance(stack, list) or not stack:
        return None
    try:
        return stack.pop()
    except Exception:
        return None


def restore_undo_snapshot(tree: tk.Widget, snapshot: dict | None) -> list[str]:
    if not isinstance(snapshot, dict):
        return []

    restored: list[str] = []
    for iid, values in snapshot.get("rows", []) or []:
        iid_txt = str(iid or "")
        if not iid_txt:
            continue
        try:
            tree.item(iid_txt, values=list(values or []))
            restored.append(iid_txt)
        except Exception:
            continue

    selection = [str(iid) for iid in snapshot.get("selection", []) or []]
    if selection:
        try:
            tree.selection_set(selection)
        except Exception:
            pass
    focus = str(snapshot.get("focus", "") or "")
    if focus:
        try:
            tree.focus(focus)
        except Exception:
            pass
    return restored
