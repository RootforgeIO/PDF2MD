"""Application dialogs with translated buttons, independent of Windows language."""
from __future__ import annotations

import tkinter as tk
from tkinter import ttk

from i18n import tr


def _show(title, message, choices, *, parent=None, cancel=None):
    parent = parent or tk._default_root
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent)
    dialog.resizable(False, False)
    result = cancel
    scale = dialog.winfo_fpixels("1i") / 96
    body = ttk.Frame(dialog, padding=round(14 * scale))
    body.pack(fill="both", expand=True)
    ttk.Label(body, text=message, wraplength=round(440 * scale), justify="left").pack(fill="x")
    buttons = ttk.Frame(body)
    buttons.pack(anchor="e", pady=(round(14 * scale), 0))

    def choose(value):
        nonlocal result
        result = value
        dialog.destroy()

    for label, value in choices:
        ttk.Button(buttons, text=tr(label), command=lambda v=value: choose(v),
                   style="Luna.TButton").pack(side="left", padx=3)
    dialog.protocol("WM_DELETE_WINDOW", lambda: choose(cancel))
    dialog.bind("<Escape>", lambda _: choose(cancel))
    dialog.update_idletasks()
    if parent is not None:
        x = parent.winfo_rootx() + max(0, (parent.winfo_width() - dialog.winfo_reqwidth()) // 2)
        y = parent.winfo_rooty() + max(0, (parent.winfo_height() - dialog.winfo_reqheight()) // 2)
        dialog.geometry(f"+{x}+{y}")
    dialog.wait_visibility()
    previous = dialog.grab_current()
    dialog.grab_set()
    dialog.focus_set()
    dialog.wait_window()
    if previous is not None and previous.winfo_exists():
        previous.grab_set()
    return result


def askyesnocancel(title, message, **kwargs):
    return _show(title, message, [("Да", True), ("Нет", False), ("Отмена", None)], **kwargs)


def askyesno(title, message, **kwargs):
    return _show(title, message, [("Да", True), ("Нет", False)], cancel=False, **kwargs)


def askokcancel(title, message, **kwargs):
    return _show(title, message, [("OK", True), ("Отмена", False)], cancel=False, **kwargs)


def showerror(title, message, **kwargs):
    return _show(title, message, [("OK", None)], **kwargs)
