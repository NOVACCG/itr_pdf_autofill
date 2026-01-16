"""Shared UI helpers (placeholders for future styling/components)."""

from __future__ import annotations

from pathlib import Path
from tkinter import messagebox, ttk

from .paths import open_in_file_explorer


def open_folder(path: Path) -> None:
    open_in_file_explorer(path)


def show_error(title: str, message: str) -> None:
    messagebox.showerror(title, message)


def show_info(title: str, message: str) -> None:
    messagebox.showinfo(title, message)


def apply_global_font(root) -> None:
    style = ttk.Style(root)
    for family in ("Microsoft YaHei UI", "Segoe UI", "Microsoft YaHei"):
        if family in root.tk.call("font", "families"):
            base_font = (family, 10)
            break
    else:
        base_font = ("TkDefaultFont", 10)
    style.configure(".", font=base_font)
