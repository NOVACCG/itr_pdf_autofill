"""Shared path helpers for output/report folders."""

from __future__ import annotations

import os
import sys
from datetime import datetime
from pathlib import Path
from tkinter import messagebox

BASE_DIR = Path(sys.executable).parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[2]
OUTPUT_ROOT = BASE_DIR / "output"
REPORT_ROOT = BASE_DIR / "report"


def get_batch_id() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def ensure_output_dir(module_name: str, kind: str, batch_id: str) -> Path:
    path = OUTPUT_ROOT / module_name / kind / batch_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def ensure_report_dir(module_name: str, batch_id: str) -> Path:
    path = REPORT_ROOT / module_name / batch_id
    path.mkdir(parents=True, exist_ok=True)
    return path


def open_in_file_explorer(path: Path) -> None:
    path.mkdir(parents=True, exist_ok=True)
    try:
        os.startfile(path)  # type: ignore[attr-defined]
    except Exception as exc:  # pragma: no cover - OS specific
        messagebox.showerror("错误", f"无法打开文件夹: {exc}")