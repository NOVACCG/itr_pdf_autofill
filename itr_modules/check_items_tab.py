"""CheckItems (non-Ex) tab: test table detection and manage manual checkmarks."""

from __future__ import annotations

import json
import queue
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import fitz

from itr_modules.shared.paths import OUTPUT_ROOT, ensure_output_dir, get_batch_id, open_in_file_explorer
from itr_modules.shared.pdf_utils import detect_checkitems_table, draw_checkmark, norm_text, row_band_from_ys

MODULE_NAME = "check_items"


def _parse_norm_list(text: str) -> list[str]:
    return [norm_text(t) for t in (text or "").split(",") if norm_text(t)]


class CheckItemsTestTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pdf_paths: list[str] = []
        self.header_norms_var = tk.StringVar(value="ITEM,DESCRIPTION,OK,NA,PL")
        self.index_col_norm_var = tk.StringVar(value="ITEM")
        self.state_col_norms_var = tk.StringVar(value="OK,NA,PL")
        self._worker_thread: threading.Thread | None = None
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._state_path = OUTPUT_ROOT / MODULE_NAME / ".state" / "session_state.json"
        self.state: dict[str, dict] = {}
        self.current_pdf: str | None = None
        self._build_ui()
        self._load_state()

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)

        ttk.Button(top, text="批量导入 PDF", command=self.pick_pdfs).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="Test（生成框图 PDF）", command=self.run_test).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="打开测试输出", command=self.open_test_folder).pack(side=tk.LEFT, padx=(0, 8))

        self.status = ttk.Label(self, text="已导入PDF：0（可多选） | 选中：0", padding=(10, 0))
        self.status.pack(fill=tk.X)

        cfg = ttk.LabelFrame(self, text="测试配置", padding=10)
        cfg.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(cfg, text="HEADER_NORMS：").grid(row=0, column=0, sticky="e")
        ttk.Entry(cfg, textvariable=self.header_norms_var, width=50).grid(row=0, column=1, sticky="we", padx=6)
        ttk.Label(
            cfg,
            text="多个归一值用英文逗号分隔，自动去空格，大小写不敏感",
        ).grid(row=1, column=1, sticky="w", padx=6)

        ttk.Label(cfg, text="INDEX_COL_NORM：").grid(row=2, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.index_col_norm_var, width=50).grid(
            row=2, column=1, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="STATE_COL_NORMS：").grid(row=3, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.state_col_norms_var, width=50).grid(
            row=3, column=1, sticky="we", padx=6, pady=(8, 0)
        )

        cfg.columnconfigure(1, weight=1)

        mid = ttk.Panedwindow(self, orient=tk.HORIZONTAL)
        mid.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        list_frame = ttk.LabelFrame(mid, text="已导入 PDF 列表（可多选；不选则默认全部）", padding=(8, 6))
        log_frame = ttk.LabelFrame(mid, text="运行日志", padding=(8, 6))
        mid.add(list_frame, weight=1)
        mid.add(log_frame, weight=2)

        self.lst_pdfs = tk.Listbox(list_frame, height=6, selectmode=tk.EXTENDED)
        self.lst_pdfs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sbx = ttk.Scrollbar(list_frame, orient="vertical", command=self.lst_pdfs.yview)
        sbx.pack(side=tk.RIGHT, fill=tk.Y)
        self.lst_pdfs.configure(yscrollcommand=sbx.set)
        self.lst_pdfs.bind("<<ListboxSelect>>", lambda _e: self._update_status())

        self.log = tk.Text(log_frame, height=10)
        self.log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_sb = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.configure(yscrollcommand=log_sb.set)

        bottom = ttk.LabelFrame(self, text="打勾操作", padding=(8, 6))
        bottom.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        bottom_paned = ttk.Panedwindow(bottom, orient=tk.HORIZONTAL)
        bottom_paned.pack(fill=tk.BOTH, expand=True)

        tag_frame = ttk.Frame(bottom_paned)
        grid_frame = ttk.Frame(bottom_paned)
        bottom_paned.add(tag_frame, weight=1)
        bottom_paned.add(grid_frame, weight=3)

        ttk.Label(tag_frame, text="Tag 列表：").pack(anchor="w")
        self.tag_list = tk.Listbox(tag_frame, height=8)
        self.tag_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tag_sb = ttk.Scrollbar(tag_frame, orient="vertical", command=self.tag_list.yview)
        tag_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tag_list.configure(yscrollcommand=tag_sb.set)
        self.tag_list.bind("<<ListboxSelect>>", self._on_tag_select)

        grid_top = ttk.Frame(grid_frame)
        grid_top.pack(fill=tk.X)
        ttk.Button(grid_top, text="导出打勾 PDF", command=self.export_filled).pack(side=tk.RIGHT)

        self.mark_tree = ttk.Treeview(grid_frame, columns=("OK", "NA", "PL"), show="tree headings", height=10)
        self.mark_tree.heading("#0", text="Item")
        self.mark_tree.heading("OK", text="OK")
        self.mark_tree.heading("NA", text="NA")
        self.mark_tree.heading("PL", text="PL")
        self.mark_tree.column("#0", width=60, anchor="center")
        self.mark_tree.column("OK", width=80, anchor="center")
        self.mark_tree.column("NA", width=80, anchor="center")
        self.mark_tree.column("PL", width=80, anchor="center")
        self.mark_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mark_sb = ttk.Scrollbar(grid_frame, orient="vertical", command=self.mark_tree.yview)
        mark_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.mark_tree.configure(yscrollcommand=mark_sb.set)
        self.mark_tree.bind("<Button-1>", self._on_mark_click)

    def _log(self, msg: str):
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)

    def _update_status(self):
        sel = self.lst_pdfs.curselection()
        self.status.config(
            text=f"已导入PDF：{len(self.pdf_paths)}（可多选） | 选中：{len(sel) if sel else 0 if self.pdf_paths else 0}"
        )

    def _get_selected_pdfs(self) -> list[str]:
        idxs = list(self.lst_pdfs.curselection())
        if not idxs:
            return []
        out = []
        for i in idxs:
            try:
                out.append(self.pdf_paths[i])
            except Exception:
                pass
        return out

    def _state_file(self) -> Path:
        self._state_path.parent.mkdir(parents=True, exist_ok=True)
        return self._state_path

    def _load_state(self) -> None:
        path = self._state_file()
        if not path.exists():
            return
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception:
            return
        if isinstance(payload, dict):
            self.state = payload

    def _save_state(self) -> None:
        path = self._state_file()
        serializable = {}
        for pdf_path, entry in self.state.items():
            if not isinstance(entry, dict):
                continue
            marks = entry.get("marks", {})
            serializable[pdf_path] = {
                "item_count": int(entry.get("item_count", 0)),
                "marks": {str(k): v for k, v in (marks or {}).items()},
            }
        path.write_text(json.dumps(serializable, ensure_ascii=False, indent=2), encoding="utf-8")

    def _ensure_state_for_pdf(self, pdf_path: str) -> None:
        entry = self.state.setdefault(pdf_path, {"item_count": 0, "marks": {}})
        if entry.get("item_count", 0) > 0:
            return

        header_norms = _parse_norm_list(self.header_norms_var.get())
        index_norm = norm_text(self.index_col_norm_var.get())
        state_norms = _parse_norm_list(self.state_col_norms_var.get())
        if not header_norms or not index_norm or not state_norms:
            return

        try:
            doc = fitz.open(pdf_path)
        except Exception:
            return

        for page in doc:
            info = detect_checkitems_table(page, header_norms, index_norm, state_norms)
            row_count = len(info.get("numbered_rows") or [])
            if row_count:
                entry["item_count"] = row_count
                break
        doc.close()
        self._save_state()

    def pick_pdfs(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个 PDF", filetypes=[("PDF", "*.pdf")])
        if not paths:
            return
        seen = set()
        new_list = []
        for p in paths:
            if p not in seen:
                seen.add(p)
                new_list.append(p)
        self.pdf_paths = list(new_list)
        self.lst_pdfs.delete(0, tk.END)
        self.tag_list.delete(0, tk.END)
        for p in self.pdf_paths:
            name = Path(p).name
            self.lst_pdfs.insert(tk.END, name)
            self.tag_list.insert(tk.END, name)
        self._update_status()
        self._log(f"已导入 PDF：{len(self.pdf_paths)} 个")

    def open_test_folder(self):
        open_in_file_explorer(self._module_root() / "test")

    def _module_root(self) -> Path:
        path = OUTPUT_ROOT / MODULE_NAME
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _on_tag_select(self, _event=None) -> None:
        sel = self.tag_list.curselection()
        if not sel:
            return
        idx = sel[0]
        if idx >= len(self.pdf_paths):
            return
        self.current_pdf = self.pdf_paths[idx]
        self._ensure_state_for_pdf(self.current_pdf)
        self._render_marks()

    def _render_marks(self) -> None:
        self.mark_tree.delete(*self.mark_tree.get_children())
        if not self.current_pdf:
            return
        entry = self.state.get(self.current_pdf, {})
        item_count = int(entry.get("item_count", 0))
        marks = entry.get("marks", {}) or {}
        for i in range(1, item_count + 1):
            mark = marks.get(str(i), "")
            values = (
                "✓" if mark == "OK" else "",
                "✓" if mark == "NA" else "",
                "✓" if mark == "PL" else "",
            )
            self.mark_tree.insert("", tk.END, iid=str(i), text=str(i), values=values)

    def _on_mark_click(self, event) -> None:
        if not self.current_pdf:
            return
        row_id = self.mark_tree.identify_row(event.y)
        col_id = self.mark_tree.identify_column(event.x)
        if not row_id or col_id == "#0":
            return
        col_map = {"#1": "OK", "#2": "NA", "#3": "PL"}
        col = col_map.get(col_id)
        if not col:
            return

        entry = self.state.setdefault(self.current_pdf, {"item_count": 0, "marks": {}})
        marks = entry.setdefault("marks", {})
        current = marks.get(str(row_id), "")
        if current == col:
            marks[str(row_id)] = ""
        else:
            marks[str(row_id)] = col
        self._save_state()
        self._render_marks()

    def run_test(self):
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return
        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("提示", "正在运行，请稍等…")
            return

        targets = self._get_selected_pdfs() or self.pdf_paths
        header_norms = _parse_norm_list(self.header_norms_var.get())
        index_norm = norm_text(self.index_col_norm_var.get())
        state_norms = _parse_norm_list(self.state_col_norms_var.get())

        if not header_norms or not index_norm or not state_norms:
            messagebox.showwarning("提示", "请检查 HEADER_NORMS / INDEX_COL_NORM / STATE_COL_NORMS 配置")
            return

        batch_id = get_batch_id()
        out_dir = ensure_output_dir(MODULE_NAME, "test", batch_id)

        self._worker_thread = threading.Thread(
            target=self._test_worker,
            args=(targets, header_norms, index_norm, state_norms, out_dir),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_queue)

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._log(msg[1])
                elif kind == "done":
                    out_dir = msg[1]
                    messagebox.showinfo("完成", f"测试PDF已生成到：\n{out_dir}")
                    return
        except queue.Empty:
            pass

        if self._worker_thread and self._worker_thread.is_alive():
            self.after(120, self._poll_queue)

    def _test_worker(
        self,
        targets: list[str],
        header_norms: list[str],
        index_norm: str,
        state_norms: list[str],
        out_dir: Path,
    ):
        for pdf_path in targets:
            stem = Path(pdf_path).stem
            out_pdf = out_dir / f"{stem}_test.pdf"
            doc = fitz.open(pdf_path)
            detected_rows = 0

            for page in doc:
                info = detect_checkitems_table(page, header_norms, index_norm, state_norms)
                header_cells = info.get("header_cells", {})
                index_bounds = info.get("index_bounds")
                ys = info.get("grid_ys", [])
                numbered_rows = info.get("numbered_rows", [])
                state_bounds = info.get("state_bounds", {})

                if not (index_bounds and ys and numbered_rows):
                    continue

                detected_rows = max(detected_rows, len(numbered_rows))

                for rect in header_cells.values():
                    page.draw_rect(rect, color=(0, 0, 1), width=1.2)

                for row_idx in numbered_rows:
                    band = row_band_from_ys(row_idx, ys)
                    if not band:
                        continue
                    y0, y1 = band
                    page.draw_rect(fitz.Rect(index_bounds[0], y0, index_bounds[1], y1), color=(0, 0, 1), width=1.2)

                for state_norm in state_norms:
                    state_bounds_rect = state_bounds.get(state_norm)
                    if not state_bounds_rect:
                        continue
                    for row_idx in numbered_rows:
                        band = row_band_from_ys(row_idx, ys)
                        if not band:
                            continue
                        y0, y1 = band
                        page.draw_rect(
                            fitz.Rect(state_bounds_rect[0], y0, state_bounds_rect[1], y1),
                            color=(0, 0, 1),
                            width=1.2,
                        )

            doc.save(out_pdf)
            doc.close()
            if detected_rows:
                entry = self.state.setdefault(pdf_path, {"item_count": 0, "marks": {}})
                entry["item_count"] = max(int(entry.get("item_count", 0)), detected_rows)
                self._save_state()
            self._q.put(("log", f"[Test] 已生成：{out_pdf}"))

        self._q.put(("done", out_dir))

    def export_filled(self) -> None:
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return

        header_norms = _parse_norm_list(self.header_norms_var.get())
        index_norm = norm_text(self.index_col_norm_var.get())
        state_norms = _parse_norm_list(self.state_col_norms_var.get())
        if not header_norms or not index_norm or not state_norms:
            messagebox.showwarning("提示", "请检查 HEADER_NORMS / INDEX_COL_NORM / STATE_COL_NORMS 配置")
            return

        batch_id = get_batch_id()
        out_dir = ensure_output_dir(MODULE_NAME, "filled", batch_id)
        for pdf_path in self.pdf_paths:
            marks = (self.state.get(pdf_path, {}) or {}).get("marks", {}) or {}
            if not marks:
                continue
            try:
                doc = fitz.open(pdf_path)
            except Exception:
                continue

            for page in doc:
                info = detect_checkitems_table(page, header_norms, index_norm, state_norms)
                ys = info.get("grid_ys", [])
                numbered_rows = info.get("numbered_rows", [])
                state_bounds = info.get("state_bounds", {})
                if not (ys and numbered_rows and state_bounds):
                    continue

                for idx, row_idx in enumerate(numbered_rows, start=1):
                    mark = marks.get(str(idx), "")
                    if not mark:
                        continue
                    bounds = state_bounds.get(mark)
                    if not bounds:
                        continue
                    band = row_band_from_ys(row_idx, ys)
                    if not band:
                        continue
                    y0, y1 = band
                    rect = fitz.Rect(bounds[0], y0, bounds[1], y1)
                    draw_checkmark(page, rect, width=1.6)

            out_pdf = out_dir / f"{Path(pdf_path).stem}_filled.pdf"
            doc.save(out_pdf)
            doc.close()
            self._log(f"[导出] {out_pdf}")

        messagebox.showinfo("完成", f"已导出到：\n{out_dir}")
