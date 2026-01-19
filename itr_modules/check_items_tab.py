"""CheckItems (non-Ex) test tab: locate tables and draw test boxes only."""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import fitz

from itr_modules.shared.paths import OUTPUT_ROOT, ensure_output_dir, get_batch_id, open_in_file_explorer
from itr_modules.shared.pdf_utils import (
    build_table_row_lines,
    extract_rulings,
    extract_table_grid_lines,
    find_cell_by_exact_norm,
    find_lowest_header_anchor,
    get_cell_text,
    header_row_band,
    is_pure_int,
    norm_text,
    row_band_from_ys,
    row_index_from_ys,
    snap_to_grid_x,
    split_columns_from_grid,
)


MODULE_NAME = "check_items"


def _parse_norm_list(text: str) -> list[str]:
    return [norm_text(t) for t in (text or "").split(",") if norm_text(t)]


def _row_bounds_from_lines(row_lines: list[float]) -> list[tuple[float, float]]:
    return [(y0, y1) for y0, y1 in zip(row_lines, row_lines[1:]) if y1 - y0 > 4]


def _row_bounds_from_words(words: list[tuple]) -> list[tuple[float, float]]:
    rows: dict[int, list[tuple[float, float]]] = {}
    for w in words:
        line_no = w[6]
        rows.setdefault(line_no, []).append((w[1], w[3]))
    bounds = []
    for ys in rows.values():
        y0 = min(y for y, _ in ys)
        y1 = max(y for _, y in ys)
        bounds.append((y0, y1))
    bounds.sort(key=lambda r: r[0])
    return bounds


def _table_rect_from_data(
    page: fitz.Page,
    verticals: list[tuple[float, float, float]],
    row_bounds: list[tuple[float, float]],
    words: list[tuple],
) -> fitz.Rect:
    if verticals:
        xs = [x for x, _, _ in verticals]
        x0 = min(xs)
        x1 = max(xs)
    elif words:
        x0 = min(w[0] for w in words)
        x1 = max(w[2] for w in words)
    else:
        x0 = page.rect.x0
        x1 = page.rect.x1

    if row_bounds:
        y0 = min(r[0] for r in row_bounds)
        y1 = max(r[1] for r in row_bounds)
    elif words:
        y0 = min(w[1] for w in words)
        y1 = max(w[3] for w in words)
    else:
        y0 = page.rect.y0
        y1 = page.rect.y1

    return fitz.Rect(x0, y0, x1, y1)


def _find_header_cells(
    page: fitz.Page,
    header_norms: list[str],
    verticals,
    horizontals,
    search_clip: fitz.Rect | None,
) -> dict[str, fitz.Rect]:
    cells: dict[str, fitz.Rect] = {}
    for norm in header_norms:
        rect = find_cell_by_exact_norm(page, norm, verticals, horizontals, search_clip=search_clip)
        if rect:
            cells[norm] = rect
    return cells


def _find_column_bounds(xs: list[float], header_rect: fitz.Rect | None) -> tuple[float, float] | None:
    if not header_rect:
        return None
    cx = (header_rect.x0 + header_rect.x1) / 2.0
    return snap_to_grid_x(cx, xs)


def _header_row_index(ys: list[float], header_rect: fitz.Rect | None) -> int:
    if not header_rect or not ys:
        return -1
    cy = (header_rect.y0 + header_rect.y1) / 2.0
    return row_index_from_ys(ys, cy)


class CheckItemsTestTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pdf_paths: list[str] = []
        self.header_norms_var = tk.StringVar(value="ITEM,DESCRIPTION,OK,NA,PL")
        self.index_col_norm_var = tk.StringVar(value="ITEM")
        self.state_col_norms_var = tk.StringVar(value="OK,NA,PL")
        self._worker_thread: threading.Thread | None = None
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._build_ui()

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
        ttk.Entry(cfg, textvariable=self.index_col_norm_var, width=50).grid(row=2, column=1, sticky="we", padx=6, pady=(8, 0))

        ttk.Label(cfg, text="STATE_COL_NORMS：").grid(row=3, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.state_col_norms_var, width=50).grid(row=3, column=1, sticky="we", padx=6, pady=(8, 0))

        cfg.columnconfigure(1, weight=1)

        list_frame = ttk.LabelFrame(self, text="已导入 PDF 列表（可多选；不选则默认全部）", padding=(8, 6))
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        self.lst_pdfs = tk.Listbox(list_frame, height=4, selectmode=tk.EXTENDED)
        self.lst_pdfs.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sbx = ttk.Scrollbar(list_frame, orient="vertical", command=self.lst_pdfs.yview)
        sbx.pack(side=tk.RIGHT, fill=tk.Y)
        self.lst_pdfs.configure(yscrollcommand=sbx.set)
        self.lst_pdfs.bind("<<ListboxSelect>>", lambda _e: self._update_status())

        ttk.Label(self, text="运行日志：").pack(anchor="w", padx=10)
        self.log = tk.Text(self, height=8)
        self.log.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

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
        for p in self.pdf_paths:
            self.lst_pdfs.insert(tk.END, Path(p).name)
        self._update_status()
        self._log(f"已导入 PDF：{len(self.pdf_paths)} 个")

    def open_test_folder(self):
        open_in_file_explorer(self._module_root() / "test")

    def _module_root(self) -> Path:
        path = OUTPUT_ROOT / MODULE_NAME
        path.mkdir(parents=True, exist_ok=True)
        return path

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

            for page in doc:
                verticals, horizontals = extract_rulings(page)
                words = page.get_text("words") or []

                header_anchor = find_lowest_header_anchor(page, header_norms, verticals, horizontals)
                header_band = header_row_band(header_anchor) if header_anchor else None

                x_left = min((x for x, _, _ in verticals), default=page.rect.x0)
                x_right = max((x for x, _, _ in verticals), default=page.rect.x1)

                row_lines = []
                if horizontals and x_left < x_right:
                    start_y = header_band.y0 if header_band else 0
                    row_lines = build_table_row_lines(page, horizontals, x_left, x_right, y_start=start_y)
                row_bounds = _row_bounds_from_lines(row_lines) if row_lines else _row_bounds_from_words(words)

                table_rect = _table_rect_from_data(page, verticals, row_bounds, words)
                xs, ys = extract_table_grid_lines(page, table_rect, verticals, horizontals)
                search_clip = header_band if header_band else table_rect
                header_cells = _find_header_cells(page, header_norms, verticals, horizontals, search_clip)
                if header_band is None and header_cells:
                    header_band = fitz.Rect(
                        table_rect.x0,
                        min(r.y0 for r in header_cells.values()) - 2,
                        table_rect.x1,
                        max(r.y1 for r in header_cells.values()) + 2,
                    )

                index_header = header_cells.get(index_norm) or header_anchor
                index_bounds = _find_column_bounds(xs, index_header)
                if not index_bounds:
                    continue

                header_row_idx = _header_row_index(ys, index_header)

                expected = 1
                numbered_rows: list[int] = []
                for i in range(header_row_idx + 1, len(ys) - 1):
                    band = row_band_from_ys(i, ys)
                    if not band:
                        continue
                    y0, y1 = band
                    cell = fitz.Rect(index_bounds[0], y0, index_bounds[1], y1)
                    cell_text = get_cell_text(page, cell)
                    if is_pure_int(cell_text) and int(cell_text) == expected:
                        numbered_rows.append(i)
                        expected += 1
                    elif expected > 1:
                        break

                if not numbered_rows:
                    continue

                for rect in header_cells.values():
                    page.draw_rect(rect, color=(0, 0, 1), width=1.2)

                for row_idx in numbered_rows:
                    band = row_band_from_ys(row_idx, ys)
                    if not band:
                        continue
                    y0, y1 = band
                    page.draw_rect(fitz.Rect(index_bounds[0], y0, index_bounds[1], y1), color=(0, 0, 1), width=1.2)

                state_bounds_map: dict[str, tuple[float, float]] = {}
                for state_norm in state_norms:
                    state_header = header_cells.get(state_norm)
                    bounds = _find_column_bounds(xs, state_header)
                    if bounds:
                        state_bounds_map[state_norm] = bounds

                if len(state_bounds_map) != len(state_norms):
                    if state_bounds_map:
                        x0 = min(b[0] for b in state_bounds_map.values())
                        x1 = max(b[1] for b in state_bounds_map.values())
                    else:
                        x0 = table_rect.x0
                        x1 = table_rect.x1
                    splits = split_columns_from_grid(xs, x0, x1, len(state_norms))
                    for state_norm, bounds in zip(state_norms, splits):
                        state_bounds_map.setdefault(state_norm, bounds)

                for state_norm in state_norms:
                    state_bounds = state_bounds_map.get(state_norm)
                    if not state_bounds:
                        continue
                    for row_idx in numbered_rows:
                        band = row_band_from_ys(row_idx, ys)
                        if not band:
                            continue
                        y0, y1 = band
                        page.draw_rect(
                            fitz.Rect(state_bounds[0], y0, state_bounds[1], y1),
                            color=(0, 0, 1),
                            width=1.2,
                        )

            doc.save(out_pdf)
            doc.close()
            self._q.put(("log", f"[Test] 已生成：{out_pdf}"))

        self._q.put(("done", out_dir))
