"""CheckItems (non-Ex) tab: table detection test + manual checkmarks."""

from __future__ import annotations

import json
import queue
import re
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
import tkinter as tk

import fitz

from itr_modules.shared.paths import OUTPUT_ROOT, ensure_output_dir, get_batch_id, open_in_file_explorer
from itr_modules.shared.pdf_utils import (
    detect_checkitems_table,
    extract_rulings,
    find_cell_by_exact_norm,
    get_cell_text_cached,
    norm_text,
    parse_pages_per_itr_regex,
    row_band_from_ys,
)

MODULE_NAME = "check_items"
DEFAULT_PAGE1_REGEX = r"Page\s*1\s*of\s*(\d+)"
DEFAULT_TAG_REGEX = r"TAG\s*NO\.\s*:\s*([A-Za-z0-9\-\._/]+)"
TAG_DIRECTIONS = ["AUTO", "RIGHT", "DOWN"]


def _parse_norm_list(text: str) -> list[str]:
    return [norm_text(t) for t in (text or "").split(",") if norm_text(t)]


def _safe_int(text: str) -> int | None:
    try:
        return int(text)
    except Exception:
        return None


def _unique_sorted(vals: list[float], tol: float = 0.6) -> list[float]:
    vals = sorted(vals)
    out = []
    for v in vals:
        if not out or abs(v - out[-1]) > tol:
            out.append(v)
    return out


def _build_word_buckets(words: list[tuple], bucket_size: float = 8.0) -> dict[int, list[tuple]]:
    buckets: dict[int, list[tuple]] = {}
    if bucket_size <= 0:
        bucket_size = 8.0
    for w in words:
        y0 = w[1]
        y1 = w[3]
        b0 = int(y0 // bucket_size)
        b1 = int(y1 // bucket_size)
        for b in range(b0, b1 + 1):
            buckets.setdefault(b, []).append(w)
    return buckets


def _row_index_from_ys(ys: list[float], y_center: float) -> int:
    for i in range(len(ys) - 1):
        if ys[i] - 1 <= y_center <= ys[i + 1] + 1:
            return i
    return -1


def _col_bounds_from_xs(xs: list[float], x_center: float) -> tuple[int, float, float] | None:
    if not xs or len(xs) < 2:
        return None
    for i in range(len(xs) - 1):
        if xs[i] - 1 <= x_center <= xs[i + 1] + 1:
            return i, xs[i], xs[i + 1]
    return None


class CheckItemsTestTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.pdf_paths: list[str] = []
        self.header_norms_var = tk.StringVar(value="ITEM,DESCRIPTION,OK,NA,PL")
        self.index_col_norm_var = tk.StringVar(value="ITEM")
        self.state_col_norms_var = tk.StringVar(value="OK,NA,PL")
        self.page1_regex_var = tk.StringVar(value=DEFAULT_PAGE1_REGEX)
        self.pages_per_itr_var = tk.StringVar(value="")
        self.matchkey_name_var = tk.StringVar(value="TAG")
        self.tag_regex_var = tk.StringVar(value=DEFAULT_TAG_REGEX)
        self.tag_dir_var = tk.StringVar(value="AUTO")

        self._worker_thread: threading.Thread | None = None
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._state_path = OUTPUT_ROOT / MODULE_NAME / ".state" / "session_state.json"

        self.parsed_tags: dict[str, list[dict]] = {}
        self.selections: dict[str, dict[int, dict[int, str]]] = {}
        self.current_tag: str | None = None
        self.current_itr_index: int | None = None

        self._build_ui()
        self._load_state()

    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)

        ttk.Button(top, text="批量导入 PDF", command=self.pick_pdfs).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="解析", command=self.parse_pdfs).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="Test（生成框图 PDF）", command=self.run_test).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(top, text="打开测试输出", command=self.open_test_folder).pack(side=tk.LEFT, padx=(0, 8))
        self.status = ttk.Label(top, text="已导入PDF：0（可多选） | 选中：0")
        self.status.pack(side=tk.LEFT, padx=(12, 0))

        cfg = ttk.LabelFrame(self, text="测试配置", padding=10)
        cfg.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Label(cfg, text="HEADER_NORMS：").grid(row=0, column=0, sticky="e")
        ttk.Entry(cfg, textvariable=self.header_norms_var, width=40).grid(row=0, column=1, sticky="we", padx=6)

        ttk.Label(cfg, text="INDEX_COL_NORM：").grid(row=1, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.index_col_norm_var, width=40).grid(
            row=1, column=1, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="STATE_COL_NORMS：").grid(row=2, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.state_col_norms_var, width=40).grid(
            row=2, column=1, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="Page1 识别正则：").grid(row=0, column=2, sticky="e")
        ttk.Entry(cfg, textvariable=self.page1_regex_var, width=40).grid(row=0, column=3, sticky="we", padx=6)

        ttk.Label(cfg, text="每个 ITR 页数：").grid(row=1, column=2, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.pages_per_itr_var, width=40).grid(
            row=1, column=3, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="MatchKey 名称：").grid(row=2, column=2, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.matchkey_name_var, width=20).grid(
            row=2, column=3, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="Tag 正则：").grid(row=3, column=0, sticky="e", pady=(8, 0))
        ttk.Entry(cfg, textvariable=self.tag_regex_var, width=40).grid(
            row=3, column=1, columnspan=3, sticky="we", padx=6, pady=(8, 0)
        )

        ttk.Label(cfg, text="Tag 值方向：").grid(row=2, column=4, sticky="e", pady=(8, 0))
        ttk.OptionMenu(cfg, self.tag_dir_var, self.tag_dir_var.get(), *TAG_DIRECTIONS).grid(
            row=2, column=5, sticky="w", padx=6, pady=(8, 0)
        )

        cfg.columnconfigure(1, weight=1)
        cfg.columnconfigure(3, weight=1)

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

        self.tag_list = tk.Listbox(tag_frame, height=8)
        self.tag_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tag_sb = ttk.Scrollbar(tag_frame, orient="vertical", command=self.tag_list.yview)
        tag_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tag_list.configure(yscrollcommand=tag_sb.set)
        self.tag_list.bind("<<ListboxSelect>>", self._on_tag_select)

        tree_frame = ttk.Frame(grid_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        self.mark_tree = ttk.Treeview(tree_frame, columns=("OK", "NA", "PL"), show="tree headings", height=12)
        self.mark_tree.heading("#0", text="Item")
        self.mark_tree.heading("OK", text="OK")
        self.mark_tree.heading("NA", text="NA")
        self.mark_tree.heading("PL", text="PL")
        self.mark_tree.column("#0", width=60, anchor="center")
        self.mark_tree.column("OK", width=80, anchor="center")
        self.mark_tree.column("NA", width=80, anchor="center")
        self.mark_tree.column("PL", width=80, anchor="center")
        self.mark_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mark_sb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.mark_tree.yview)
        mark_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.mark_tree.configure(yscrollcommand=mark_sb.set)
        self.mark_tree.bind("<Button-1>", self._on_mark_click)

        self._set_parsed_ready(False)

    def _set_parsed_ready(self, ready: bool) -> None:
        state = tk.NORMAL if ready else tk.DISABLED
        self.tag_list.configure(state=state)
        if not ready:
            self.tag_list.delete(0, tk.END)
            self.mark_tree.delete(*self.mark_tree.get_children())
            self.current_tag = None
            self.current_itr_index = None

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
        if not isinstance(payload, dict):
            return
        selections = payload.get("selections", {})
        if isinstance(selections, dict):
            parsed = {}
            for tag, itr_map in selections.items():
                if not isinstance(itr_map, dict):
                    continue
                tag_map: dict[int, dict[int, str]] = {}
                for itr_key, items in itr_map.items():
                    try:
                        itr_idx = int(itr_key)
                    except Exception:
                        continue
                    if not isinstance(items, dict):
                        continue
                    item_map: dict[int, str] = {}
                    for item_no, val in items.items():
                        try:
                            item_idx = int(item_no)
                        except Exception:
                            continue
                        item_map[item_idx] = str(val)
                    tag_map[itr_idx] = item_map
                parsed[tag] = tag_map
            self.selections = parsed

    def _save_state(self) -> None:
        path = self._state_file()
        payload = {
            "selections": {
                tag: {str(itr_idx): {str(k): v for k, v in items.items()} for itr_idx, items in itr_map.items()}
                for tag, itr_map in self.selections.items()
            }
        }
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def _reset_parsed(self) -> None:
        self.parsed_tags = {}
        self._set_parsed_ready(False)

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
        self._reset_parsed()
        for p in self.pdf_paths:
            name = Path(p).name
            self.lst_pdfs.insert(tk.END, name)
        self._update_status()
        self._log(f"已导入 PDF：{len(self.pdf_paths)} 个")

    def open_test_folder(self):
        open_in_file_explorer(self._module_root() / "test")

    def _module_root(self) -> Path:
        path = OUTPUT_ROOT / MODULE_NAME
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _find_itr_segments(self, doc: fitz.Document, page1_regex: str, pages_per_itr: int | None) -> list[tuple[int, int]]:
        starts = []
        try:
            page1_re = re.compile(page1_regex, re.IGNORECASE)
        except re.error:
            page1_re = re.compile(DEFAULT_PAGE1_REGEX, re.IGNORECASE)
            self._log("[解析] Page1 正则非法，已回退默认规则。")

        for i in range(doc.page_count):
            text = doc[i].get_text("text") or ""
            if page1_re.search(text):
                starts.append(i)

        if starts:
            segments = []
            for idx, start in enumerate(starts):
                end = (starts[idx + 1] - 1) if idx + 1 < len(starts) else doc.page_count - 1
                segments.append((start, end))
            return segments

        if pages_per_itr:
            segments = []
            for start in range(0, doc.page_count, pages_per_itr):
                end = min(start + pages_per_itr - 1, doc.page_count - 1)
                segments.append((start, end))
            return segments

        return []

    def _find_tag_cells(
        self,
        page: fitz.Page,
        matchkey_norm: str,
        tag_regex: str,
        direction: str,
    ) -> tuple[fitz.Rect | None, fitz.Rect | None, str]:
        words = page.get_text("words") or []
        buckets = _build_word_buckets(words)
        cache: dict[tuple[float, float, float, float], str] = {}
        verticals, horizontals = extract_rulings(page)
        xs = _unique_sorted([x for x, _, _ in verticals])
        ys = _unique_sorted([y for y, _, _ in horizontals])

        match_cell = find_cell_by_exact_norm(
            page,
            matchkey_norm,
            verticals,
            horizontals,
            search_clip=None,
            words=words,
            buckets=buckets,
            cache=cache,
        )
        if not match_cell:
            return None, None, ""

        y_idx = _row_index_from_ys(ys, (match_cell.y0 + match_cell.y1) / 2.0)
        row_band = row_band_from_ys(y_idx, ys) if y_idx >= 0 else None
        col_info = _col_bounds_from_xs(xs, (match_cell.x0 + match_cell.x1) / 2.0)
        if not row_band or not col_info:
            return match_cell, None, ""

        col_idx, x0, x1 = col_info
        value_cell = None
        if direction in ("AUTO", "RIGHT") and col_idx + 2 < len(xs):
            value_cell = fitz.Rect(xs[col_idx + 1], row_band[0], xs[col_idx + 2], row_band[1])
        if direction == "DOWN" or (direction == "AUTO" and value_cell is None):
            if y_idx + 2 < len(ys):
                value_cell = fitz.Rect(x0, ys[y_idx + 1], x1, ys[y_idx + 2])

        tag_value = ""
        if value_cell:
            raw = get_cell_text_cached(value_cell, words, buckets, cache)
            try:
                rx = re.compile(tag_regex, re.IGNORECASE)
            except re.error:
                rx = re.compile(DEFAULT_TAG_REGEX, re.IGNORECASE)
            match = rx.search(raw)
            tag_value = match.group(1).strip() if match else raw.strip()

        return match_cell, value_cell, tag_value

    def _detect_table_info(self, doc: fitz.Document, segment: tuple[int, int]) -> tuple[int, dict]:
        header_norms = _parse_norm_list(self.header_norms_var.get())
        index_norm = norm_text(self.index_col_norm_var.get())
        state_norms = _parse_norm_list(self.state_col_norms_var.get())
        for page_index in range(segment[0], segment[1] + 1):
            info = detect_checkitems_table(doc[page_index], header_norms, index_norm, state_norms)
            count = len(info.get("numbered_rows") or [])
            if count:
                return count, info
        return 0, {}

    def parse_pdfs(self) -> None:
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return

        page1_regex = self.page1_regex_var.get().strip() or DEFAULT_PAGE1_REGEX
        pages_per_itr_manual = _safe_int(self.pages_per_itr_var.get().strip())
        tag_regex = self.tag_regex_var.get().strip() or DEFAULT_TAG_REGEX
        matchkey_name = self.matchkey_name_var.get().strip() or "TAG"
        matchkey_norm = norm_text(matchkey_name)
        direction = (self.tag_dir_var.get() or "AUTO").upper()

        self.parsed_tags = {}
        unknown_idx = 1

        for pdf_path in self.pdf_paths:
            try:
                doc = fitz.open(pdf_path)
            except Exception as exc:
                self._log(f"[解析失败] 无法打开 PDF：{Path(pdf_path).name} ({exc})")
                continue

            pages_per_itr_auto = parse_pages_per_itr_regex(doc, page1_regex, scan_pages=min(4, doc.page_count))
            if pages_per_itr_auto:
                self._log(f"[解析] {Path(pdf_path).name} 自动识别每套页数：{pages_per_itr_auto}")
            else:
                if pages_per_itr_manual:
                    self._log(f"[解析] {Path(pdf_path).name} 自动识别失败，使用手动页数：{pages_per_itr_manual}")
                else:
                    self._log(f"[解析失败] {Path(pdf_path).name} 无法识别每套页数，请填写手动页数。")

            segments = self._find_itr_segments(
                doc,
                page1_regex,
                pages_per_itr_auto or pages_per_itr_manual,
            )
            if not segments:
                self._log(f"[解析失败] {Path(pdf_path).name} 未能拆分 ITR 页段。")
                doc.close()
                continue

            for itr_idx, segment in enumerate(segments, start=1):
                tag_value = ""
                match_cell = None
                value_cell = None
                for page_index in range(segment[0], min(segment[1] + 1, segment[0] + 2)):
                    match_cell, value_cell, tag_value = self._find_tag_cells(
                        doc[page_index],
                        matchkey_norm,
                        tag_regex,
                        direction,
                    )
                    if tag_value:
                        break

                if not tag_value:
                    tag_value = f"UNKNOWN-{unknown_idx}"
                    unknown_idx += 1
                    self._log(
                        f"[解析] {Path(pdf_path).name} 第{itr_idx}套未找到{matchkey_name}，使用 {tag_value}"
                    )

                item_count, table_info = self._detect_table_info(doc, segment)
                if item_count == 0:
                    self._log(f"[解析] {Path(pdf_path).name} 第{itr_idx}套未识别到序号行。")

                entry = {
                    "pdf_path": pdf_path,
                    "itr_index": itr_idx,
                    "segment": segment,
                    "item_count": item_count,
                    "table_info": table_info,
                    "header_texts": table_info.get("header_texts", {}) if table_info else {},
                    "tag_match_cell": match_cell,
                    "tag_value_cell": value_cell,
                }
                self.parsed_tags.setdefault(tag_value, []).append(entry)

            doc.close()

        if not self.parsed_tags:
            self._set_parsed_ready(False)
            messagebox.showwarning("提示", "未解析到任何 Tag，请检查配置或日志。")
            return

        self.tag_list.delete(0, tk.END)
        for tag in sorted(self.parsed_tags.keys()):
            self.tag_list.insert(tk.END, tag)
        self._set_parsed_ready(True)
        self._log("[解析完成] Tag 列表已生成。")

    def _on_tag_select(self, _event=None) -> None:
        if not self.tag_list.curselection():
            return
        idx = self.tag_list.curselection()[0]
        tag = self.tag_list.get(idx)
        if not tag:
            return
        self.current_tag = tag
        self.current_itr_index = 1
        self.mark_tree.delete(*self.mark_tree.get_children())
        self._render_marks()

    def _render_marks(self) -> None:
        self.mark_tree.delete(*self.mark_tree.get_children())
        if self.current_tag is None or self.current_itr_index is None:
            return
        itr_entries = self.parsed_tags.get(self.current_tag, [])
        if self.current_itr_index - 1 >= len(itr_entries):
            return
        entry = itr_entries[self.current_itr_index - 1]
        item_count = int(entry.get("item_count", 0))
        header_texts = entry.get("header_texts", {})
        self.mark_tree.heading("#0", text=header_texts.get("ITEM", "Item") or "Item")
        self.mark_tree.heading("OK", text=header_texts.get("OK", "OK") or "OK")
        self.mark_tree.heading("NA", text=header_texts.get("NA", "NA") or "NA")
        self.mark_tree.heading("PL", text=header_texts.get("PL", "PL") or "PL")
        marks = self.selections.get(self.current_tag, {}).get(self.current_itr_index, {})
        for i in range(1, item_count + 1):
            mark = marks.get(i, "")
            values = (
                "✓" if mark == "OK" else "",
                "✓" if mark == "NA" else "",
                "✓" if mark == "PL" else "",
            )
            self.mark_tree.insert("", tk.END, iid=str(i), text=str(i), values=values)

    def _on_mark_click(self, event) -> None:
        if self.current_tag is None or self.current_itr_index is None:
            return
        row_id = self.mark_tree.identify_row(event.y)
        col_id = self.mark_tree.identify_column(event.x)
        if not row_id or col_id == "#0":
            return
        col_map = {"#1": "OK", "#2": "NA", "#3": "PL"}
        col = col_map.get(col_id)
        if not col:
            return

        tag_map = self.selections.setdefault(self.current_tag, {})
        marks = tag_map.setdefault(self.current_itr_index, {})
        try:
            row_num = int(row_id)
        except Exception:
            return
        current = marks.get(row_num, "")
        if current == col:
            marks[row_num] = ""
        else:
            marks[row_num] = col
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
        ensure_output_dir(MODULE_NAME, "filled", batch_id)

        self._worker_thread = threading.Thread(
            target=self._test_worker,
            args=(
                targets,
                header_norms,
                index_norm,
                state_norms,
                self.matchkey_name_var.get().strip() or "TAG",
                self.tag_regex_var.get().strip() or DEFAULT_TAG_REGEX,
                (self.tag_dir_var.get() or "AUTO").upper(),
                out_dir,
            ),
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

    def _draw_tag_boxes(self, page: fitz.Page, matchkey_name: str, tag_regex: str, direction: str) -> None:
        match_cell, value_cell, _ = self._find_tag_cells(page, norm_text(matchkey_name), tag_regex, direction)
        if match_cell:
            page.draw_rect(match_cell, color=(0, 0.6, 0), width=1.2)
        if value_cell:
            page.draw_rect(value_cell, color=(0, 0.6, 0), width=1.2)

    def _test_worker(
        self,
        targets: list[str],
        header_norms: list[str],
        index_norm: str,
        state_norms: list[str],
        matchkey_name: str,
        tag_regex: str,
        direction: str,
        out_dir: Path,
    ):
        for pdf_path in targets:
            stem = Path(pdf_path).stem
            out_pdf = out_dir / f"{stem}_test.pdf"
            doc = fitz.open(pdf_path)

            for page in doc:
                info = detect_checkitems_table(page, header_norms, index_norm, state_norms)
                header_cells = info.get("header_cells", {})
                index_bounds = info.get("index_bounds")
                ys = info.get("grid_ys", [])
                numbered_rows = info.get("numbered_rows", [])
                state_bounds = info.get("state_bounds", {})

                if not (index_bounds and ys and numbered_rows):
                    self._log("[Test] 表格识别失败，跳过该页框选。")
                    self._draw_tag_boxes(page, matchkey_name, tag_regex, direction)
                    continue

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

                self._draw_tag_boxes(page, matchkey_name, tag_regex, direction)

            doc.save(out_pdf)
            doc.close()
            self._q.put(("log", f"[Test] 已生成：{out_pdf}"))

        self._q.put(("done", out_dir))
