# -*- coding: utf-8 -*-
"""
NA 自动勾选模块（Tab 正式版 UI）- v3.5（支持批量导入PDF + NA自动打勾）
================================================

本版新增：
- ✅ 方案A：一次选择多个 PDF 批量导入（askopenfilenames）
- ✅ 必须先【解析】（对选中PDF或全部PDF），解析成功后【测试/打勾】才可用（强制流程）
- ✅ 测试输出目录：output/na_check/test/<batch>
- ✅ 测试PDF命名：原文件名 + "_test_boxes.pdf"
- ✅ “测试PDF显示标注字”开关：给每个框都写小字标注（可开关）

说明：
- 目前支持“解析 + 测试画框 + NA自动打勾（列格式化：用表头列边界 + 横线切行）”。
- Phase 2 说明：该文件仅作为技术路线参考，UI 将在后续版本移除。
"""

import csv
import datetime
import os
import queue
import re
import threading
import time
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path

import fitz  # PyMuPDF

from itr_modules.shared.paths import (
    BASE_DIR,
    OUTPUT_ROOT,
    REPORT_ROOT,
    ensure_output_dir,
    ensure_report_dir,
    get_batch_id,
    open_in_file_explorer,
)
from itr_modules.shared.pdf_utils import (
    build_table_row_lines,
    cell_rect_for_word,
    draw_checkmark,
    extract_rulings,
    find_cell_by_exact_norm,
    find_ex_concept_cells,
    find_lowest_header_anchor,
    find_ok_na_pl_cells,
    fit_text_to_box,
    get_cell_text,
    header_row_band,
    is_pure_int,
    norm_text,
    parse_pages_per_itr_regex,
    rect_between_lines,
    _cell_text_from_row_words,
    _snap_col_bounds,
    _unique_sorted_x_from_verticals,
)

# -------------------------
# 输出目录（统一 output/，按模块/输出类型/批次）
# output/
#   na_check/
#     filled/<batch>/
#     test/<batch>/
#
# 报告目录（统一 report/，按模块/批次）
# report/
#   na_check/
#     <batch>/
# -------------------------
MODULE_NAME = "na_check"
REPORT_MODULE_NAME = "na_check"


def _label_rect_above(cell: fitz.Rect, height: float = 8.0) -> fitz.Rect:
    return fitz.Rect(cell.x0, cell.y0 - height, cell.x1, cell.y0)


def open_folder(path: Path):
    open_in_file_explorer(path)


class NACheckTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        # 批量PDF：列表
        self.pdf_paths: list[str] = []

        # 解析结果每个PDF一份
        self.parsed_map: dict[str, list[dict]] = {}
        self.pages_per_itr_map: dict[str, int] = {}

        # 页数模式：正则优先 + 手动兜底
        self.page_mode = tk.StringVar(value="regex")
        self.page_regex = tk.StringVar(value=r"Page\s*\d+\s*of\s*(\d+)")
        self.page_scan_n = tk.StringVar(value="5")
        self.page_manual = tk.StringVar(value="4")

        # 底表序号表头候选
        self.num_header_var = tk.StringVar(value="NO,NUMBER,ITM,ITEM")

        # 流程状态：必须解析成功，才能测试/打勾
        self.is_parsed_ok = False  # 全局：当前“选中范围”是否已解析

        # 调试标注开关：在测试PDF里给每个框写一个小字
        self.debug_labels = tk.BooleanVar(value=True)

        self._worker_thread: threading.Thread | None = None
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._tick_total_pages: int = 0
        self._tick_start_time: float | None = None

        self._build_ui()
        self._apply_state()

    # ---------- UI ----------
    def _build_ui(self):
        top = ttk.Frame(self, padding=(10, 8))
        top.pack(fill=tk.X)

        self.btn_pick = ttk.Button(top, text="批量导入 PDF", command=self.pick_pdfs)
        self.btn_pick.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_parse = ttk.Button(top, text="解析（抓锚点）", command=self.parse_selected)
        self.btn_parse.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_test = ttk.Button(top, text="测试（生成框图 PDF）", command=self.test_selected)
        self.btn_test.pack(side=tk.LEFT, padx=(0, 8))

        self.btn_tick = ttk.Button(top, text="打勾（NA）", command=self.tick_na_placeholder)
        self.btn_tick.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Button(top, text="打开导出文件夹", command=self.open_filled_folder).pack(side=tk.LEFT, padx=(0, 8))

        ttk.Separator(top, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)

        ttk.Button(top, text="打开测试输出 test", command=lambda: open_folder(self._module_root() / "test")).pack(
            side=tk.LEFT, padx=(0, 8)
        )

        ttk.Checkbutton(top, text="测试PDF显示标注字", variable=self.debug_labels).pack(side=tk.LEFT, padx=(8, 0))

        self.status = ttk.Label(self, text="已导入PDF：0（可多选） | 选中：0", padding=(10, 0))
        self.status.pack(fill=tk.X)

        self.progress_text = ttk.Label(self, text="进度：未开始", padding=(10, 4))
        self.progress_text.pack(fill=tk.X)
        style = ttk.Style()
        style.configure("Green.Horizontal.TProgressbar", troughcolor="#e6e6e6", background="#4caf50")
        self.progress = ttk.Progressbar(self, mode="determinate", style="Green.Horizontal.TProgressbar")
        self.progress.pack(fill=tk.X, padx=10, pady=(0, 6))

        paned = ttk.Panedwindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        upper = ttk.Frame(paned)
        lower = ttk.Frame(paned)
        paned.add(upper, weight=3)
        paned.add(lower, weight=2)

        style = ttk.Style()
        style.configure("TPanedwindow", sashrelief="raised")

        # --- PDF列表 + 配置（上半部分） ---
        upper_top = ttk.Frame(upper)
        upper_top.pack(fill=tk.BOTH, expand=False)

        lf_pdfs = ttk.LabelFrame(upper_top, text="已导入 PDF 列表（可多选；不选则默认全部）", padding=(8, 6))
        lf_pdfs.pack(fill=tk.X, pady=(0, 10))

        self.lst_pdfs = tk.Listbox(lf_pdfs, height=4, selectmode=tk.EXTENDED)
        self.lst_pdfs.pack(side=tk.LEFT, fill=tk.X, expand=True)
        sbx = ttk.Scrollbar(lf_pdfs, orient="vertical", command=self.lst_pdfs.yview)
        sbx.pack(side=tk.RIGHT, fill=tk.Y)
        self.lst_pdfs.configure(yscrollcommand=sbx.set)
        self.lst_pdfs.bind("<<ListboxSelect>>", lambda e: self._on_select_changed())

        cfg = ttk.LabelFrame(upper, text="配置（页数：正则优先 + 手动兜底）", padding=10)
        cfg.pack(fill=tk.X, pady=(0, 10))

        r = 0
        ttk.Radiobutton(cfg, text="正则方式（优先）", variable=self.page_mode, value="regex").grid(
            row=r, column=0, sticky="w"
        )
        ttk.Label(cfg, text="正则：").grid(row=r, column=1, sticky="e")
        ttk.Entry(cfg, textvariable=self.page_regex, width=36).grid(row=r, column=2, sticky="we", padx=5)
        ttk.Label(cfg, text="扫描前 N 页：").grid(row=r, column=3, sticky="e")
        ttk.Entry(cfg, textvariable=self.page_scan_n, width=6).grid(row=r, column=4, sticky="w")
        r += 1

        ttk.Radiobutton(cfg, text="手动方式", variable=self.page_mode, value="manual").grid(
            row=r, column=0, sticky="w", pady=(6, 0)
        )
        ttk.Label(cfg, text="每套页数：").grid(row=r, column=1, sticky="e", pady=(6, 0))
        ttk.Entry(cfg, textvariable=self.page_manual, width=8).grid(row=r, column=2, sticky="w", padx=5, pady=(6, 0))

        r += 1
        ttk.Label(cfg, text="底部矩阵表序号表头候选（逗号分隔）：").grid(
            row=r, column=0, sticky="w", pady=(10, 0), columnspan=2
        )
        ttk.Entry(cfg, textvariable=self.num_header_var).grid(
            row=r, column=2, sticky="we", padx=5, pady=(10, 0), columnspan=3
        )

        cfg.columnconfigure(2, weight=1)

        # --- 解析概览列表 ---
        list_frame = ttk.LabelFrame(upper, text="解析概览（按 ITR）", padding=(8, 6))
        list_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("pdf", "itr", "page", "no", "desc", "excols", "okna", "exconcept", "exvalue")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        for k, title, w, anchor in [
            ("pdf", "PDF", 220, "w"),
            ("itr", "ITR#", 60, "center"),
            ("page", "页(起)", 60, "center"),
            ("no", "NO表头", 70, "center"),
            ("desc", "Description", 90, "center"),
            ("excols", "EX列数", 70, "center"),
            ("okna", "OK/NA/PL", 90, "center"),
            ("exconcept", "Ex Concept", 90, "center"),
            ("exvalue", "Ex值格", 80, "center"),
        ]:
            self.tree.heading(k, text=title)
            self.tree.column(k, width=w, anchor=anchor)

        ysb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=ysb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        ysb.pack(side=tk.RIGHT, fill=tk.Y)

        ttk.Label(lower, text="运行日志：").pack(anchor="w")
        self.log = tk.Text(lower, height=10)
        self.log.pack(fill=tk.BOTH, expand=True)

        self._log(f"输出根目录：{OUTPUT_ROOT}")
        self._log(f"模块目录：{self._module_root()}")
        self._log(f"测试输出目录：{self._module_root() / 'test'}")
        self._log("流程：批量导入PDF → 解析（必须）→ 测试（看框是否正确）→ 打勾（下一步实现）")

    # ---------- 状态控制 ----------
    def _apply_state(self):
        has_pdf = len(self.pdf_paths) > 0
        self.btn_parse.configure(state=("normal" if has_pdf else "disabled"))

        # test/tick：必须对“当前选中范围”解析成功
        self.btn_test.configure(state=("normal" if self.is_parsed_ok else "disabled"))
        self.btn_tick.configure(state=("normal" if self.is_parsed_ok else "disabled"))

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._log(msg[1])
                elif kind == "parse_done":
                    parsed_map, pages_map, parsed_pdf_count, total_itr_ok = msg[1:]
                    self.parsed_map = parsed_map
                    self.pages_per_itr_map = pages_map

                    for item in self.tree.get_children():
                        self.tree.delete(item)

                    for pdf_path, parsed_list in self.parsed_map.items():
                        pages_per_itr = self.pages_per_itr_map.get(pdf_path, 4)
                        for entry in parsed_list:
                            i = entry["itr"]
                            pidx = entry["page_index"]
                            has_no = bool(entry.get("no_cell"))
                            self.tree.insert(
                                "",
                                "end",
                                values=(
                                    Path(pdf_path).name,
                                    i,
                                    pidx + 1,
                                    "✅" if has_no else "❌",
                                    "✅" if entry.get("desc_cell") else "❌",
                                    len(entry.get("ex_cells") or []),
                                    ",".join(sorted((entry.get("okna") or {}).keys())) if (entry.get("okna") or {}) else "❌",
                                    "✅" if entry.get("ex_label") else "❌",
                                    "✅" if entry.get("ex_value") else "❌",
                                ),
                            )

                    self._recompute_parsed_ok_for_selection()
                    self._update_status()
                    self.status.config(text=f"解析完成：PDF={parsed_pdf_count} | ITR={total_itr_ok}")
                    self._apply_state()
                    self._enable_actions()

                    if parsed_pdf_count == 0:
                        messagebox.showwarning(
                            "解析失败",
                            "选中范围内没有任何 PDF 解析成功。\n请检查页数配置/NO候选或PDF结构。",
                        )
                    else:
                        messagebox.showinfo(
                            "解析完成",
                            f"解析完成：PDF={parsed_pdf_count} 个，命中 ITR={total_itr_ok} 套。\n现在可点【测试】。",
                        )
                    return
                elif kind == "ex_value_request":
                    pdf_name, itr_idx, box, event = msg[1:]
                    result = self._prompt_missing_ex_value(pdf_name, itr_idx)
                    box["status"] = result["status"]
                    box["value"] = result["value"]
                    event.set()
                elif kind == "test_done":
                    out_dir = msg[1]
                    self.status.config(text="测试完成")
                    self._enable_actions()
                    messagebox.showinfo("完成", f"测试PDF已生成到：\n{out_dir}\n（按原文件名 + _test_boxes.pdf 命名）")
                    return
                elif kind == "tick_done":
                    pdf_done, checked_total, out_dir, report_path, skipped_pages = msg[1:]
                    self.status.config(text="打勾完成")
                    if self._tick_total_pages:
                        self.progress["maximum"] = self._tick_total_pages
                        self.progress["value"] = self._tick_total_pages
                        self.progress_text.config(
                            text=f"已处理 {self._tick_total_pages} / {self._tick_total_pages} 页 | ETA 0秒 | 当前：完成"
                        )
                    self._enable_actions()
                    if skipped_pages:
                        messagebox.showwarning(
                            "完成（有跳过页）",
                            f"已处理 {pdf_done} 个PDF。\n共勾选 NA：{checked_total} 处。\n"
                            f"跳过页数：{len(skipped_pages)}。\n"
                            f"报告：{report_path}",
                        )
                    else:
                        messagebox.showinfo(
                            "完成",
                            f"已处理 {pdf_done} 个PDF。\n共勾选 NA：{checked_total} 处。\n输出目录：{out_dir}",
                        )
                    return
                elif kind == "progress":
                    done_pages, total_pages, pdf_name, page_num = msg[1:]
                    self.progress["maximum"] = max(total_pages, 1)
                    self.progress["value"] = done_pages
                    eta_text = "ETA --"
                    if self._tick_start_time and done_pages > 0:
                        elapsed = max(time.time() - self._tick_start_time, 0.0)
                        remaining = max(total_pages - done_pages, 0)
                        eta_seconds = int(remaining * (elapsed / done_pages))
                        eta_text = f"ETA {eta_seconds // 60}分{eta_seconds % 60}秒"
                    self.progress_text.config(
                        text=f"已处理 {done_pages} / {total_pages} 页 | {eta_text} | 当前：{pdf_name} 第 {page_num} 页"
                    )
        except queue.Empty:
            pass

        if self._worker_thread and self._worker_thread.is_alive():
            self.after(120, self._poll_queue)
        else:
            self._enable_actions()

    def _update_status(self):
        sel = self._get_selected_pdfs()
        self.status.config(
            text=f"已导入PDF：{len(self.pdf_paths)}（可多选） | 选中：{len(sel) if sel else 0 if self.pdf_paths else 0}"
            f"{'（默认全部）' if (self.pdf_paths and not sel) else ''}"
        )

    def _on_select_changed(self):
        # 选择变化后：如果“选中范围”里有未解析的，就锁住测试/打勾
        self._recompute_parsed_ok_for_selection()
        self._update_status()
        self._apply_state()

    def _recompute_parsed_ok_for_selection(self):
        if not self.pdf_paths:
            self.is_parsed_ok = False
            return
        sel = self._get_selected_pdfs()
        targets = sel if sel else self.pdf_paths  # 不选则默认全部
        self.is_parsed_ok = all((p in self.parsed_map and self.parsed_map[p]) for p in targets)

    # ---------- 日志 ----------
    def _log(self, s: str):
        self.log.insert(tk.END, s + "\n")
        self.log.see(tk.END)

    def _prompt_missing_ex_value(self, pdf_name: str, itr_idx: int) -> dict:
        win = tk.Toplevel(self)
        win.title("需要手动输入防爆类型")
        win.geometry("720x220")
        win.resizable(False, False)

        msg = (
            "未检测到防爆类型值格（ExConcept 值格）。请输入本套 ITR 的防爆类型/缩写，用于确定要打勾的 EX 列。"
        )
        ttk.Label(win, text=msg, wraplength=680, justify="left").pack(anchor="w", padx=12, pady=(12, 6))
        ttk.Label(win, text=f"PDF：{pdf_name} | ITR：{itr_idx}").pack(anchor="w", padx=12, pady=(0, 6))

        entry = ttk.Entry(win, width=60)
        entry.pack(padx=12, pady=(0, 12))
        entry.focus()

        result = {"status": "skip", "value": ""}

        def confirm():
            result["status"] = "confirm"
            result["value"] = entry.get()
            win.destroy()

        def skip():
            result["status"] = "skip"
            result["value"] = ""
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Button(btns, text="确认", command=confirm).pack(side="right")
        ttk.Button(btns, text="没有值格（跳过本套）", command=skip).pack(side="right", padx=10)

        win.grab_set()
        self.wait_window(win)
        return result

    def _module_root(self) -> Path:
        path = OUTPUT_ROOT / MODULE_NAME
        path.mkdir(parents=True, exist_ok=True)
        return path

    def open_filled_folder(self) -> None:
        open_in_file_explorer(self._module_root() / "filled")

    def _report_root(self) -> Path:
        path = REPORT_ROOT / REPORT_MODULE_NAME
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _batch_id(self) -> str:
        return get_batch_id()

    def _batch_dir(self, output_type: str, batch_id: str) -> Path:
        if output_type == "report":
            return ensure_report_dir(MODULE_NAME, batch_id)
        return ensure_output_dir(MODULE_NAME, output_type, batch_id)

    def _write_skipped_report(self, report_dir: Path, skipped: list[dict]) -> Path:
        report_dir.mkdir(parents=True, exist_ok=True)
        report_path = report_dir / "skipped_pages.csv"
        with open(report_path, "w", encoding="utf-8", newline="") as f:
            writer = csv.DictWriter(f, fieldnames=["pdf", "page", "reason"])
            writer.writeheader()
            writer.writerows(skipped)
        return report_path

    def _disable_actions(self):
        self.btn_pick.configure(state="disabled")
        self.btn_parse.configure(state="disabled")
        self.btn_test.configure(state="disabled")
        self.btn_tick.configure(state="disabled")

    def _enable_actions(self):
        self.btn_pick.configure(state="normal")
        self._apply_state()

    # ---------- 选择帮助 ----------
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

    # ---------- 动作 ----------
    def pick_pdfs(self):
        paths = filedialog.askopenfilenames(title="选择一个或多个 ITR PDF", filetypes=[("PDF", "*.pdf")])
        if not paths:
            return

        # 去重（保持选择顺序）
        seen = set()
        new_list = []
        for p in paths:
            if p not in seen:
                seen.add(p)
                new_list.append(p)

        self.pdf_paths = list(new_list)
        self.parsed_map.clear()
        self.pages_per_itr_map.clear()
        self.is_parsed_ok = False

        # 刷新列表
        self.lst_pdfs.delete(0, tk.END)
        for p in self.pdf_paths:
            self.lst_pdfs.insert(tk.END, Path(p).name)

        # 清空概览
        for item in self.tree.get_children():
            self.tree.delete(item)

        self._log(f"已批量导入 PDF：{len(self.pdf_paths)} 个")
        for p in self.pdf_paths:
            self._log(f"  - {p}")
        self._log("请先点击【解析】（对选中PDF或全部PDF）再进行测试/打勾。")

        self._update_status()
        self._apply_state()

    def _get_pages_per_itr(self, doc: fitz.Document) -> int:
        if self.page_mode.get() == "regex":
            pattern = self.page_regex.get().strip() or r"Page\s*\d+\s*of\s*(\d+)"
            try:
                scan_n = int(self.page_scan_n.get().strip() or "5")
            except Exception:
                scan_n = 5
            v = parse_pages_per_itr_regex(doc, pattern, scan_n)
            if v and v > 0:
                return v

        try:
            m = int(self.page_manual.get().strip() or "4")
            if m > 0:
                return m
        except Exception:
            pass
        return 4

    def parse_selected(self):
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return

        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("提示", "正在解析，请稍等…")
            return

        sel = self._get_selected_pdfs()
        targets = sel if sel else self.pdf_paths
        num_candidates = [x.strip() for x in (self.num_header_var.get() or "").split(",") if x.strip()]
        if not num_candidates:
            num_candidates = ["NO", "NUMBER", "ITM", "ITEM"]

        self._disable_actions()
        self.status.config(text="解析中…")
        self._worker_thread = threading.Thread(
            target=self._parse_worker,
            args=(targets, num_candidates),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_queue)

    def _parse_worker(self, targets: list[str], num_candidates: list[str]):
        parsed_map: dict[str, list[dict]] = {}
        pages_per_itr_map: dict[str, int] = {}
        parsed_pdf_count = 0
        total_itr_ok = 0

        self._q.put(("log", f"[解析] 目标PDF数={len(targets)}（不选则默认全部）"))
        self._q.put(("log", f"[解析] 底表序号候选={list(map(norm_text, num_candidates))}"))

        for pdf_path in targets:
            try:
                doc = fitz.open(pdf_path)
            except Exception as e:
                self._q.put(("log", f"❌ 打开失败：{Path(pdf_path).name} -> {e}"))
                continue

            pages_per_itr = self._get_pages_per_itr(doc)
            pages_per_itr_map[pdf_path] = pages_per_itr
            itr_count = (doc.page_count + pages_per_itr - 1) // pages_per_itr

            ok_count = 0
            parsed_list = []

            for i in range(itr_count):
                pidx = i * pages_per_itr
                if pidx >= doc.page_count:
                    break

                page = doc[pidx]
                verticals, horizontals = extract_rulings(page)

                no_cell = find_lowest_header_anchor(page, num_candidates, verticals, horizontals)
                if not no_cell:
                    parsed_list.append(
                        {
                            "itr": i + 1,
                            "page_index": pidx,
                            "no_cell": None,
                            "band": None,
                            "desc_cell": None,
                            "ex_cells": [],
                            "okna": {},
                            "ex_label": None,
                            "ex_value": None,
                        }
                    )
                    continue

                band = header_row_band(no_cell)
                desc_cell = find_cell_by_exact_norm(page, "Description", verticals, horizontals, search_clip=band)
                ex_cells = collect_ex_header_cells(page, band, verticals, horizontals)
                okna = find_ok_na_pl_cells(page, band, verticals, horizontals)
                ex_label, ex_value = find_ex_concept_cells(page, verticals, horizontals)

                parsed_list.append(
                    {
                        "itr": i + 1,
                        "page_index": pidx,
                        "no_cell": no_cell,
                        "band": band,
                        "desc_cell": desc_cell,
                        "ex_cells": ex_cells,
                        "okna": okna,
                        "ex_label": ex_label,
                        "ex_value": ex_value,
                    }
                )
                ok_count += 1

            doc.close()

            if ok_count > 0:
                parsed_map[pdf_path] = parsed_list
                parsed_pdf_count += 1
                total_itr_ok += ok_count
                self._q.put(("log", f"✅ {Path(pdf_path).name} 解析成功：{ok_count} 套 ITR（每套 {pages_per_itr} 页）"))
            else:
                self._q.put(("log", f"❌ {Path(pdf_path).name} 解析失败：未找到任何 ITR 的 NO 表头，请检查配置或PDF结构。"))

        self._q.put(
            (
                "parse_done",
                parsed_map,
                pages_per_itr_map,
                parsed_pdf_count,
                total_itr_ok,
            )
        )

    def test_selected(self):
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return

        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("提示", "正在运行，请稍等…")
            return

        sel = self._get_selected_pdfs()
        targets = sel if sel else self.pdf_paths

        # 强制流程：选中范围必须全部解析成功
        not_ok = [p for p in targets if (p not in self.parsed_map or not self.parsed_map[p])]
        if not_ok:
            messagebox.showwarning(
                "提示",
                "以下 PDF 尚未解析成功，无法测试：\n"
                + "\n".join(Path(p).name for p in not_ok)
                + "\n\n请先点【解析】。",
            )
            return

        self._disable_actions()
        self.status.config(text="测试中…")

        batch_id = self._batch_id()
        out_dir = self._batch_dir("test", batch_id)
        debug_labels = bool(self.debug_labels.get())

        self._worker_thread = threading.Thread(
            target=self._test_worker,
            args=(targets, out_dir, debug_labels),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_queue)

    def _test_worker(self, targets: list[str], out_dir: Path, debug_labels: bool):
        for pdf_path in targets:
            stem = Path(pdf_path).stem
            out_pdf = out_dir / f"{stem}_test_boxes.pdf"

            doc = fitz.open(pdf_path)
            parsed_list = self.parsed_map[pdf_path]

            for entry in parsed_list:
                pidx = entry["page_index"]
                if pidx >= doc.page_count:
                    continue
                page = doc[pidx]

                # NO/Description：蓝框
                if entry.get("no_cell"):
                    rr = entry["no_cell"]
                    page.draw_rect(rr, color=(0, 0, 1), width=2.0)
                    if debug_labels:
                        fit_text_to_box(page, _label_rect_above(rr), "NO", {"max_font_size": 6, "min_font_size": 3})

                if entry.get("desc_cell"):
                    rr = entry["desc_cell"]
                    page.draw_rect(rr, color=(0, 0, 1), width=2.0)
                    if debug_labels:
                        fit_text_to_box(page, _label_rect_above(rr), "DESC", {"max_font_size": 6, "min_font_size": 3})

                # OK/NA/PL：蓝框
                for k, rr in (entry.get("okna") or {}).items():
                    page.draw_rect(rr, color=(0, 0, 1), width=2.0)
                    if debug_labels:
                        fit_text_to_box(page, _label_rect_above(rr), k, {"max_font_size": 6, "min_font_size": 3})

                # EX 列：橙框
                for name, rr in (entry.get("ex_cells") or []):
                    page.draw_rect(rr, color=(1, 0.5, 0), width=2.0)
                    if debug_labels:
                        fit_text_to_box(page, _label_rect_above(rr), name, {"max_font_size": 6, "min_font_size": 3})

                # Ex Concept：紫（label）+ 红（值）
                if entry.get("ex_label"):
                    rr = entry["ex_label"]
                    page.draw_rect(rr, color=(0.6, 0, 0.8), width=2.0)
                    if debug_labels:
                        fit_text_to_box(
                            page,
                            _label_rect_above(rr),
                            "EX_CONCEPT",
                            {"max_font_size": 6, "min_font_size": 3},
                        )

                if entry.get("ex_value"):
                    rr = entry["ex_value"]
                    page.draw_rect(rr, color=(1, 0, 0), width=2.0)
                    if debug_labels:
                        fit_text_to_box(
                            page,
                            _label_rect_above(rr),
                            "EX_VALUE",
                            {"max_font_size": 6, "min_font_size": 3},
                        )

            doc.save(out_pdf)
            doc.close()
            self._q.put(("log", f"[测试] 已生成：{out_pdf}"))

        self._q.put(("test_done", out_dir))

    def tick_na_placeholder(self):
        """
        打勾（NA）——第一阶段（只做“NA 自动勾选”）
        ------------------------------------------------
        列格式化思路（你昨晚认可的那套）：
        1) 解析阶段已抓到：NO 表头单元格、九个 EX 缩写列头单元格、OK/NA/PL 列头单元格、Ex Concept 及其值格。
        2) 以这些列头的 x0/x1 作为“列边界”，再用横线把表格切成一行一行。
        3) 对每一行：读取目标 EX 列的交叉格文本；如果该格是 NA，则在同一行的 NA 选项格里写入 “✓”。
        """
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先批量导入 PDF")
            return

        if self._worker_thread and self._worker_thread.is_alive():
            messagebox.showinfo("提示", "正在运行，请稍等…")
            return

        sel = self._get_selected_pdfs()
        targets = sel if sel else self.pdf_paths

        # 强制流程：选中范围须全部解析成功
        not_ok = [p for p in targets if (p not in self.parsed_map or not self.parsed_map[p])]
        if not_ok:
            messagebox.showwarning(
                "提示",
                "以下 PDF 尚未解析成功，无法打勾：\n"
                + "\n".join(Path(p).name for p in not_ok)
                + "\n\n请先点【解析】并用【测试】确认抓取。",
            )
            return

        # 重要：这里不做“全写→缩写”的映射（你后续要做成可配置）。
        # 当前版本：直接把 Ex 值格里提取到的文本，和表头九列缩写做“包含/相等归一化”匹配；
        # 如果匹配不到，就跳过该 ITR（并在日志提示）。
        self._disable_actions()
        self.status.config(text="打勾中…")
        self.progress["value"] = 0
        self.progress_text.config(text="进度：初始化…")
        self._tick_start_time = time.time()

        batch_id = self._batch_id()
        out_dir = self._batch_dir("filled", batch_id)
        report_dir = self._batch_dir("report", batch_id)

        total_pages = self._estimate_tick_total_pages(targets)
        self._tick_total_pages = total_pages
        self.progress["maximum"] = max(total_pages, 1)

        self._worker_thread = threading.Thread(
            target=self._tick_worker,
            args=(targets, out_dir, report_dir, total_pages),
            daemon=True,
        )
        self._worker_thread.start()
        self.after(120, self._poll_queue)

    def _tick_worker(self, targets: list[str], out_dir: Path, report_dir: Path, total_pages: int):
        checked_total = 0
        pdf_done = 0
        skipped_pages: list[dict] = []
        done_pages = 0

        for pdf_path in targets:
            name = Path(pdf_path).name
            out_pdf = out_dir / name
            try:
                if out_pdf.exists():
                    out_pdf.unlink()
            except Exception:
                pass

            doc = fitz.open(pdf_path)
            parsed_list = self.parsed_map[pdf_path]
            pages_per_itr = self.pages_per_itr_map.get(pdf_path, 4)

            self._q.put(("log", f"[打勾] {name} 开始（每套 {pages_per_itr} 页）"))

            for entry in parsed_list:
                itr_idx = entry["itr"]
                base_pidx = entry["page_index"]
                if base_pidx >= doc.page_count:
                    skipped_pages.append(
                        {
                            "pdf": name,
                            "page": base_pidx + 1,
                            "reason": "page_out_of_range",
                        }
                    )
                    done_pages += 1
                    self._q.put(("progress", done_pages, total_pages, name, base_pidx + 1))
                    continue

                ex_val = ""
                ex_val_norm = ""
                if entry.get("ex_value") is None:
                    event = threading.Event()
                    box = {"status": "", "value": ""}
                    self._q.put(("ex_value_request", name, itr_idx, box, event))
                    event.wait()

                    status = box.get("status")
                    manual_val = box.get("value", "")

                    if status != "confirm":
                        pages_to_scan = range(base_pidx, min(base_pidx + pages_per_itr, doc.page_count))
                        for pidx in pages_to_scan:
                            skipped_pages.append(
                                {
                                    "pdf": name,
                                    "page": pidx + 1,
                                    "reason": "missing_ex_value_cell",
                                }
                            )
                            done_pages += 1
                            self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                        self._q.put(("log", f"  - ITR#{itr_idx} 缺少ExConcept值格，用户选择跳过"))
                        continue

                    if not manual_val.strip():
                        pages_to_scan = range(base_pidx, min(base_pidx + pages_per_itr, doc.page_count))
                        for pidx in pages_to_scan:
                            skipped_pages.append(
                                {
                                    "pdf": name,
                                    "page": pidx + 1,
                                    "reason": "empty_manual_ex_input",
                                }
                            )
                            done_pages += 1
                            self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                        self._q.put(("log", f"  - ITR#{itr_idx} 手动输入为空，跳过"))
                        continue

                    ex_val = manual_val.strip()
                    ex_val_norm = norm_text(ex_val)
                else:
                    if entry.get("ex_value"):
                        ex_val = (get_cell_text(doc[base_pidx], entry["ex_value"]) or "").strip()
                    ex_val_norm = norm_text(ex_val)

                ex_headers = entry.get("ex_cells") or []
                if not entry.get("no_cell") or not ex_headers or not (entry.get("okna") or {}).get("NA"):
                    pages_to_scan = range(base_pidx, min(base_pidx + pages_per_itr, doc.page_count))
                    for pidx in pages_to_scan:
                        skipped_pages.append(
                            {
                                "pdf": name,
                                "page": pidx + 1,
                                "reason": "missing_anchor_on_page",
                            }
                        )
                        done_pages += 1
                        self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                    self._q.put(("log", f"  - ITR#{itr_idx} 关键锚点缺失（NO/EX列/NA列），跳过"))
                    continue

                target_ex = None
                for short, rect in ex_headers:
                    if short and (short in ex_val_norm or ex_val_norm in short or short == ex_val_norm):
                        target_ex = (short, rect)
                        break

                if target_ex is None:
                    m = re.search(r"EX([A-Z0-9]{1,3})", ex_val_norm)
                    if m:
                        cand = "EX" + m.group(1)
                        for short, rect in ex_headers:
                            if short == cand:
                                target_ex = (short, rect)
                                break

                if target_ex is None:
                    pages_to_scan = range(base_pidx, min(base_pidx + pages_per_itr, doc.page_count))
                    for pidx in pages_to_scan:
                        skipped_pages.append(
                            {
                                "pdf": name,
                                "page": pidx + 1,
                                "reason": "manual_ex_no_match" if entry.get("ex_value") is None else "target_ex_not_found",
                            }
                        )
                        done_pages += 1
                        self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                    if entry.get("ex_value") is None:
                        self._q.put(("log", f"  - ITR#{itr_idx} 手动输入无法匹配EX缩写列（输入='{ex_val}'），跳过"))
                    else:
                        self._q.put(("log", f"  - ITR#{itr_idx} 无法确定目标EX缩写列（Ex值格='{ex_val}'），跳过"))
                    continue

                target_ex_name, target_ex_rect = target_ex
                no_rect = entry["no_cell"]
                okna = entry["okna"]
                na_hdr = okna.get("NA")
                ok_hdr = okna.get("OK")
                pl_hdr = okna.get("PL")

                x_left = min(no_rect.x0, target_ex_rect.x0, na_hdr.x0)
                x_right = max(
                    no_rect.x1,
                    target_ex_rect.x1,
                    na_hdr.x1,
                    (ok_hdr.x1 if ok_hdr else 0),
                    (pl_hdr.x1 if pl_hdr else 0),
                )

                x_no_c = (no_rect.x0 + no_rect.x1) / 2
                x_ex_c = (target_ex_rect.x0 + target_ex_rect.x1) / 2
                x_na_c = (na_hdr.x0 + na_hdr.x1) / 2

                for pidx in range(base_pidx, min(base_pidx + pages_per_itr, doc.page_count)):
                    page = doc[pidx]

                    verticals, horizontals = extract_rulings(page)
                    xs = _unique_sorted_x_from_verticals(verticals)

                    num_candidates = [x.strip() for x in (self.num_header_var.get() or "").split(",") if x.strip()]
                    if not num_candidates:
                        num_candidates = ["NO", "NUMBER", "ITM", "ITEM"]

                    no_cell2 = find_lowest_header_anchor(page, num_candidates, verticals, horizontals)
                    band = header_row_band(no_cell2) if no_cell2 else None

                    target_rect2 = None
                    na_hdr2 = None
                    no_col_x0, no_col_x1 = None, None
                    ex_col_x0, ex_col_x1 = None, None
                    na_col_x0, na_col_x1 = None, None

                    if band is not None:
                        ex_cells = collect_ex_header_cells(page, band, verticals, horizontals)
                        okna_cells = find_ok_na_pl_cells(page, band, verticals, horizontals)
                        na_hdr2 = okna_cells.get("NA")
                        if ex_cells and na_hdr2:
                            for short, rr in ex_cells:
                                if short == target_ex_name:
                                    target_rect2 = rr
                                    break
                        if no_cell2 and target_rect2 and na_hdr2:
                            no_col_x0, no_col_x1 = no_cell2.x0, no_cell2.x1
                            ex_col_x0, ex_col_x1 = target_rect2.x0, target_rect2.x1
                            na_col_x0, na_col_x1 = na_hdr2.x0, na_hdr2.x1

                    if no_col_x0 is None or ex_col_x0 is None or na_col_x0 is None:
                        if len(xs) < 2:
                            skipped_pages.append(
                                {
                                    "pdf": name,
                                    "page": pidx + 1,
                                    "reason": "insufficient_verticals",
                                }
                            )
                            done_pages += 1
                            self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                            continue
                        no_bounds = _snap_col_bounds(xs, x_no_c)
                        ex_bounds = _snap_col_bounds(xs, x_ex_c)
                        na_bounds = _snap_col_bounds(xs, x_na_c)
                        if not no_bounds or not ex_bounds or not na_bounds:
                            skipped_pages.append(
                                {
                                    "pdf": name,
                                    "page": pidx + 1,
                                    "reason": "missing_column_bounds",
                                }
                            )
                            done_pages += 1
                            self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                            continue
                        no_col_x0, no_col_x1 = no_bounds
                        ex_col_x0, ex_col_x1 = ex_bounds
                        na_col_x0, na_col_x1 = na_bounds

                    y_start = band.y1 if band is not None else 0
                    row_lines = build_table_row_lines(page, horizontals, x_left, x_right, y_start=y_start)
                    _page_words_by_row = None

                    if len(row_lines) < 3:
                        skipped_pages.append(
                            {
                                "pdf": name,
                                "page": pidx + 1,
                                "reason": "insufficient_row_lines",
                            }
                        )
                        done_pages += 1
                        self._q.put(("progress", done_pages, total_pages, name, pidx + 1))
                        continue

                    found_first_number = False
                    for row_idx, (y0, y1) in enumerate(zip(row_lines, row_lines[1:])):
                        if y1 - y0 < 6:
                            continue

                        if _page_words_by_row is None:
                            table_clip = fitz.Rect(x_left, min(row_lines), x_right, max(row_lines))
                            words = page.get_text("words", clip=table_clip) or []
                            _page_words_by_row = [[] for _ in range(len(row_lines) - 1)]
                            import bisect as _bisect

                            for w in words:
                                cy = (w[1] + w[3]) / 2.0
                                ridx = _bisect.bisect_right(row_lines, cy) - 1
                                if 0 <= ridx < len(_page_words_by_row):
                                    _page_words_by_row[ridx].append(w)

                        row_words = _page_words_by_row[row_idx]

                        no_txt = _cell_text_from_row_words(row_words, no_col_x0, no_col_x1)

                        if not found_first_number:
                            if not is_pure_int(no_txt):
                                continue
                            found_first_number = True

                        target_txt = _cell_text_from_row_words(row_words, ex_col_x0, ex_col_x1)
                        if norm_text(target_txt) != "NA":
                            continue

                        na_opt_rect = rect_between_lines(na_col_x0, na_col_x1, y0, y1)

                        if (page.get_text("text", clip=na_opt_rect) or "").strip():
                            continue

                        draw_checkmark(page, na_opt_rect, width=1.6)
                        checked_total += 1
                    done_pages += 1
                    self._q.put(("progress", done_pages, total_pages, name, pidx + 1))

            doc.save(out_pdf)
            doc.close()
            pdf_done += 1
            self._q.put(("log", f"[打勾] {name} 完成，输出：{out_pdf}"))

        report_path = None
        if skipped_pages:
            report_path = self._write_skipped_report(report_dir, skipped_pages)

        self._q.put(("tick_done", pdf_done, checked_total, out_dir, report_path, skipped_pages))

    def _estimate_tick_total_pages(self, targets: list[str]) -> int:
        total = 0
        for pdf_path in targets:
            try:
                doc = fitz.open(pdf_path)
            except Exception:
                continue
            page_count = doc.page_count
            doc.close()
            parsed_list = self.parsed_map.get(pdf_path, [])
            pages_per_itr = self.pages_per_itr_map.get(pdf_path, 4)
            for entry in parsed_list:
                base_pidx = entry.get("page_index", 0)
                if base_pidx >= page_count:
                    total += 1
                else:
                    total += min(pages_per_itr, page_count - base_pidx)
        return total
