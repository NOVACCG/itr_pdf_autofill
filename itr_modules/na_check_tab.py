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


def norm_text(s: str) -> str:
    """归一化：大写 + 去掉非字母数字（空格/换行/符号都去掉）"""
    return re.sub(r"[^A-Z0-9]+", "", (s or "").upper())


def parse_pages_per_itr_regex(doc: fitz.Document, pattern: str, scan_pages: int) -> int | None:
    """从前 scan_pages 页，用正则抓取 'Page x of y' 的 y（每套 ITR 页数）。"""
    try:
        reg = re.compile(pattern, flags=re.IGNORECASE)
    except re.error:
        return None

    n = min(max(scan_pages, 1), doc.page_count)
    for i in range(n):
        txt = doc[i].get_text("text") or ""
        m = reg.search(txt)
        if m:
            try:
                return int(m.group(1))
            except Exception:
                return None
    return None


# -------------------------
# 表格线提取：横/竖线
# -------------------------
def extract_rulings(page: fitz.Page, tol=1.5):
    verticals = []
    horizontals = []

    drawings = page.get_drawings()
    for d in drawings:
        for it in d.get("items", []):
            if not it:
                continue
            kind = it[0]
            if kind == "l":  # line
                (x0, y0) = it[1]
                (x1, y1) = it[2]
                if abs(x0 - x1) <= tol:
                    x = (x0 + x1) / 2.0
                    verticals.append((x, min(y0, y1), max(y0, y1)))
                elif abs(y0 - y1) <= tol:
                    y = (y0 + y1) / 2.0
                    horizontals.append((y, min(x0, x1), max(x0, x1)))
            elif kind == "re":  # rectangle path
                r = it[1]
                if isinstance(r, fitz.Rect):
                    x0, y0, x1, y1 = r.x0, r.y0, r.x1, r.y1
                    verticals.extend([(x0, y0, y1), (x1, y0, y1)])
                    horizontals.extend([(y0, x0, x1), (y1, x0, x1)])

    verticals = [(x, y0, y1) for (x, y0, y1) in verticals if (y1 - y0) > 6]
    horizontals = [(y, x0, x1) for (y, x0, x1) in horizontals if (x1 - x0) > 6]
    return verticals, horizontals


def cell_rect_for_word(word_rect: fitz.Rect, verticals, horizontals):
    cx0, cy0, cx1, cy1 = word_rect.x0, word_rect.y0, word_rect.x1, word_rect.y1
    y_mid = (cy0 + cy1) / 2.0
    x_mid = (cx0 + cx1) / 2.0

    left = None
    right = None
    for x, y0, y1 in verticals:
        if y0 - 2 <= y_mid <= y1 + 2:
            if x <= cx0 + 2 and (left is None or x > left):
                left = x
            if x >= cx1 - 2 and (right is None or x < right):
                right = x

    top = None
    bottom = None
    for y, x0, x1 in horizontals:
        if x0 - 2 <= x_mid <= x1 + 2:
            if y <= cy0 + 2 and (top is None or y > top):
                top = y
            if y >= cy1 - 2 and (bottom is None or y < bottom):
                bottom = y

    if left is None or right is None or top is None or bottom is None:
        return None
    return fitz.Rect(left + 0.3, top + 0.3, right - 0.3, bottom - 0.3)


def get_cell_text(page: fitz.Page, cell: fitz.Rect) -> str:
    return (page.get_text("text", clip=cell) or "").strip()


def _norm_join_words(words_in_row):
    """把一行内的 words (PyMuPDF words 元组)按从左到右拼接成文本。"""
    if not words_in_row:
        return ""
    # words: (x0,y0,x1,y1, text, block, line, word)
    words_in_row = sorted(words_in_row, key=lambda w: (w[0], w[1]))
    return " ".join((w[4] or "").strip() for w in words_in_row if (w[4] or "").strip())


def _cell_text_from_row_words(row_words, x0, x1):
    """从“该行的 words 列表”里取出中心点落在 [x0,x1] 的词，拼成该单元格文本。"""
    if not row_words:
        return ""
    picked = []
    for w in row_words:
        wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
        cx = (wx0 + wx1) / 2.0
        if x0 <= cx <= x1:
            picked.append(w)
    return _norm_join_words(picked)


# -------------------------
# “列格式化”辅助：用表头列边界 + 横线切行
# -------------------------
def _uniq_sorted(vals, tol=0.8):
    """排序+去重：同一位置(差<=tol)视为一条线。"""
    vals = sorted(vals)
    out = []
    for v in vals:
        if not out or abs(v - out[-1]) > tol:
            out.append(v)
    return out


def build_table_row_lines(page: fitz.Page, horizontals, x_left: float, x_right: float, y_start: float, min_span_pad=8.0):
    """从页面横线里筛出能覆盖表格宽度的横线，返回 y 坐标列表（已去重排序）。"""
    ys = []
    for y, x0, x1 in horizontals:
        # 这条横线要足够“横跨”表格宽度
        if x0 <= x_left + min_span_pad and x1 >= x_right - min_span_pad and y >= y_start - 2:
            ys.append(y)
    return _uniq_sorted(ys)


def is_pure_int(s: str) -> bool:
    s = (s or "").strip()
    return bool(s) and s.isdigit()


def rect_between_lines(x0, x1, y0, y1, pad=0.6):
    return fitz.Rect(x0 + pad, y0 + pad, x1 - pad, y1 - pad)


def _unique_sorted_x_from_verticals(verticals) -> list[float]:
    """从竖线集合里提取去重后的 x 坐标（排序）。

    verticals 可能是:
    - (x, y0, y1)  由 extract_rulings() 生成
    - (x0, y0, x1, y1)  兼容旧写法/外部传入
    本函数只关心 x 坐标。
    """
    xs: list[float] = []
    for v in verticals or []:
        if not v:
            continue
        # 兼容 3 元/4 元元组
        if len(v) == 3:
            x, _, _ = v
            xs.append(float(x))
        elif len(v) >= 4:
            x0, _, _, _ = v[:4]
            xs.append(float(x0))
        else:
            # 不认识的结构，跳过
            continue
    xs = sorted({round(x, 2) for x in xs})
    return xs


def _snap_col_bounds(xs: list[float], x_center: float) -> tuple[float, float] | None:
    """给定一堆竖线 x 坐标，返回能“包住”x_center 的相邻两条竖线 (xL, xR)。

    如果找不到严格包住的，就取距离最近的一段。
    """
    if not xs or len(xs) < 2:
        return None
    # 先尝试严格包住
    for i in range(len(xs) - 1):
        if xs[i] - 1.0 <= x_center <= xs[i + 1] + 1.0:
            return (xs[i], xs[i + 1])
    # 兜底：取最近的段（按段中心距离）
    best = None
    best_d = 1e18
    for i in range(len(xs) - 1):
        c = (xs[i] + xs[i + 1]) / 2.0
        d = abs(c - x_center)
        if d < best_d:
            best_d = d
            best = (xs[i], xs[i + 1])
    return best


# -------------------------
# 方法A：先粗搜后“单元格归一化严格匹配”
# -------------------------
def find_cell_by_exact_norm(page: fitz.Page, target: str, verticals, horizontals, search_clip: fitz.Rect | None = None):
    target_norm = norm_text(target)
    words = page.get_text("words", clip=search_clip) if search_clip else page.get_text("words")

    candidates = []
    for x0, y0, x1, y1, w, *_ in words:
        wn = norm_text(w)
        if not wn:
            continue
        if wn in target_norm or target_norm in wn:
            cell = cell_rect_for_word(fitz.Rect(x0, y0, x1, y1), verticals, horizontals)
            if cell:
                candidates.append(cell)

    # 去重
    uniq = []
    for c in candidates:
        if all(
            not (
                abs(c.x0 - u.x0) < 1
                and abs(c.y0 - u.y0) < 1
                and abs(c.x1 - u.x1) < 1
                and abs(c.y1 - u.y1) < 1
            )
            for u in uniq
        ):
            uniq.append(c)

    for cell in uniq:
        if norm_text(get_cell_text(page, cell)) == target_norm:
            return cell
    return None


def find_lowest_header_anchor(page: fitz.Page, candidates: list[str], verticals, horizontals):
    cand_norms = [norm_text(x) for x in candidates]
    words = page.get_text("words")
    hits = []
    for x0, y0, x1, y1, w, *_ in words:
        if norm_text(w) in cand_norms:
            cell = cell_rect_for_word(fitz.Rect(x0, y0, x1, y1), verticals, horizontals)
            if not cell:
                continue
            if norm_text(get_cell_text(page, cell)) in cand_norms:
                hits.append(cell)
    if not hits:
        return None
    hits.sort(key=lambda r: r.y0, reverse=True)
    return hits[0]


def header_row_band(no_cell: fitz.Rect, pad=3.0):
    return fitz.Rect(0, no_cell.y0 - pad, 10000, no_cell.y1 + pad)


def collect_ex_header_cells(page: fitz.Page, row_band: fitz.Rect, verticals, horizontals):
    """
    收集底表表头里“九个防爆缩写列”（ EXE / EXDE / EXD / EXI ...）。

    关键修复：以前从 words 里抓 EX 开头的词，遇到换行（EX\nDE）会只抓到 "EX"。
    现在改成“按格子取文本”：
    - 在表头 band 内，利用竖线把表头行切成一格一格
    - 对每一格读取完整文本并归一化
    - 如果格子归一化文本符合 EX[A-Z0-9]{1,3}，则认为是一个缩写列头
    """
    band = row_band
    cy = (band.y0 + band.y1) / 2

    # 找到能覆盖表头 band 的所有竖线 x
    xs = []
    for vx, vy0, vy1 in verticals:
        if vy0 <= cy <= vy1:
            xs.append(float(vx))
    xs = sorted(set(round(x, 2) for x in xs))
    if len(xs) < 2:
        return []

    ex_pat = re.compile(r"^EX[A-Z0-9]{1,3}$")
    out = []

    # 逐列扫描（相邻两条竖线定义一个格子宽度）
    for i in range(len(xs) - 1):
        x0, x1 = xs[i], xs[i + 1]
        # 排除极窄的“假列”
        if x1 - x0 < 6:
            continue

        rr = fitz.Rect(x0 + 0.6, band.y0 + 0.6, x1 - 0.6, band.y1 - 0.6)
        txt = norm_text(get_cell_text(page, rr))
        if not txt:
            continue
        # 只要格子里完整文本符合 EX* 缩写，就收录
        if ex_pat.match(txt):
            out.append((txt, rr))

    # 去重并按 x0 排序
    uniq = {}
    for k, rr in out:
        if k not in uniq:
            uniq[k] = rr
    items = sorted(uniq.items(), key=lambda kv: kv[1].x0)
    return [(k, rr) for k, rr in items]


def find_ok_na_pl_cells(page: fitz.Page, row_band: fitz.Rect, verticals, horizontals):
    res = {}
    for key in ["OK", "NA", "PL"]:
        cell = find_cell_by_exact_norm(page, key, verticals, horizontals, search_clip=row_band)
        if cell:
            res[key] = cell
    return res


def find_ex_concept_cells(page: fitz.Page, verticals, horizontals):
    label = find_cell_by_exact_norm(page, "Ex Concept", verticals, horizontals)
    if not label:
        return None, None

    y_mid = (label.y0 + label.y1) / 2
    right_lines = [x for x, y0, y1 in verticals if (y0 - 2 <= y_mid <= y1 + 2) and x > label.x1 + 1]
    if not right_lines:
        return label, None
    right_edge = min(right_lines)

    value_cell = fitz.Rect(label.x1 + 0.3, label.y0 + 0.3, right_edge - 0.3, label.y1 - 0.3)
    if not norm_text(get_cell_text(page, value_cell)).startswith("EX"):
        return label, None
    return label, value_cell


def draw_checkmark(page: fitz.Page, rr: fitz.Rect, width: float = 1.6):
    """
    画勾：用两条线组成 √，避免写入“✓”字符时因字体缺失变成小点。

    ✅ 关键优化：不再按“长方形整格”比例画勾，而是：
    - 取 rr 的最短边作为边长，在单元格中心构造一个正方形绘制区域
    - 在该正方形内画 √，保证勾永远不被拉长，打印更清晰
    """
    if rr is None:
        return

    # 以最短边构造“正方形画布”
    side = max(min(rr.width, rr.height), 1.0)
    cx = (rr.x0 + rr.x1) / 2.0
    cy = (rr.y0 + rr.y1) / 2.0
    sq = fitz.Rect(cx - side / 2.0, cy - side / 2.0, cx + side / 2.0, cy + side / 2.0)

    # 内缩：避免碰到格子边框
    inset = max(side * 0.18, 1.0)
    r = fitz.Rect(sq.x0 + inset, sq.y0 + inset, sq.x1 - inset, sq.y1 - inset)

    w = r.width
    h = r.height

    # √ 的三点（相对于正方形 r）
    p1 = (r.x0 + 0.18 * w, r.y0 + 0.55 * h)
    p2 = (r.x0 + 0.42 * w, r.y0 + 0.78 * h)
    p3 = (r.x0 + 0.82 * w, r.y0 + 0.22 * h)

    page.draw_line(p1, p2, width=width)
    page.draw_line(p2, p3, width=width)


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
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), "NO", fontsize=6)

                if entry.get("desc_cell"):
                    rr = entry["desc_cell"]
                    page.draw_rect(rr, color=(0, 0, 1), width=2.0)
                    if debug_labels:
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), "DESC", fontsize=6)

                # OK/NA/PL：蓝框
                for k, rr in (entry.get("okna") or {}).items():
                    page.draw_rect(rr, color=(0, 0, 1), width=2.0)
                    if debug_labels:
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), k, fontsize=6)

                # EX 列：橙框
                for name, rr in (entry.get("ex_cells") or []):
                    page.draw_rect(rr, color=(1, 0.5, 0), width=2.0)
                    if debug_labels:
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), name, fontsize=6)

                # Ex Concept：紫（label）+ 红（值）
                if entry.get("ex_label"):
                    rr = entry["ex_label"]
                    page.draw_rect(rr, color=(0.6, 0, 0.8), width=2.0)
                    if debug_labels:
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), "EX_CONCEPT", fontsize=6)

                if entry.get("ex_value"):
                    rr = entry["ex_value"]
                    page.draw_rect(rr, color=(1, 0, 0), width=2.0)
                    if debug_labels:
                        page.insert_text((rr.x0 + 1, rr.y0 - 1), "EX_VALUE", fontsize=6)

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