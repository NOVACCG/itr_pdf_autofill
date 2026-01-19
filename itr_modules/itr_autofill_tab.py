# itr_universal_autofill_tool_v2.py
# ============================================================
# 通用 ITR PDF 自动预填工具（带预设管理 / 字段映射 / PDF定位测试 / 导出不卡顿）
# 依赖：pip install pymupdf pandas openpyxl
#
# 目录说明（脚本同目录自动生成）：
# - presets/                预设文件（.json）
# - output/itr_autofill/     导出结果（filled PDF + 测试PDF）
# - report/itr_autofill/     报告输出（report.xlsx）
# - match_memory.json       模糊匹配记忆库（短Key -> ExcelKey）
# - config_global.json      当前使用预设名
# ============================================================

import os
import sys
import re
import json
import datetime
import threading
import queue
from dataclasses import dataclass
from typing import Dict, List, Tuple, Optional, Set
from pathlib import Path

import fitz  # PyMuPDF
import pandas as pd

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# -------------------------
# 路径与常量
# -------------------------
APP_NAME = "ITR 自动预填工具"
APP_VERSION = "V1.0.3"

from itr_modules.shared.paths import BASE_DIR, ensure_output_dir, ensure_report_dir, get_batch_id, open_in_file_explorer
from itr_modules.shared.pdf_utils import (
    extract_tag_by_cell_adjacency,
    fit_text_to_box,
    norm_text,
    normalize_cell_text,
)

PRESETS_DIR = os.path.join(BASE_DIR, "presets")
MODULE_NAME = "itr_autofill"
OUTPUT_TEST_ROOT = os.path.join(BASE_DIR, "output", MODULE_NAME, "test")
OUTPUT_FILLED_ROOT = os.path.join(BASE_DIR, "output", MODULE_NAME, "filled")
GLOBAL_CONFIG_PATH = os.path.join(BASE_DIR, "config_global.json")
MATCH_MEMORY_PATH = os.path.join(BASE_DIR, "match_memory.json")
TAG_CHOICE_MEMORY_PATH = os.path.join(BASE_DIR, "tag_choice_memory.json")

os.makedirs(PRESETS_DIR, exist_ok=True)
os.makedirs(OUTPUT_TEST_ROOT, exist_ok=True)
os.makedirs(OUTPUT_FILLED_ROOT, exist_ok=True)
ensure_report_dir(MODULE_NAME, get_batch_id())

DEFAULT_PAGE1_MARK_RE = re.compile(r"Page\s*1\s*of\s*(\d+)", re.IGNORECASE)
DEFAULT_TAG_PATTERN = r"TAG\s*NO\.\s*:\s*([A-Za-z0-9\-\._/]+)"
DEFAULT_TAG_RE = re.compile(DEFAULT_TAG_PATTERN, re.IGNORECASE)
DEFAULT_VALUE_REGEX = r"([A-Za-z0-9\-\._/]+)"
TAG_DIRECTION_OPTIONS = ["RIGHT", "LEFT", "DOWN", "UP"]

SIDE_OPTIONS = ["", "LEFT", "RIGHT"]
RULE_OPTIONS = ["", "SHEET_NAME", "PDF_NAME", "TODAY", "EMPTY"]
SOURCE_OPTIONS = ["EXCEL", "MANUAL", "CONST", "RULE"]


def now_iso() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def today_ymd() -> str:
    return datetime.datetime.now().strftime("%Y-%m-%d")


def batch_id() -> str:
    return get_batch_id()


def ensure_output_batch_dir(output_type: str, batch: str) -> str:
    return str(ensure_output_dir(MODULE_NAME, output_type, batch))


def ensure_report_batch_dir(batch: str) -> str:
    return str(ensure_report_dir(MODULE_NAME, batch))


def load_json_safe(path: str, default):
    if not os.path.exists(path):
        return default
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, type(default)) else default
    except Exception:
        return default


def save_json_safe(path: str, data):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_global_config() -> dict:
    return load_json_safe(GLOBAL_CONFIG_PATH, {"active_preset": ""})


def save_global_config(cfg: dict):
    save_json_safe(GLOBAL_CONFIG_PATH, cfg)


def load_match_memory() -> dict:
    return load_json_safe(MATCH_MEMORY_PATH, {})


def save_match_memory(mem: dict):
    save_json_safe(MATCH_MEMORY_PATH, mem)


def norm_header(s: str) -> str:
    s = str(s).upper()
    return re.sub(r"[^A-Z0-9]", "", s)


def norm_key_value(v: str) -> str:
    if v is None:
        return ""
    return str(v).replace(" ", "").strip().upper()


def safe_value(v) -> str:
    """Convert an Excel cell value to text WITHOUT 'invalid value' cleanup.

    Behavior:
    - None -> "" (keep PDF writer stable; avoids literal 'None')
    - Anything else -> string form as-is (no special-casing for '-', '—', '', NaN, etc.)
    """
    if v is None:
        return ""
    try:
        return str(v)
    except Exception:
        return ""


def default_preset() -> dict:
    """给一个开箱即用的默认预设（适配你们当前 Ex ITR 结构，可自行改）。"""
    return {
        "preset_name": "Default_Ex_ITR",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "notes": "",
        "itr_pages_per_set": 4,
        "page1_mark_regex": r"Page\s*1\s*of\s*(\d+)",
        "match": {
            "key_name": "TAG",
            "pdf_extract_regex": DEFAULT_VALUE_REGEX,
            "tag_direction": "RIGHT",
            "strip_suffixes": ["-EX"],  # 例如 PDF 里是 627-xx-Ex，而 Excel 里是 627-xx
            "excel_key_col_candidates_norm": ["TAGNO", "TAG", "TAGNUMBER", "EQUIPMENTTAG"],
            "enable_fuzzy": True,
            "fuzzy_require_confirm": True,
            "fuzzy_show_topn": 10,
        },
        "excel": {
            # ⚠️ 注意：这里是“从 0 开始计数”的表头行。
            # 例如：你在 Excel 里肉眼看到列名在第 3 行，这里填 2。
            "header_row": 2
        },
        "text": {"max_font_size": 9, "min_font_size": 5, "padding": 2, "line_gap": 1.15},
        "fields": [
            {"name": "Location", "pdf_label": "Location", "pdf_label_side": "", "page_scope": [1],
             "source": "RULE", "excel_col_norm": "", "const_value": "", "rule": "SHEET_NAME"},
            {"name": "Zone", "pdf_label": "Zone", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "HAZARDOUSCLASS", "const_value": "", "rule": ""},

            {"name": "GasGroup_Env", "pdf_label": "Gas Group", "pdf_label_side": "LEFT", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "ENVIRONMENTGASGROUP", "const_value": "", "rule": ""},
            {"name": "TempClass_Env", "pdf_label": "Temp Class", "pdf_label_side": "LEFT", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "ENVIRONMENTTEMPCLASS", "const_value": "", "rule": ""},

            {"name": "ExCertificate", "pdf_label": "Ex Certificate", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "IECATEXCERTIFICATION", "const_value": "", "rule": ""},
            {"name": "Model", "pdf_label": "Model", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "TYPEMODEL", "const_value": "", "rule": ""},
            {"name": "ExConcept", "pdf_label": "Ex Concept", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "EXCLASS", "const_value": "", "rule": ""},
            {"name": "Manufacturer", "pdf_label": "Manufacturer", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "MANUFACTURER", "const_value": "", "rule": ""},

            {"name": "CertBody", "pdf_label": "Cert. Body", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "IECATEXNOTIFIEDBODY", "const_value": "", "rule": ""},

            {"name": "GasGroup_Equip", "pdf_label": "Gas Group", "pdf_label_side": "RIGHT", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "EQUIPMENTGROUP", "const_value": "", "rule": ""},
            {"name": "TempClass_Equip", "pdf_label": "Temp Class", "pdf_label_side": "RIGHT", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "EQUIPMENTTEMPCLASS", "const_value": "", "rule": ""},
            {"name": "IPRating", "pdf_label": "IP Rating", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "EQUIPMENTIPRATING", "const_value": "", "rule": ""},
            {"name": "ProductDate", "pdf_label": "Product Date", "pdf_label_side": "", "page_scope": [1],
             "source": "EXCEL", "excel_col_norm": "PRODUCTDATE", "const_value": "", "rule": ""},

            # Serial Number：默认不从 Excel 找，留给你在界面上手填（会写进PDF）
            {"name": "SerialNumber", "pdf_label": "Serial Number", "pdf_label_side": "", "page_scope": [1],
             "source": "MANUAL", "excel_col_norm": "", "const_value": "", "rule": ""},
        ],
    }


def preset_path(name: str) -> str:
    return os.path.join(PRESETS_DIR, f"{name}.json")


def list_presets() -> List[str]:
    names = []
    for fn in os.listdir(PRESETS_DIR):
        if fn.lower().endswith(".json"):
            names.append(os.path.splitext(fn)[0])
    return sorted(names)


def load_preset(name: str) -> Optional[dict]:
    p = preset_path(name)
    if not os.path.exists(p):
        return None
    try:
        with open(p, "r", encoding="utf-8") as f:
            d = json.load(f)
            return d if isinstance(d, dict) else None
    except Exception:
        return None


def save_preset(name: str, data: dict):
    data["preset_name"] = name
    if not data.get("created_at"):
        data["created_at"] = now_iso()
    data["updated_at"] = now_iso()

    # 补齐结构（旧预设升级）
    d0 = default_preset()
    data.setdefault("match", d0["match"])
    data.setdefault("excel", d0["excel"])
    data.setdefault("text", d0["text"])
    data.setdefault("itr_pages_per_set", d0["itr_pages_per_set"])
    data.setdefault("page1_mark_regex", d0["page1_mark_regex"])
    if "fields" not in data or not isinstance(data["fields"], list):
        data["fields"] = d0["fields"]

    with open(preset_path(name), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def compile_re(pattern: str, fallback: re.Pattern) -> re.Pattern:
    try:
        return re.compile(pattern, re.IGNORECASE)
    except Exception:
        return fallback


# -------------------------
# Excel 索引
# -------------------------
def build_excel_index(excel_path: str, preset: dict) -> dict:
    if not str(excel_path).lower().endswith(".xlsx"):
        raise ValueError("仅支持 .xlsx 格式的 Excel，请先另存为 .xlsx")
    """
    建索引：ExcelKey -> (sheet_name, row_dict, col_map_norm, key_col_name)
    其中 ExcelKey 由 match.excel_key_col_candidates_norm 指定的列读取。
    """
    header_row = int(preset.get("excel", {}).get("header_row", 0))
    key_candidates = [norm_header(x) for x in preset.get("match", {}).get("excel_key_col_candidates_norm", [])]

    xls = pd.ExcelFile(excel_path)
    idx = {}

    for sheet in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=sheet, header=header_row)
        except Exception:
            continue

        cols = list(df.columns)
        col_map_norm = {norm_header(c): c for c in cols}
        key_col = None
        for cand in key_candidates:
            if cand in col_map_norm:
                key_col = col_map_norm[cand]
                break
        if key_col is None:
            continue

        key_idx = cols.index(key_col)
        for row in df.itertuples(index=False, name=None):
            key_val = norm_key_value(row[key_idx])
            if not key_val or key_val in idx:
                continue
            row_dict = dict(zip(cols, row))
            idx[key_val] = (sheet, row_dict, col_map_norm, key_col)

    return idx


# -------------------------
# PDF 解析：找 ITR 套起始页 & 从 Page1 提取 MatchKey
# -------------------------
def find_itr_start_pages(pdf_path: str, preset: dict, doc: Optional[fitz.Document] = None) -> List[int]:
    """
    返回每套 ITR 的起始页（1-based）。
    优先：用 Page1 标记正则（例如 Page 1 of 4）。
    若 PDF 没有页码标记，则兜底：按 itr_pages_per_set 固定分组（1, 1+N, 1+2N ...）。
    （可用于“两页一套但无页码”的情况——只要你在预设里把 itr_pages_per_set 设为 2）
    """
    should_close = False
    if doc is None:
        doc = fitz.open(pdf_path)
        should_close = True
    pages_per_set = int(preset.get("itr_pages_per_set", 1)) or 1

    mark_re = compile_re(preset.get("page1_mark_regex", ""), DEFAULT_PAGE1_MARK_RE)
    starts = []
    for i in range(len(doc)):
        t = doc[i].get_text("text")
        if t and mark_re.search(t):
            starts.append(i + 1)  # 1-based

    if not starts:
        # 没有页码标记：按固定页数拆分
        starts = list(range(1, len(doc) + 1, pages_per_set))

    if should_close:
        doc.close()
    return starts


def extract_match_key_from_page(page: fitz.Page, preset: dict) -> str:
    rx = compile_re(preset.get("match", {}).get("pdf_extract_regex", ""), DEFAULT_TAG_RE)
    t = page.get_text("text")
    m = rx.search(t)
    if not m:
        return ""
    key = norm_key_value(m.group(1).strip())

    # 去除后缀（仅末尾）
    for suf in preset.get("match", {}).get("strip_suffixes", []):
        suf_up = norm_key_value(suf)
        if suf_up and key.endswith(suf_up):
            key = key[: -len(suf_up)]
    return key


def candidate_keys(key_pdf: str, preset: dict) -> List[str]:
    """给一些候选：本体 + 可能的 -EX 变体（用于容错）。"""
    key = norm_key_value(key_pdf)
    cands: Set[str] = {key}
    strip_sufs = [norm_key_value(s) for s in preset.get("match", {}).get("strip_suffixes", [])]
    if "EX" in strip_sufs or "-EX" in strip_sufs:
        cands.add(key + "-EX")
    return [c for c in cands if c]


def fuzzy_find_keys(excel_index: dict, short_key: str) -> List[str]:
    keys = list(excel_index.keys())
    ends = [k for k in keys if k.endswith(short_key)]
    cont = [k for k in keys if short_key in k and k not in ends]
    return ends + cont


def match_one(key_pdf: str, excel_index: dict, preset: dict, match_memory: dict, chooser_func=None):
    """
    返回：
    status, excel_key, sheet_name, payload, short_key
    payload=(sheet,row_dict,col_map_norm) 或 None
    """
    key_pdf_norm = norm_key_value(key_pdf)
    if not key_pdf_norm:
        return "not_found", "", "", None, ""

    short_key = key_pdf_norm
    for suf in preset.get("match", {}).get("strip_suffixes", []):
        suf2 = norm_key_value(suf)
        if suf2 and short_key.endswith(suf2):
            short_key = short_key[: -len(suf2)]

    remembered = match_memory.get(short_key, "")
    if remembered and remembered in excel_index:
        sheet, row_dict, col_map_norm, _ = excel_index[remembered]
        return "memory", remembered, sheet, (sheet, row_dict, col_map_norm), short_key

    for c in candidate_keys(key_pdf_norm, preset):
        if c in excel_index:
            sheet, row_dict, col_map_norm, _ = excel_index[c]
            status = "exact" if c == key_pdf_norm else "variant"
            return status, c, sheet, (sheet, row_dict, col_map_norm), short_key

    if not preset.get("match", {}).get("enable_fuzzy", True):
        return "not_found", "", "", None, short_key

    fuzzy = fuzzy_find_keys(excel_index, short_key)
    if not fuzzy:
        return "not_found", "", "", None, short_key

    topn = int(preset.get("match", {}).get("fuzzy_show_topn", 10))
    show = fuzzy[:topn]

    if preset.get("match", {}).get("fuzzy_require_confirm", True):
        if chooser_func is None:
            return "skipped", "", "", None, short_key
        pick = chooser_func(key_pdf_norm, short_key, show)
        if not pick:
            return "skipped", "", "", None, short_key
        chosen = pick
    else:
        chosen = show[0]

    sheet, row_dict, col_map_norm, _ = excel_index[chosen]
    return "fuzzy", chosen, sheet, (sheet, row_dict, col_map_norm), short_key


# -------------------------
# 字段取值：EXCEL / MANUAL / CONST / RULE
# -------------------------
def compute_filled(
    preset: dict,
    sheet_name: str,
    row_dict: Optional[dict],
    col_map_norm: Optional[dict],
    pdf_name: str,
) -> Dict[str, str]:
    filled: Dict[str, str] = {}
    for f in preset.get("fields", []):
        name = f.get("name", "")
        if not name:
            continue
        src = (f.get("source", "MANUAL") or "MANUAL").upper()

        if src == "MANUAL":
            filled[name] = ""
        elif src == "CONST":
            filled[name] = str(f.get("const_value", "") or "")
        elif src == "RULE":
            rule = (f.get("rule", "") or "").upper()
            if rule == "SHEET_NAME":
                filled[name] = sheet_name or ""
            elif rule == "PDF_NAME":
                filled[name] = os.path.splitext(pdf_name)[0] if pdf_name else ""
            elif rule == "TODAY":
                filled[name] = today_ymd()
            elif rule == "EMPTY":
                filled[name] = ""
            else:
                filled[name] = ""
        elif src == "EXCEL":
            col_norm = norm_header(f.get("excel_col_norm", ""))
            if not row_dict or not col_map_norm or not col_norm:
                filled[name] = ""
            else:
                raw_col = col_map_norm.get(col_norm)
                filled[name] = safe_value(row_dict.get(raw_col)) if raw_col else ""
        else:
            filled[name] = ""
    return filled


# -------------------------
# 表格线提取 / 定位 / 写字（自动换行 & 字号自适应）
# -------------------------
def collect_line_segments(page: fitz.Page):
    drawings = page.get_drawings()
    segs = []
    for d in drawings:
        for it in d["items"]:
            if it[0] == "l":
                segs.append((it[1], it[2]))
    return segs


def row_verticals(line_segments, mid_y: float) -> List[float]:
    xs = set()
    for p1, p2 in line_segments:
        if abs(p1.x - p2.x) < 0.6 and abs(p1.y - p2.y) > 6:
            ymin, ymax = sorted([p1.y, p2.y])
            if ymin - 1 <= mid_y <= ymax + 1:
                xs.add(round(p1.x, 1))
    return sorted(xs)


def col_horizontals(line_segments, mid_x: float) -> List[float]:
    ys = set()
    for p1, p2 in line_segments:
        if abs(p1.y - p2.y) < 0.6 and abs(p1.x - p2.x) > 6:
            xmin, xmax = sorted([p1.x, p2.x])
            if xmin - 1 <= mid_x <= xmax + 1:
                ys.add(round(p1.y, 1))
    return sorted(ys)


def label_variants(raw_label: str) -> List[str]:
    s = str(raw_label).strip()
    if not s:
        return []
    variants = {s, s + ":", s.replace(":", ""), s.replace("  ", " ")}
    variants.add(s + " :")
    variants.add(s + "：")
    variants.add(s.replace(" ", "  "))
    return [v for v in variants if v]


def search_label_rect(page: fitz.Page, field: dict) -> List[fitz.Rect]:
    lab = field.get("pdf_label", "")
    vars_ = label_variants(lab)
    found = []
    for v in vars_:
        rects = page.search_for(v)
        if rects:
            found.extend(rects)
    uniq = []
    for r in found:
        key = (round(r.x0, 1), round(r.y0, 1), round(r.x1, 1), round(r.y1, 1))
        if all(
            (round(u.x0, 1), round(u.y0, 1), round(u.x1, 1), round(u.y1, 1)) != key for u in uniq
        ):
            uniq.append(r)
    return uniq


def pick_label_rect_for_side(rects: List[fitz.Rect], side: str) -> Optional[fitz.Rect]:
    if not rects:
        return None
    if len(rects) == 1 or not side:
        return sorted(rects, key=lambda r: (r.y0, r.x0))[0]
    rects_sorted = sorted(rects, key=lambda r: r.x0)
    side = (side or "").upper()
    if side == "LEFT":
        return rects_sorted[0]
    if side == "RIGHT":
        return rects_sorted[-1]
    return sorted(rects, key=lambda r: (r.y0, r.x0))[0]


def find_cell_right_of_label(page: fitz.Page, line_segments, label_rect: fitz.Rect) -> Optional[fitz.Rect]:
    """通过表格线，找 label 右侧的单元格框（写入区域）。"""
    mid_y = (label_rect.y0 + label_rect.y1) / 2
    mid_x = (label_rect.x0 + label_rect.x1) / 2

    xs = row_verticals(line_segments, mid_y)
    ys = col_horizontals(line_segments, mid_x)

    if not ys:
        y0 = label_rect.y0 - 2
        y1 = label_rect.y1 + 2
    else:
        y0_candidates = [y for y in ys if y <= mid_y]
        y1_candidates = [y for y in ys if y >= mid_y]
        if not y0_candidates or not y1_candidates:
            y0 = label_rect.y0 - 2
            y1 = label_rect.y1 + 2
        else:
            y0 = max(y0_candidates)
            y1 = min(y1_candidates)

    page_w = page.rect.width

    right_lines = [x for x in xs if x > label_rect.x1 + 0.5]
    x0 = min(right_lines) if right_lines else (label_rect.x1 + 1)

    after_x0 = [x for x in xs if x > x0 + 0.5]
    if after_x0:
        x1 = min(after_x0)
    elif xs:
        x1 = max(xs)
        if x1 <= x0 + 5:
            x1 = page_w - 36
    else:
        x1 = page_w - 36

    if x1 <= x0 + 5 or y1 <= y0 + 4:
        return None
    return fitz.Rect(x0, y0, x1, y1)


# -------------------------
# PDF 定位测试：画框
# -------------------------
def pdf_position_test(pdf_path: str, preset: dict, fields: List[dict]) -> Tuple[str, List[str]]:
    logs = []
    if not os.path.exists(pdf_path):
        return "", ["PDF不存在"]

    doc = fitz.open(pdf_path)
    pages_per_set = int(preset.get("itr_pages_per_set", 1))
    starts = find_itr_start_pages(pdf_path, preset, doc=doc)
    if not starts:
        starts = [1]
        logs.append("WARN: 未找到 Page1 标记，默认从第1页作为ITR开始")
    set_start = starts[0]

    # 画出 Page1 标记（红框）：方便检查“拆分每套ITR起始页”的定位是否正确
    mark_re = compile_re(preset.get("page1_mark_regex", ""), DEFAULT_PAGE1_MARK_RE)
    for pi in range(len(doc)):
        page = doc[pi]
        t = page.get_text("text")
        if not t:
            continue
        if not mark_re.search(t):
            continue
        # 逐个匹配片段，尽量定位到页面上的文本位置
        for mm in mark_re.finditer(t):
            token = (mm.group(0) or "").strip()
            if not token:
                continue
            # search_for 按字符串查找位置；如果原文含多空格，这里做一个“压缩空格”兜底
            rects = page.search_for(token)
            if not rects:
                token2 = re.sub(r"\s+", " ", token)
                rects = page.search_for(token2)
            for rct in rects:
                page.draw_rect(rct, color=(0, 0, 1), width=1.2)
                logs.append(f"OK: Page1 标记页 -> page={pi + 1} text='{token}'")

    for f in fields:
        name = f.get("name", "")
        scope = f.get("page_scope", [1]) or [1]
        side = (f.get("pdf_label_side", "") or "").upper()

        for rel in scope:
            try:
                rel_i = int(rel)
            except Exception:
                rel_i = 1
            if rel_i < 1 or rel_i > pages_per_set:
                continue

            page_index = (set_start - 1) + (rel_i - 1)
            if page_index < 0 or page_index >= len(doc):
                continue
            page = doc[page_index]

            line_segments = collect_line_segments(page)
            rects = search_label_rect(page, f)
            label_rect = pick_label_rect_for_side(rects, side)
            if not label_rect:
                logs.append(f"[MISS] {name} (p{rel_i}): label not found")
                continue

            page.draw_rect(label_rect, color=(0, 0, 1), width=1)  # 蓝框 label
            cell_rect = find_cell_right_of_label(page, line_segments, label_rect)
            if not cell_rect:
                logs.append(f"[WARN] {name} (p{rel_i}): label found but cell not found")
                continue

            page.draw_rect(cell_rect, color=(1, 0, 0), width=1)  # 红框：可填写方框(cell)
            logs.append(f"[OK] {name} (p{rel_i})")

    match_cfg = preset.get("match", {})
    key_norm = norm_text(match_cfg.get("key_name", "TAG"))
    tag_direction = (match_cfg.get("tag_direction", "RIGHT") or "RIGHT").upper()
    page = doc[set_start - 1]
    value_regex = (match_cfg.get("pdf_extract_regex", "") or "").strip() or DEFAULT_VALUE_REGEX
    tag_text, tag_debug = extract_tag_by_cell_adjacency(page, key_norm, tag_direction, value_regex)
    key_cell_rect = tag_debug.get("key_cell_rect")
    value_cell_rect = tag_debug.get("value_cell_rect")
    if key_cell_rect:
        page.draw_rect(key_cell_rect, color=(0, 0, 1), width=1.2)
    if value_cell_rect:
        page.draw_rect(value_cell_rect, color=(1, 0, 0), width=1.2)
    value_raw_preview = (tag_debug.get("value_cell_text_raw") or "")[:200]
    chosen_line = tag_debug.get("chosen_line")
    vcover_count = tag_debug.get("vcover_count")
    hcover_count = tag_debug.get("hcover_count")
    error = tag_debug.get("error", "")
    logs.append(
        f"TAG pdf={os.path.basename(pdf_path)} start_page=P1 anchor_norm={key_norm} "
        f"key_cell_rect={key_cell_rect} key_cell_text_raw=\"{tag_debug.get('key_cell_text_raw', '')}\" "
        f"key_cell_text_norm=\"{tag_debug.get('key_cell_text_norm', '')}\" direction={tag_direction} "
        f"vcover_count={vcover_count} hcover_count={hcover_count} chosen_line={chosen_line} "
        f"value_cell_rect={value_cell_rect} value_cell_text_raw_preview=\"{value_raw_preview}\" "
        f"value_regex=\"{value_regex}\" tag_pick=\"{tag_text}\" error=\"{error}\" "
        "tag_source=CELL_ADJACENT"
    )

    base = os.path.basename(pdf_path)
    ts = batch_id()
    out_name = f"{os.path.splitext(base)[0]}__test_{preset.get('preset_name', 'preset')}__{ts}.pdf"
    batch_dir = os.path.join(OUTPUT_TEST_ROOT, ts)
    os.makedirs(batch_dir, exist_ok=True)
    out_path = os.path.join(batch_dir, out_name)
    doc.save(out_path)
    doc.close()

    log_path = os.path.join(batch_dir, f"{os.path.splitext(out_name)[0]}.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(f"Preset: {preset.get('preset_name', '')}\nPDF: {base}\nTime: {now_iso()}\n\n")
        for line in logs:
            f.write(line + "\n")

    return out_path, logs


@dataclass
class ITRItem:
    pdf_file: str
    set_start_page_1based: int
    key_pdf: str
    match_status: str
    excel_key: str
    sheet_name: str
    filled: Dict[str, str]
    tag_source: str


def write_one_itr(
    doc: fitz.Document,
    set_start_page_1based: int,
    preset: dict,
    filled: Dict[str, str],
    field_rect_cache: dict,
):
    text_cfg = preset.get("text", {})
    pages_per_set = int(preset.get("itr_pages_per_set", 1))

    for f in preset.get("fields", []):
        name = f.get("name", "")
        if not name:
            continue
        value = filled.get(name, "")
        if not str(value).strip():
            continue

        scope = f.get("page_scope", [1]) or [1]
        side = (f.get("pdf_label_side", "") or "").upper()

        for rel in scope:
            try:
                rel_i = int(rel)
            except Exception:
                rel_i = 1
            if rel_i < 1 or rel_i > pages_per_set:
                continue
            page_index = (set_start_page_1based - 1) + (rel_i - 1)
            if page_index < 0 or page_index >= len(doc):
                continue

            page = doc[page_index]
            cache_key = (page_index, name)
            if cache_key in field_rect_cache:
                label_rect, cell_rect = field_rect_cache[cache_key]
            else:
                line_segments = collect_line_segments(page)
                rects = search_label_rect(page, f)
                label_rect = pick_label_rect_for_side(rects, side)
                if not label_rect:
                    continue
                cell_rect = find_cell_right_of_label(page, line_segments, label_rect)
                if not cell_rect:
                    continue
                field_rect_cache[cache_key] = (label_rect, cell_rect)

            fit_text_to_box(page, cell_rect, str(value), text_cfg)


def simple_input(parent, title: str, prompt: str, default: str = "") -> str:
    win = tk.Toplevel(parent)
    win.title(title)
    win.geometry("480x170")
    win.resizable(False, False)

    ttk.Label(win, text=prompt).pack(anchor="w", padx=10, pady=(12, 6))
    ent = ttk.Entry(win, width=60)
    ent.pack(padx=10, pady=6)
    ent.insert(0, default)
    ent.focus()

    out = {"val": ""}

    def ok():
        out["val"] = ent.get()
        win.destroy()

    def cancel():
        out["val"] = ""
        win.destroy()

    btns = ttk.Frame(win)
    btns.pack(fill="x", padx=10, pady=10)
    ttk.Button(btns, text="确定", command=ok).pack(side="right")
    ttk.Button(btns, text="取消", command=cancel).pack(side="right", padx=8)

    win.grab_set()
    parent.wait_window(win)
    return out["val"]


class ITRAutofillTab(ttk.Frame):
    def __init__(self, master):
        super().__init__(master)

        self.global_cfg = load_global_config()
        self.match_memory = load_match_memory()
        self.tag_choice_memory = load_json_safe(TAG_CHOICE_MEMORY_PATH, {})
        self._loaded_preset_name: Optional[str] = None

        self.excel_path: Optional[str] = None
        self.pdf_paths: List[str] = []
        self.excel_index: Optional[dict] = None

        self.items: List[ITRItem] = []
        self.current_idx: int = -1
        self.tag_choice_cache: Dict[str, dict] = {}

        self._active_entry: Optional[tk.Entry] = None
        self._active_item_id: Optional[str] = None

        self._export_thread: Optional[threading.Thread] = None
        self._q: "queue.Queue[tuple]" = queue.Queue()
        self._preview_thread: Optional[threading.Thread] = None
        self._preview_q: "queue.Queue[tuple]" = queue.Queue()
        self._presets_window: Optional[tk.Toplevel] = None
        self.preset_confirmed: bool = False

        self._build_ui()
        self._build_presets_window()
        self._load_initial_preset()
        self._update_main_preset_status()

    def _build_presets_window(self):
        win = tk.Toplevel(self)
        win.title("预设管理")
        win.geometry("1380x860")
        win.withdraw()
        root = self.winfo_toplevel()
        win.transient(root)
        win.lift()
        win.attributes("-topmost", True)
        win.after(200, lambda: win.attributes("-topmost", False))
        win.focus_force()

        def on_close():
            win.withdraw()

        win.protocol("WM_DELETE_WINDOW", on_close)
        self._build_presets_ui(win)
        self._presets_window = win

    def open_presets_window(self):
        if self._presets_window and self._presets_window.winfo_exists():
            self._presets_window.deiconify()
            self._presets_window.lift()
            self._presets_window.focus_force()
        else:
            self._build_presets_window()
            if self._presets_window:
                self._presets_window.deiconify()
                self._presets_window.lift()
                self._presets_window.focus_force()

    def open_pdf_test_folder(self):
        path = OUTPUT_TEST_ROOT
        os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开测试文件夹: {e}")

    def open_output_folder(self):
        path = os.path.join(BASE_DIR, "output", MODULE_NAME)
        os.makedirs(path, exist_ok=True)
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开输出文件夹: {e}")

    def open_filled_folder(self):
        path = Path(BASE_DIR) / "output" / MODULE_NAME / "filled"
        open_in_file_explorer(path)

    # ------------------ UI ------------------
    def _build_ui(self):
        self._build_main_tab()

    # ------------------ Presets Tab ------------------
    def _build_presets_ui(self, root):
        pan = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashwidth=12, sashrelief=tk.RAISED, bd=1, relief=tk.GROOVE)
        pan.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(pan)
        right = ttk.Frame(pan)
        pan.add(left, minsize=280)
        pan.add(right, minsize=920)

        # 左：预设列表
        ttk.Label(left, text="预设列表").pack(anchor="w")
        lf = ttk.Frame(left)
        lf.pack(fill="both", expand=True)

        self.preset_list = tk.Listbox(lf, height=20)
        self.preset_list.pack(side="left", fill="both", expand=True)
        self.preset_list.bind("<<ListboxSelect>>", self.on_select_preset)

        sb = ttk.Scrollbar(lf, orient="vertical", command=self.preset_list.yview)
        sb.pack(side="right", fill="y")
        self.preset_list.config(yscrollcommand=sb.set)

        btns = ttk.Frame(left)
        btns.pack(fill="x", pady=10)
        ttk.Button(btns, text="新建", command=self.preset_new).pack(fill="x", pady=2)
        ttk.Button(btns, text="删除", command=self.preset_delete).pack(fill="x", pady=2)
        ttk.Button(btns, text="设为当前使用", command=self.preset_set_active).pack(fill="x", pady=2)

        # 右：预设信息
        meta = ttk.LabelFrame(right, text="预设信息")
        meta.pack(fill="x")

        row1 = ttk.Frame(meta)
        row1.pack(fill="x", padx=10, pady=6)
        ttk.Label(row1, text="预设名").pack(side="left")
        self.ent_preset_name = ttk.Entry(row1, width=40)
        self.ent_preset_name.pack(side="left", padx=8)
        ttk.Label(row1, text="创建时间").pack(side="left", padx=(20, 4))
        self.lbl_created = ttk.Label(row1, text="-")
        self.lbl_created.pack(side="left")
        ttk.Label(row1, text="修改时间").pack(side="left", padx=(20, 4))
        self.lbl_updated = ttk.Label(row1, text="-")
        self.lbl_updated.pack(side="left")

        row2 = ttk.Frame(meta)
        row2.pack(fill="x", padx=10, pady=6)
        ttk.Label(row2, text="ITR每套页数").pack(side="left")
        self.ent_pages_per_set = ttk.Entry(row2, width=10)
        self.ent_pages_per_set.pack(side="left", padx=8)
        ttk.Label(row2, text="Excel表头行(从0开始)").pack(side="left", padx=(20, 4))
        self.ent_header_row = ttk.Entry(row2, width=10)
        self.ent_header_row.pack(side="left", padx=8)
        ttk.Label(row2, text="例：列名在Excel第3行→填2").pack(side="left", padx=(12, 0))

        row3 = ttk.Frame(meta)
        row3.pack(fill="x", padx=10, pady=6)
        ttk.Label(row3, text="Page1标记规则(正则)").pack(side="left")
        self.ent_page1_re = ttk.Entry(row3, width=70)
        self.ent_page1_re.pack(side="left", padx=8)
        ttk.Label(row3, text="(用于找每套ITR的第1页)").pack(side="left", padx=8)

        # Match Key
        mk = ttk.LabelFrame(right, text="Match Key（匹配键）配置")
        mk.pack(fill="x", pady=10)

        mk1 = ttk.Frame(mk)
        mk1.pack(fill="x", padx=10, pady=6)
        ttk.Label(mk1, text="Key 归一值(锚点，例：TAGNO)").pack(side="left")
        self.ent_key_name = ttk.Entry(mk1, width=20)
        self.ent_key_name.pack(side="left", padx=8)

        mk1b = ttk.Frame(mk)
        mk1b.pack(fill="x", padx=10, pady=6)
        ttk.Label(mk1b, text="Tag 值方向").pack(side="left", padx=(0, 4))
        self.tag_direction_var = tk.StringVar(value="RIGHT")
        self.tag_direction_menu = ttk.OptionMenu(
            mk1b, self.tag_direction_var, self.tag_direction_var.get(), *TAG_DIRECTION_OPTIONS
        )
        self.tag_direction_menu.pack(side="left")
        ttk.Label(mk1b, text="值提取正则").pack(side="left", padx=(20, 4))
        self.ent_pdf_key_re = ttk.Entry(mk1b, width=60)
        self.ent_pdf_key_re.pack(side="left", padx=8)

        mk2 = ttk.Frame(mk)
        mk2.pack(fill="x", padx=10, pady=6)
        ttk.Label(mk2, text="去除后缀(逗号分隔)").pack(side="left")
        self.ent_strip_suf = ttk.Entry(mk2, width=50)
        self.ent_strip_suf.pack(side="left", padx=8)
        ttk.Label(mk2, text="例：-EX,-JB  (匹配前删掉末尾后缀)").pack(side="left", padx=(12, 0))

        mk3 = ttk.Frame(mk)
        mk3.pack(fill="x", padx=10, pady=6)
        ttk.Label(mk3, text="Excel Key列候选(归一化,逗号分隔)").pack(side="left")
        self.ent_key_cols = ttk.Entry(mk3, width=70)
        self.ent_key_cols.pack(side="left", padx=8)

        mk4 = ttk.Frame(mk)
        mk4.pack(fill="x", padx=10, pady=6)
        self.var_enable_fuzzy = tk.BooleanVar(value=True)
        self.var_require_confirm = tk.BooleanVar(value=True)
        ttk.Checkbutton(mk4, text="启用模糊匹配", variable=self.var_enable_fuzzy).pack(side="left")
        ttk.Checkbutton(mk4, text="模糊匹配需确认", variable=self.var_require_confirm).pack(side="left", padx=20)

        # 字段映射表
        ff = ttk.LabelFrame(right, text="字段映射（可增删；双击单元格编辑；可多选后测试PDF定位）")
        ff.pack(fill="both", expand=True)

        toolrow = ttk.Frame(ff)
        toolrow.pack(fill="x", padx=10, pady=6)
        ttk.Button(toolrow, text="添加字段", command=self.field_add).pack(side="left")
        ttk.Button(toolrow, text="删除字段", command=self.field_delete).pack(side="left", padx=8)
        ttk.Button(toolrow, text="保存预设", command=self.preset_save).pack(side="right")
        ttk.Button(toolrow, text="另存为...", command=self.preset_save_as).pack(side="right", padx=8)
        ttk.Button(toolrow, text="测试PDF定位(画框)", command=self.field_test_pdf).pack(side="right", padx=8)
        ttk.Button(toolrow, text="打开测试文件夹", command=self.open_pdf_test_folder).pack(side="right", padx=8)

        cols = ("name", "pdf_label", "side", "page_scope", "source", "excel_col_norm", "const", "rule")
        self.field_tree = ttk.Treeview(ff, columns=cols, show="headings", selectmode="extended", height=16)
        for c in cols:
            self.field_tree.heading(c, text=c)
        self.field_tree.column("name", width=170, anchor="w")
        self.field_tree.column("pdf_label", width=230, anchor="w")
        self.field_tree.column("side", width=80, anchor="center")
        self.field_tree.column("page_scope", width=130, anchor="w")
        self.field_tree.column("source", width=90, anchor="center")
        self.field_tree.column("excel_col_norm", width=190, anchor="w")
        self.field_tree.column("const", width=140, anchor="w")
        self.field_tree.column("rule", width=150, anchor="w")
        self.field_tree.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        self.field_tree.bind("<Double-1>", self.on_field_edit)

        self._bind_preset_change_events()
        self.refresh_preset_list()

    def refresh_preset_list(self):
        self.preset_list.delete(0, tk.END)
        for n in list_presets():
            self.preset_list.insert(tk.END, n)

    def _bind_preset_change_events(self):
        entries = [
            self.ent_preset_name,
            self.ent_pages_per_set,
            self.ent_header_row,
            self.ent_page1_re,
            self.ent_key_name,
            self.ent_pdf_key_re,
            self.ent_strip_suf,
            self.ent_key_cols,
        ]
        for ent in entries:
            ent.bind("<KeyRelease>", lambda _e: self._mark_preset_modified())
        self.var_enable_fuzzy.trace_add("write", lambda *_: self._mark_preset_modified())
        self.tag_direction_var.trace_add("write", lambda *_: self._mark_preset_modified())
        self.var_require_confirm.trace_add("write", lambda *_: self._mark_preset_modified())

    def _mark_preset_modified(self):
        if self.preset_confirmed:
            self.preset_confirmed = False
            self.status.config(text="预设已修改，需重新确认")
        self._update_main_preset_status()

    def _set_preset_confirmed(self, confirmed: bool):
        self.preset_confirmed = confirmed
        if confirmed:
            self.status.config(text="预设已确认")
        else:
            self.status.config(text="预设已修改，需重新确认")
        self._update_main_preset_status()

    def _load_initial_preset(self):
        names = list_presets()
        if not names:
            d = default_preset()
            save_preset(d["preset_name"], d)
            names = list_presets()

        active = self.global_cfg.get("active_preset", "")
        if active and active in names:
            self.load_preset_into_editor(active)
            idx = names.index(active)
            self.preset_list.selection_set(idx)
            self.preset_list.see(idx)
        else:
            self.load_preset_into_editor(names[0])
            self.preset_list.selection_set(0)
        self._set_preset_confirmed(False)

    def on_select_preset(self, _evt=None):
        sel = self.preset_list.curselection()
        if not sel:
            return
        self.load_preset_into_editor(self.preset_list.get(sel[0]))

    def load_preset_into_editor(self, name: str):
        d = load_preset(name)
        if not d:
            messagebox.showerror("错误", f"无法加载预设：{name}")
            return
        self._loaded_preset_name = name
        save_preset(name, d)  # 旧预设升级
        d = load_preset(name) or d

        if not hasattr(self, "ent_preset_name"):
            return

        self.ent_preset_name.delete(0, tk.END)
        self.ent_preset_name.insert(0, name)
        self.lbl_created.config(text=d.get("created_at", "-"))
        self.lbl_updated.config(text=d.get("updated_at", "-"))

        self.ent_pages_per_set.delete(0, tk.END)
        self.ent_pages_per_set.insert(0, str(d.get("itr_pages_per_set", 4)))
        self.ent_header_row.delete(0, tk.END)
        self.ent_header_row.insert(0, str(d.get("excel", {}).get("header_row", 0)))
        self.ent_page1_re.delete(0, tk.END)
        self.ent_page1_re.insert(0, str(d.get("page1_mark_regex", default_preset()["page1_mark_regex"])))

        m = d.get("match", {})
        self.ent_key_name.delete(0, tk.END)
        self.ent_key_name.insert(0, str(m.get("key_name", "TAG")))
        self.ent_pdf_key_re.delete(0, tk.END)
        self.ent_pdf_key_re.insert(0, str(m.get("pdf_extract_regex", default_preset()["match"]["pdf_extract_regex"])))
        self.tag_direction_var.set(str(m.get("tag_direction", "RIGHT")).upper())
        self.ent_strip_suf.delete(0, tk.END)
        self.ent_strip_suf.insert(0, ",".join(m.get("strip_suffixes", ["-EX"])))
        self.ent_key_cols.delete(0, tk.END)
        self.ent_key_cols.insert(0, ",".join(m.get("excel_key_col_candidates_norm", [])))
        self.var_enable_fuzzy.set(bool(m.get("enable_fuzzy", True)))
        self.var_require_confirm.set(bool(m.get("fuzzy_require_confirm", True)))

        self._render_fields_tree(d.get("fields", []))
        self._set_preset_confirmed(False)
        self._update_main_preset_status()

    def _render_fields_tree(self, fields: List[dict]):
        self.field_tree.delete(*self.field_tree.get_children())
        for f in fields:
            self.field_tree.insert(
                "",
                tk.END,
                values=(
                    f.get("name", ""),
                    f.get("pdf_label", ""),
                    (f.get("pdf_label_side", "") or "").upper(),
                    ",".join(str(x) for x in (f.get("page_scope", [1]) or [1])),
                    (f.get("source", "") or "MANUAL").upper(),
                    f.get("excel_col_norm", ""),
                    f.get("const_value", ""),
                    (f.get("rule", "") or "").upper(),
                ),
            )

    def _read_fields_tree(self) -> List[dict]:
        res = []
        for iid in self.field_tree.get_children():
            name, pdf_label, side, page_scope, source, excel_col_norm, const_val, rule = self.field_tree.item(
                iid, "values"
            )
            scope_list = []
            for p in str(page_scope).split(","):
                p = p.strip()
                if not p:
                    continue
                try:
                    scope_list.append(int(p))
                except Exception:
                    pass
            if not scope_list:
                scope_list = [1]
            res.append(
                {
                    "name": str(name).strip(),
                    "pdf_label": str(pdf_label).strip(),
                    "pdf_label_side": str(side).strip().upper(),
                    "page_scope": scope_list,
                    "source": (str(source).strip().upper() or "MANUAL"),
                    "excel_col_norm": str(excel_col_norm).strip(),
                    "const_value": str(const_val),
                    "rule": str(rule).strip().upper(),
                }
            )
        return res

    def preset_new(self):
        """新建一个预设并载入到编辑区。"""
        base = "NewPreset"
        names = set(list_presets())
        i = 1
        name = base
        while name in names:
            i += 1
            name = f"{base}_{i}"
        d = default_preset()
        d["preset_name"] = name
        d["created_at"] = now_iso()
        d["updated_at"] = now_iso()
        save_preset(name, d)
        self.refresh_preset_list()
        idx = list_presets().index(name)
        self.preset_list.selection_clear(0, tk.END)
        self.preset_list.selection_set(idx)
        self.preset_list.see(idx)
        self.load_preset_into_editor(name)
        self._set_preset_confirmed(False)

    def preset_delete(self):
        sel = self.preset_list.curselection()
        if not sel:
            return
        name = self.preset_list.get(sel[0])
        if not messagebox.askyesno("确认", f"删除预设：{name} ？"):
            return
        try:
            os.remove(preset_path(name))
        except Exception as e:
            messagebox.showerror("错误", f"删除失败：{e}")
            return
        self.refresh_preset_list()
        names = list_presets()
        if names:
            self.load_preset_into_editor(names[0])
            self.preset_list.selection_set(0)
        else:
            self._load_initial_preset()
        self._set_preset_confirmed(False)

    def _collect_preset_from_editor(self) -> Tuple[str, dict]:
        d = load_preset(self.ent_preset_name.get().strip()) or default_preset()
        name = self.ent_preset_name.get().strip() or "UnnamedPreset"
        try:
            d["itr_pages_per_set"] = int(self.ent_pages_per_set.get().strip() or "4")
        except Exception:
            d["itr_pages_per_set"] = 4
        try:
            d.setdefault("excel", {})["header_row"] = int(self.ent_header_row.get().strip() or "0")
        except Exception:
            d.setdefault("excel", {})["header_row"] = 0

        d["page1_mark_regex"] = self.ent_page1_re.get().strip() or default_preset()["page1_mark_regex"]

        d.setdefault("match", {})
        d["match"]["key_name"] = self.ent_key_name.get().strip() or "TAG"
        d["match"]["pdf_extract_regex"] = self.ent_pdf_key_re.get().strip() or default_preset()["match"][
            "pdf_extract_regex"
        ]
        d["match"]["tag_direction"] = (self.tag_direction_var.get() or "RIGHT").upper()
        d["match"]["strip_suffixes"] = [x.strip() for x in self.ent_strip_suf.get().split(",") if x.strip()]
        d["match"]["excel_key_col_candidates_norm"] = [x.strip() for x in self.ent_key_cols.get().split(",") if x.strip()]
        d["match"]["enable_fuzzy"] = bool(self.var_enable_fuzzy.get())
        d["match"]["fuzzy_require_confirm"] = bool(self.var_require_confirm.get())
        d["fields"] = self._read_fields_tree()
        return name, d

    def preset_save(self):
        """保存当前预设（如果用户在右侧把预设名改了，则视为‘重命名’，不会留下旧文件）。"""
        name, d = self._collect_preset_from_editor()
        old = getattr(self, "_loaded_preset_name", None)

        save_preset(name, d)

        if old and old != name:
            try:
                old_path = preset_path(old)
                if os.path.exists(old_path):
                    os.remove(old_path)
            except Exception:
                pass

        self._loaded_preset_name = name
        self.refresh_preset_list()
        self.load_preset_into_editor(name)
        self._set_preset_confirmed(True)
        messagebox.showinfo("完成", f"已保存预设：{name}")

    def preset_save_as(self):
        name, d = self._collect_preset_from_editor()
        new = simple_input(self, "另存为", "请输入新预设名：", default=name + "_copy")
        if not new:
            return
        new = new.strip()
        if not new:
            return
        if os.path.exists(preset_path(new)) and not messagebox.askyesno("提示", f"预设 {new} 已存在，是否覆盖？"):
            return
        save_preset(new, d)
        self.refresh_preset_list()
        self.load_preset_into_editor(new)
        self._set_preset_confirmed(True)
        messagebox.showinfo("完成", f"已另存为：{new}")

    def preset_set_active(self):
        name = self.ent_preset_name.get().strip()
        if not name:
            return
        self.global_cfg["active_preset"] = name
        save_global_config(self.global_cfg)
        self._set_preset_confirmed(True)
        messagebox.showinfo("完成", f"当前使用预设：{name}")
        self._update_main_preset_status()
        if self._presets_window and self._presets_window.winfo_exists():
            self._presets_window.withdraw()

    def preset_apply(self):
        """兼容可能存在的“应用”入口：等同于设为当前使用。"""
        self.preset_set_active()

    def preset_edit(self):
        """预留的编辑入口，保证外部调用不报错。"""
        messagebox.showinfo("提示", "当前已在右侧面板直接编辑预设字段。")

    def on_field_edit(self, event):
        iid = self.field_tree.identify_row(event.y)
        col = self.field_tree.identify_column(event.x)
        if not iid or not col:
            return
        col_index = int(col.replace("#", "")) - 1
        cols = ("name", "pdf_label", "side", "page_scope", "source", "excel_col_norm", "const", "rule")
        if col_index < 0 or col_index >= len(cols):
            return
        x, y, w, h = self.field_tree.bbox(iid, col)
        old = self.field_tree.item(iid, "values")[col_index]
        key = cols[col_index]

        # source -> 下拉
        if key == "source":
            cb = ttk.Combobox(self.field_tree, values=SOURCE_OPTIONS, state="readonly")
            cb.place(x=x, y=y, width=w, height=h)
            cb.set((old or "MANUAL").upper())
            cb.focus()

            def commit(_=None):
                vals = list(self.field_tree.item(iid, "values"))
                vals[col_index] = cb.get().upper()
                self.field_tree.item(iid, values=vals)
                cb.destroy()
                self._mark_preset_modified()

            cb.bind("<<ComboboxSelected>>", commit)
            cb.bind("<FocusOut>", commit)
            return

        # side -> 下拉
        if key == "side":
            cb = ttk.Combobox(self.field_tree, values=SIDE_OPTIONS, state="readonly")
            cb.place(x=x, y=y, width=w, height=h)
            cb.set((old or "").upper())
            cb.focus()

            def commit(_=None):
                vals = list(self.field_tree.item(iid, "values"))
                vals[col_index] = cb.get().upper()
                self.field_tree.item(iid, values=vals)
                cb.destroy()
                self._mark_preset_modified()

            cb.bind("<<ComboboxSelected>>", commit)
            cb.bind("<FocusOut>", commit)
            return

        # rule -> 下拉
        if key == "rule":
            cb = ttk.Combobox(self.field_tree, values=RULE_OPTIONS, state="readonly")
            cb.place(x=x, y=y, width=w, height=h)
            cb.set((old or "").upper())
            cb.focus()

            def commit(_=None):
                vals = list(self.field_tree.item(iid, "values"))
                vals[col_index] = cb.get().upper()
                self.field_tree.item(iid, values=vals)
                cb.destroy()
                self._mark_preset_modified()

            cb.bind("<<ComboboxSelected>>", commit)
            cb.bind("<FocusOut>", commit)
            return

        # 其他 -> 普通输入框
        ent = ttk.Entry(self.field_tree)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, old)
        ent.focus()

        def commit(_=None):
            vals = list(self.field_tree.item(iid, "values"))
            val = ent.get()
            if key in ("side", "source", "rule"):
                val = val.upper().strip()
            vals[col_index] = val
            self.field_tree.item(iid, values=vals)
            ent.destroy()
            self._mark_preset_modified()

        ent.bind("<Return>", commit)
        ent.bind("<FocusOut>", commit)
        ent.bind("<Escape>", commit)

    def field_add(self):
        self.field_tree.insert("", tk.END, values=("NewField", "Label", "", "1", "MANUAL", "", "", ""))
        self._mark_preset_modified()

    def field_delete(self):
        sel = self.field_tree.selection()
        for iid in sel:
            self.field_tree.delete(iid)
        if sel:
            self._mark_preset_modified()

    def field_test_pdf(self):
        name, preset = self._collect_preset_from_editor()
        preset["preset_name"] = name

        pdf = filedialog.askopenfilename(title="选择要测试的 PDF", filetypes=[("PDF files", "*.pdf")])
        if not pdf:
            return

        fields = preset.get("fields", [])
        sel = self.field_tree.selection()
        if sel:
            selected_names = {self.field_tree.item(iid, "values")[0] for iid in sel}
            fields = [f for f in fields if f.get("name") in selected_names]

        out_pdf, _logs = pdf_position_test(pdf, preset, fields)
        if out_pdf:
            messagebox.showinfo(
                "完成",
                f"已生成测试PDF（蓝框label 红框cell）：\n{os.path.basename(out_pdf)}\n目录：output/itr_autofill/test/",
            )
        else:
            messagebox.showerror("错误", "测试失败")

    # ------------------ Main Tab ------------------
    def _build_main_tab(self):
        root = self

        top = ttk.Frame(root)
        top.pack(fill="x", padx=10, pady=10)
        self.lbl_active = ttk.Label(top, text="当前预设：未设置")
        self.lbl_active.pack(side="left")
        ttk.Button(top, text="打开预设管理", command=self.open_presets_window).pack(side="right")

        filebar = ttk.Frame(root)
        filebar.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(filebar, text="选择 Excel", command=self.pick_excel).pack(side="left")
        self.lbl_excel = ttk.Label(filebar, text="未选择")
        self.lbl_excel.pack(side="left", padx=8)
        ttk.Button(filebar, text="选择 PDF(可多选)", command=self.pick_pdfs).pack(side="left", padx=10)
        self.lbl_pdfs = ttk.Label(filebar, text="未选择")
        self.lbl_pdfs.pack(side="left", padx=8)
        self.btn_parse = ttk.Button(filebar, text="解析&预填", command=self.run_preview)
        self.btn_parse.pack(side="left", padx=10)
        ttk.Button(filebar, text="管理映射记忆", command=self.open_memory_manager).pack(side="left", padx=10)

        # 进度条（导出用）
        progbar = ttk.Frame(root)
        progbar.pack(fill="x", padx=10)
        self.export_progress = ttk.Progressbar(progbar, mode="determinate")
        self.export_progress.pack(fill="x")
        self.export_progress["value"] = 0

        mid = tk.PanedWindow(root, orient=tk.HORIZONTAL, sashwidth=14, sashrelief=tk.RAISED, bd=1, relief=tk.GROOVE)
        mid.pack(fill="both", expand=True, padx=10, pady=10)

        left = ttk.Frame(mid)
        ttk.Label(left, text="ITR 列表（按套）").pack(anchor="w")
        lf = ttk.Frame(left)
        lf.pack(fill="both", expand=True)
        self.listbox = tk.Listbox(lf)
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.bind("<<ListboxSelect>>", self.on_select_item)
        sb_y = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        sb_y.pack(side="right", fill="y")
        self.listbox.config(yscrollcommand=sb_y.set)
        sb_x = ttk.Scrollbar(left, orient="horizontal", command=self.listbox.xview)
        sb_x.pack(fill="x")
        self.listbox.config(xscrollcommand=sb_x.set)

        right = ttk.Frame(mid)
        ttk.Label(right, text="字段预填/人工修改（双击值列编辑）").pack(anchor="w")
        self.tree = ttk.Treeview(right, columns=("key", "value"), show="headings", height=22)
        self.tree.heading("key", text="字段")
        self.tree.heading("value", text="值")
        self.tree.column("key", width=220, anchor="w")
        self.tree.column("value", width=920, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<Double-1>", self.on_edit_cell)

        mid.add(left, minsize=380)
        mid.add(right, minsize=820)

        bottom = ttk.Frame(root)
        bottom.pack(fill="x", padx=10, pady=10)
        self.btn_open_output = ttk.Button(bottom, text="打开导出文件夹", command=self.open_filled_folder)
        self.btn_open_output.pack(side="right")
        self.btn_export = ttk.Button(bottom, text="导出填好的PDF + report.xlsx", command=self.export_all_async)
        self.btn_export.pack(side="right", padx=10)
        self.btn_save_edits = ttk.Button(bottom, text="保存当前修改", command=self.save_current_edits)
        self.btn_save_edits.pack(side="right")
        self.status = ttk.Label(bottom, text="就绪")
        self.status.pack(side="left")

    def _update_main_preset_status(self):
        active = self.global_cfg.get("active_preset", "")
        if not active:
            self.lbl_active.config(text="当前预设：未设置（请到预设管理中设为当前使用）")
            self.btn_parse.config(state="disabled")
            self.btn_export.config(state="disabled")
            return
        p = load_preset(active)
        if not p:
            self.lbl_active.config(text="当前预设：加载失败（请重新设定）")
            self.btn_parse.config(state="disabled")
            self.btn_export.config(state="disabled")
            return
        if not self.preset_confirmed:
            self.lbl_active.config(text="当前预设：未确认（请保存或设为当前使用）")
            self.btn_parse.config(state="disabled")
            self.btn_export.config(state="disabled")
            return
        self.lbl_active.config(
            text=f"当前预设：{active} | 创建：{p.get('created_at', '-')} | 修改：{p.get('updated_at', '-')}"
        )
        m = p.get("match", {})
        ok = bool(m.get("key_name", "").strip()) and bool(p.get("fields", []))
        self.btn_parse.config(state=("normal" if ok else "disabled"))

    def pick_excel(self):
        path = filedialog.askopenfilename(
            title="选择 Excel", filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if path:
            if not path.lower().endswith(".xlsx"):
                messagebox.showerror("错误", "Excel 读取仅支持 .xlsx 格式，请用 Excel 另存为 .xlsx 后再导入")
                return
            self.excel_path = path
            self.lbl_excel.config(text=os.path.basename(path))

    def pick_pdfs(self):
        paths = filedialog.askopenfilenames(
            title="选择 PDF（可多选）", filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        if paths:
            self.pdf_paths = list(paths)
            self.lbl_pdfs.config(text=f"{len(self.pdf_paths)} 个PDF")

    def gui_tag_candidate_chooser(self, pdf_name: str, candidates: List[str]) -> Tuple[str, bool]:
        win = tk.Toplevel(self)
        win.title("Tag 候选选择")
        win.geometry("600x420")

        ttk.Label(win, text=f"PDF: {pdf_name}").pack(anchor="w", padx=10, pady=(10, 2))
        ttk.Label(win, text="请选择一个 Tag 候选：").pack(anchor="w", padx=10, pady=(0, 8))

        lb = tk.Listbox(win, height=12, width=80)
        lb.pack(fill="both", expand=True, padx=10, pady=6)
        for v in candidates:
            lb.insert(tk.END, v)

        remember_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            win,
            text="记住该选择（同模板下次自动选中）",
            variable=remember_var,
        ).pack(anchor="w", padx=10, pady=(0, 6))

        chosen = {"val": "", "remember": False}

        def pick():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("提示", "请先选择一个候选")
                return
            chosen["val"] = candidates[sel[0]]
            chosen["remember"] = bool(remember_var.get())
            win.destroy()

        ttk.Button(win, text="确定", command=pick).pack(pady=8)
        win.grab_set()
        self.wait_window(win)

        return chosen["val"], chosen["remember"]

    def gui_fuzzy_chooser(self, key_pdf: str, short_key: str, candidates: List[str]) -> str:
        win = tk.Toplevel(self)
        win.title("模糊匹配确认")
        win.geometry("860x520")

        ttk.Label(win, text=f"PDF Key: {key_pdf}").pack(anchor="w", padx=10, pady=(10, 2))
        ttk.Label(win, text=f"短Key(用于记忆): {short_key}").pack(anchor="w", padx=10, pady=(0, 10))
        ttk.Label(win, text="请选择一个 Excel Key：").pack(anchor="w", padx=10)

        lb = tk.Listbox(win, height=14, width=110)
        lb.pack(fill="both", expand=True, padx=10, pady=8)
        for c in candidates:
            lb.insert(tk.END, c)

        remember_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            win,
            text="记住该映射（下次遇到同样短Key不再提示）",
            variable=remember_var,
        ).pack(anchor="w", padx=10, pady=(0, 10))

        chosen = {"val": "", "remember": False}

        def pick():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning("提示", "请先点选一条候选")
                return
            chosen["val"] = candidates[sel[0]]
            chosen["remember"] = bool(remember_var.get())
            win.destroy()

        ttk.Button(win, text="确定", command=pick).pack(pady=8)
        win.grab_set()
        self.wait_window(win)

        if chosen["val"] and chosen["remember"]:
            self.match_memory[short_key] = chosen["val"]
            save_match_memory(self.match_memory)

        return chosen["val"]

    def run_preview(self):
        if self._preview_thread and self._preview_thread.is_alive():
            messagebox.showinfo("提示", "正在解析，请稍等…")
            return

        if not self.preset_confirmed:
            messagebox.showwarning("提示", "请先确认预设")
            return

        active = self.global_cfg.get("active_preset", "")
        preset = load_preset(active) if active else None
        if not preset:
            messagebox.showwarning("提示", "请先在预设管理中设定当前预设")
            return
        if not self.excel_path:
            messagebox.showwarning("提示", "请先选择 Excel")
            return
        if not self.pdf_paths:
            messagebox.showwarning("提示", "请先选择 PDF")
            return

        self.btn_parse.config(state="disabled")
        self.btn_export.config(state="disabled")
        self.btn_save_edits.config(state="disabled")
        self.status.config(text="正在建立 Excel 索引...")
        self.export_progress["value"] = 0
        self.update_idletasks()

        self._preview_thread = threading.Thread(
            target=self._preview_worker,
            args=(preset,),
            daemon=True,
        )
        self._preview_thread.start()
        self.after(120, self._poll_preview_queue)

    def _tag_cache_key(self, pdf_path: str, start_page: int, tag_mode: str) -> str:
        return f"{pdf_path}::p{start_page}::{tag_mode}"

    def _resolve_itr_tag(self, doc: fitz.Document, pdf_path: str, start_page: int, preset: dict) -> Tuple[str, str]:
        match_cfg = preset.get("match", {})
        cache_key = self._tag_cache_key(pdf_path, start_page, "CELL_A")
        cached = self.tag_choice_cache.get(cache_key)
        if cached and cached.get("value"):
            return cached["value"], cached.get("source", "regex_manual")
        key_name = match_cfg.get("key_name", "TAG")
        key_norm = norm_text(key_name)
        direction = (match_cfg.get("tag_direction", "RIGHT") or "RIGHT").upper()
        page1 = doc[start_page - 1]

        value_regex = (match_cfg.get("pdf_extract_regex", "") or "").strip() or DEFAULT_VALUE_REGEX
        print(f"pdf={os.path.basename(pdf_path)} start_page=P{start_page} anchor_norm={key_norm}")
        value_normed, debug = extract_tag_by_cell_adjacency(page1, key_norm, direction, value_regex)
        key_cell_rect = debug.get("key_cell_rect")
        key_cell_text_raw = debug.get("key_cell_text_raw", "")
        key_cell_text_norm = debug.get("key_cell_text_norm", "")
        value_cell_rect = debug.get("value_cell_rect")
        value_raw = debug.get("value_cell_text_raw", "") or ""
        value_raw_preview = (value_raw or "")[:200]
        vcover_count = debug.get("vcover_count")
        hcover_count = debug.get("hcover_count")
        chosen_line = debug.get("chosen_line")
        error = debug.get("error", "")
        print(
            f"key_cell_rect={key_cell_rect} key_cell_text_raw=\"{key_cell_text_raw}\" "
            f"key_cell_text_norm=\"{key_cell_text_norm}\" direction={direction} "
            f"vcover_count={vcover_count} hcover_count={hcover_count} chosen_line={chosen_line} "
            f"value_cell_rect={value_cell_rect} value_cell_text_raw_preview=\"{value_raw_preview}\" "
            f"value_regex=\"{value_regex}\" error=\"{error}\""
        )
        if not value_normed:
            if debug.get("error") == "anchor_cell_not_found":
                print(f"未找到 key_norm={key_norm} 的单元格（严格等值匹配）")
            elif debug.get("error") == "regex_no_match":
                print(f"value_cell 有文本，但 regex 截取失败；value_raw={value_raw_preview}; regex={value_regex}")
            else:
                print(f"找到 key_norm 单元格，但 direction={direction} 的相邻单元格不存在/为空")
            return "", "missing"

        tag_value = norm_key_value(value_normed)
        for suf in match_cfg.get("strip_suffixes", []):
            suf_up = norm_key_value(suf)
            if suf_up and tag_value.endswith(suf_up):
                tag_value = tag_value[: -len(suf_up)]
        print(
            f"value_cell_rect={value_cell_rect} value_cell_text_raw_preview=\"{value_raw_preview}\" "
            f"value_regex=\"{value_regex}\" tag_pick=\"{tag_value}\" error=\"{error}\" "
            "tag_source=CELL_ADJACENT"
        )
        self.tag_choice_cache[cache_key] = {"value": tag_value, "source": "CELL_ADJACENT"}
        return tag_value, "CELL_ADJACENT"

    def _preview_worker(self, preset: dict):
        try:
            excel_index = build_excel_index(self.excel_path, preset)
        except Exception as e:
            self._preview_q.put(("error", f"Excel 读取失败：{e}"))
            return

        items: List[ITRItem] = []
        total_files = len(self.pdf_paths)
        for idx, pdf_path in enumerate(self.pdf_paths):
            pdf_name = os.path.basename(pdf_path)
            try:
                doc = fitz.open(pdf_path)
            except Exception:
                done_files = idx + 1
                self._preview_q.put(("progress", done_files, total_files, pdf_name))
                continue

            starts = find_itr_start_pages(pdf_path, preset, doc=doc)
            if not starts:
                starts = [1]

            pages_per_set = int(preset.get("itr_pages_per_set", 1))
            valid_starts = [s for s in starts if s + pages_per_set - 1 <= len(doc)] or starts[:1]
            starts_set = set(valid_starts)

            for start_page in starts_set:
                try:
                    doc[start_page - 1]
                except Exception:
                    continue
                key_pdf, tag_source = self._resolve_itr_tag(doc, pdf_path, start_page, preset)
                if key_pdf:
                    key_pdf = norm_key_value(key_pdf)
                    for suf in preset.get("match", {}).get("strip_suffixes", []):
                        suf_up = norm_key_value(suf)
                        if suf_up and key_pdf.endswith(suf_up):
                            key_pdf = key_pdf[: -len(suf_up)]

                status, excel_key, sheet, payload, short_key = match_one(
                    key_pdf,
                    excel_index,
                    preset,
                    self.match_memory,
                    chooser_func=self._thread_fuzzy_chooser,
                )

                if payload is None:
                    filled = compute_filled(preset, "", None, None, pdf_name)
                    item = ITRItem(
                        pdf_file=pdf_name,
                        set_start_page_1based=start_page,
                        key_pdf=key_pdf,
                        match_status=status,
                        excel_key=excel_key,
                        sheet_name="",
                        filled=filled,
                        tag_source=tag_source,
                    )
                else:
                    sheet_name, row_dict, col_map_norm = payload
                    filled = compute_filled(preset, sheet_name, row_dict, col_map_norm, pdf_name)
                    item = ITRItem(
                        pdf_file=pdf_name,
                        set_start_page_1based=start_page,
                        key_pdf=key_pdf,
                        match_status=status,
                        excel_key=excel_key,
                        sheet_name=sheet_name,
                        filled=filled,
                        tag_source=tag_source,
                    )
                items.append(item)

            doc.close()
            done_files = idx + 1
            self._preview_q.put(("progress", done_files, total_files, pdf_name))

        self._preview_q.put(("done", items, excel_index))

    def _thread_tag_candidate_chooser(self, pdf_name: str, candidates: List[str]) -> Tuple[str, bool]:
        event = threading.Event()
        box = {"val": "", "remember": False}
        self._preview_q.put(("tag_request", pdf_name, candidates, box, event))
        event.wait()
        return box["val"], box["remember"]

    def _thread_fuzzy_chooser(self, key_pdf: str, short_key: str, candidates: List[str]) -> str:
        event = threading.Event()
        box = {"val": ""}
        self._preview_q.put(("fuzzy_request", key_pdf, short_key, candidates, box, event))
        event.wait()
        return box["val"]

    def _poll_preview_queue(self):
        try:
            while True:
                msg = self._preview_q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    done, total, pdf_name = msg[1:]
                    self.export_progress["maximum"] = max(total, 1)
                    self.export_progress["value"] = done
                    self.status.config(text=f"正在处理 {pdf_name}（{done}/{total}）")
                elif kind == "fuzzy_request":
                    key_pdf, short_key, candidates, box, event = msg[1:]
                    pick = self.gui_fuzzy_chooser(key_pdf, short_key, candidates)
                    box["val"] = pick
                    event.set()
                elif kind == "tag_request":
                    pdf_name, candidates, box, event = msg[1:]
                    pick, remember = self.gui_tag_candidate_chooser(pdf_name, candidates)
                    box["val"] = pick
                    box["remember"] = remember
                    event.set()
                elif kind == "done":
                    items, excel_index = msg[1], msg[2]
                    self.excel_index = excel_index
                    self.items = items
                    self.current_idx = -1
                    self.listbox.delete(0, tk.END)
                    self.tree.delete(*self.tree.get_children())

                    for i, it in enumerate(self.items):
                        txt = (
                            f"{i + 1:03d} | {it.pdf_file} | start p{it.set_start_page_1based:>3} | "
                            f"key={it.key_pdf} ({it.tag_source}) | {it.match_status} | {it.excel_key}"
                        )
                        self.listbox.insert(tk.END, txt)

                    self.status.config(text=f"解析完成：{len(self.items)} 套 ITR | 记忆映射：{len(self.match_memory)}")
                    self.export_progress["value"] = 0
                    self.btn_parse.config(state="normal")
                    self.btn_export.config(state="normal")
                    self.btn_save_edits.config(state="normal")
                    self._update_main_preset_status()
                    return
                elif kind == "error":
                    err = msg[1]
                    self.status.config(text="解析失败")
                    self.btn_parse.config(state="normal")
                    self.btn_export.config(state="normal")
                    self.btn_save_edits.config(state="normal")
                    self._update_main_preset_status()
                    messagebox.showerror("错误", err)
                    return
        except queue.Empty:
            pass

        if self._preview_thread and self._preview_thread.is_alive():
            self.after(120, self._poll_preview_queue)
        else:
            self.btn_parse.config(state="normal")
            self.btn_export.config(state="normal")
            self.btn_save_edits.config(state="normal")
            self._update_main_preset_status()

    def on_select_item(self, _evt=None):
        self.save_current_edits()
        sel = self.listbox.curselection()
        if not sel:
            return
        self.current_idx = sel[0]
        self.render_current_item()

    def render_current_item(self):
        self.tree.delete(*self.tree.get_children())
        if not (0 <= self.current_idx < len(self.items)):
            return
        it = self.items[self.current_idx]
        active = self.global_cfg.get("active_preset", "")
        preset = load_preset(active) if active else None
        order = (
            [f.get("name", "") for f in preset.get("fields", []) if f.get("name")]
            if preset
            else list(it.filled.keys())
        )
        for k in order:
            self.tree.insert("", tk.END, values=(k, it.filled.get(k, "")))
        self.status.config(
            text=(
                f"当前：{it.pdf_file} | start p{it.set_start_page_1based} | "
                f"{it.key_pdf} ({it.tag_source}) | {it.match_status}"
            )
        )

    def _commit_active_editor_if_any(self):
        if self._active_entry is None or self._active_item_id is None:
            return
        try:
            new_val = self._active_entry.get()
            vals = list(self.tree.item(self._active_item_id, "values"))
            vals[1] = new_val
            self.tree.item(self._active_item_id, values=vals)
        finally:
            try:
                self._active_entry.destroy()
            except Exception:
                pass
            self._active_entry = None
            self._active_item_id = None

    def on_edit_cell(self, event):
        self._commit_active_editor_if_any()
        iid = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        if not iid or col != "#2":
            return
        x, y, w, h = self.tree.bbox(iid, col)
        old = self.tree.item(iid, "values")[1]
        ent = tk.Entry(self.tree)
        ent.place(x=x, y=y, width=w, height=h)
        ent.insert(0, old)
        ent.focus()
        self._active_entry = ent
        self._active_item_id = iid
        ent.bind("<Return>", lambda _e: self._commit_active_editor_if_any())
        ent.bind("<FocusOut>", lambda _e: self._commit_active_editor_if_any())
        ent.bind("<Escape>", lambda _e: self._commit_active_editor_if_any())

    def save_current_edits(self):
        self._commit_active_editor_if_any()
        if not (0 <= self.current_idx < len(self.items)):
            return
        it = self.items[self.current_idx]
        for iid in self.tree.get_children():
            k, v = self.tree.item(iid, "values")
            it.filled[str(k)] = str(v)

    # ------------------ 导出：后台线程，避免 GUI 未响应 ------------------
    def export_all_async(self):
        self.save_current_edits()
        if self._export_thread and self._export_thread.is_alive():
            messagebox.showinfo("提示", "正在导出，请稍等…")
            return

        if not self.preset_confirmed:
            messagebox.showwarning("提示", "请先确认预设")
            return

        active = self.global_cfg.get("active_preset", "")
        preset = load_preset(active) if active else None
        if not preset:
            messagebox.showwarning("提示", "未加载当前预设")
            return
        if not self.items or not self.pdf_paths:
            messagebox.showwarning("提示", "请先解析&预填")
            return

        # UI: disable
        self.btn_export.config(state="disabled")
        self.btn_parse.config(state="disabled")
        self.btn_save_edits.config(state="disabled")
        self.status.config(text="导出中…")
        self.export_progress["value"] = 0
        self.update_idletasks()

        # start thread
        batch = batch_id()
        self._export_thread = threading.Thread(target=self._export_worker, args=(preset, batch), daemon=True)
        self._export_thread.start()
        self.after(120, self._poll_queue)

    def _export_worker(self, preset: dict, batch: str):
        try:
            # 分组：按 PDF 文件名
            by_pdf: Dict[str, List[ITRItem]] = {}
            for it in self.items:
                by_pdf.setdefault(it.pdf_file, []).append(it)

            total_files = len(self.pdf_paths)
            out_pdfs = []
            filled_dir = ensure_output_batch_dir("filled", batch)
            report_dir = ensure_report_batch_dir(batch)

            for idx, pdf_path in enumerate(self.pdf_paths):
                pdf_name = os.path.basename(pdf_path)
                group = by_pdf.get(pdf_name, [])
                if not group:
                    done_files = idx + 1
                    self._q.put(("progress", done_files, total_files, pdf_name))
                    continue

                try:
                    doc = fitz.open(pdf_path)
                except Exception:
                    done_files = idx + 1
                    self._q.put(("progress", done_files, total_files, pdf_name))
                    continue
                cache = {}
                for it in group:
                    write_one_itr(doc, it.set_start_page_1based, preset, it.filled, cache)

                out_name = os.path.splitext(pdf_name)[0] + "_filled.pdf"
                out_path = os.path.join(filled_dir, out_name)
                doc.save(out_path)
                doc.close()
                out_pdfs.append(out_path)
                done_files = idx + 1
                self._q.put(("progress", done_files, total_files, pdf_name))

            report_path = self._save_report(preset, report_dir)
            self._q.put(("done", len(out_pdfs), report_path, filled_dir, report_dir))

        except Exception as e:
            self._q.put(("error", str(e)))

    def _poll_queue(self):
        try:
            while True:
                msg = self._q.get_nowait()
                kind = msg[0]
                if kind == "progress":
                    done, total, pdf_name = msg[1:]
                    self.export_progress["maximum"] = max(total, 1)
                    self.export_progress["value"] = done
                    self.status.config(text=f"正在处理 {pdf_name}（{done}/{total}）")
                elif kind == "done":
                    pdf_count, report_path, filled_dir, report_dir = msg[1], msg[2], msg[3], msg[4]
                    self.status.config(text="导出完成")
                    self.btn_export.config(state="normal")
                    self.btn_save_edits.config(state="normal")
                    self._update_main_preset_status()  # parse按钮状态恢复
                    messagebox.showinfo(
                        "完成",
                        f"输出PDF：{pdf_count}\n填充目录：{filled_dir}\n报告：{report_path}\n报告目录：{report_dir}",
                    )
                    return
                elif kind == "error":
                    err = msg[1]
                    self.status.config(text="导出失败")
                    self.btn_export.config(state="normal")
                    self.btn_save_edits.config(state="normal")
                    self._update_main_preset_status()
                    messagebox.showerror("错误", f"导出失败：{err}")
                    return
        except queue.Empty:
            pass

        if self._export_thread and self._export_thread.is_alive():
            self.after(120, self._poll_queue)
        else:
            # 线程结束但没收到done/error，兜底恢复
            self.btn_export.config(state="normal")
            self.btn_save_edits.config(state="normal")
            self._update_main_preset_status()

    def _save_report(self, preset: dict, report_dir: str) -> str:
        out_path = os.path.join(report_dir, "report.xlsx")
        fields_order = [f.get("name", "") for f in preset.get("fields", []) if f.get("name")]
        rows = []
        for it in self.items:
            r = {
                "PDF": it.pdf_file,
                "StartPage": it.set_start_page_1based,
                "PDF_Key": it.key_pdf,
                "MatchStatus": it.match_status,
                "ExcelKey": it.excel_key,
                "Sheet": it.sheet_name,
            }
            empties = []
            for k in fields_order:
                v = it.filled.get(k, "")
                r[k] = v
                if not str(v).strip():
                    empties.append(k)
            r["EmptyFields"] = ", ".join(empties)
            rows.append(r)
        pd.DataFrame(rows).to_excel(out_path, index=False)
        return out_path

    def open_memory_manager(self):
        win = tk.Toplevel(self)
        win.title(f"管理映射记忆（{os.path.basename(MATCH_MEMORY_PATH)}）")
        win.geometry("920x580")

        ttk.Label(win, text="短Key -> ExcelKey（用于模糊匹配记忆）").pack(anchor="w", padx=10, pady=10)
        tree = ttk.Treeview(win, columns=("short", "excel"), show="headings", height=18)
        tree.heading("short", text="短Key")
        tree.heading("excel", text="ExcelKey")
        tree.column("short", width=320, anchor="w")
        tree.column("excel", width=560, anchor="w")
        tree.pack(fill="both", expand=True, padx=10, pady=8)

        def refresh():
            tree.delete(*tree.get_children())
            for k, v in sorted(self.match_memory.items()):
                tree.insert("", tk.END, values=(k, v))

        refresh()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=10)

        def delete_selected():
            sel = tree.selection()
            if not sel:
                return
            for iid in sel:
                short_key = tree.item(iid, "values")[0]
                self.match_memory.pop(short_key, None)
            save_match_memory(self.match_memory)
            refresh()

        def clear_all():
            if not messagebox.askyesno("确认", "确定要清空所有记忆映射吗？"):
                return
            self.match_memory.clear()
            save_match_memory(self.match_memory)
            refresh()

        ttk.Button(btns, text="删除选中", command=delete_selected).pack(side="left")
        ttk.Button(btns, text="清空全部", command=clear_all).pack(side="left", padx=10)
        ttk.Button(btns, text="关闭", command=win.destroy).pack(side="right")
