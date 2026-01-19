"""Shared PDF helpers."""

from __future__ import annotations

import hashlib
import re
from typing import List

import fitz


def _tokenize_for_wrap(text: str) -> List[str]:
    """Split text into wrap-friendly tokens (space first, then slash/dash)."""
    text = str(text).strip()
    if not text:
        return []
    parts = re.split(r"\s+", text)
    tokens: List[str] = []
    for part in parts:
        if not part:
            continue
        sub = re.split(r"([/-])", part)
        buffer = ""
        for item in sub:
            if item in ("/", "-"):
                buffer += item
                tokens.append(buffer)
                buffer = ""
            else:
                buffer = item if buffer == "" else buffer + item
        if buffer:
            tokens.append(buffer)
    return tokens


def _split_long_token(token: str, max_width: float, fontname: str, fontsize: float) -> List[str]:
    if not token:
        return []
    if fitz.get_text_length(token, fontname=fontname, fontsize=fontsize) <= max_width:
        return [token]
    parts: List[str] = []
    buf = ""
    for ch in token:
        trial = buf + ch
        if fitz.get_text_length(trial, fontname=fontname, fontsize=fontsize) <= max_width:
            buf = trial
        else:
            if buf:
                parts.append(buf)
            buf = ch
    if buf:
        parts.append(buf)
    return parts


def _wrap_tokens(tokens: List[str], max_width: float, fontname: str, fontsize: float) -> List[str]:
    lines: List[str] = []
    current = ""
    for token in tokens:
        for piece in _split_long_token(token, max_width, fontname, fontsize):
            trial = piece if current == "" else current + " " + piece
            width = fitz.get_text_length(trial, fontname=fontname, fontsize=fontsize)
            if width <= max_width:
                current = trial
            else:
                if current:
                    lines.append(current)
                current = piece
    if current:
        lines.append(current)
    return lines


def fit_text_to_box(page: fitz.Page, rect: fitz.Rect, text: str, text_cfg: dict) -> None:
    """Fit text into rect with wrapping + font scaling.

    Layout logic:
    - text is rendered from the top-left inner box (rect minus padding).
    - line height = font_size * line_gap.
    - choose the largest font_size that allows all lines to fit the inner height.
    - fallback to min_font_size with textbox rendering (clipped to inner box).
    """
    if text is None:
        return
    text = str(text).strip()
    if text == "":
        return

    max_font = float(text_cfg.get("max_font_size", 9))
    min_font = float(text_cfg.get("min_font_size", 5))
    padding = float(text_cfg.get("padding", 2))
    line_gap = float(text_cfg.get("line_gap", 1.15))

    inner = fitz.Rect(rect.x0 + padding, rect.y0 + 1, rect.x1 - padding, rect.y1 - 1)
    max_width = inner.width
    max_height = inner.height

    fontname = "helv"
    tokens = _tokenize_for_wrap(text)

    font_size = max_font
    target_min = max(min_font, 1.0)
    while font_size >= target_min:
        lines = _wrap_tokens(tokens, max_width, fontname, font_size)
        needed_height = len(lines) * font_size * line_gap
        if needed_height <= max_height:
            y = inner.y0 + font_size
            for line in lines:
                page.insert_text(
                    (inner.x0, y),
                    line,
                    fontsize=font_size,
                    fontname=fontname,
                    color=(0, 0, 0),
                )
                y += font_size * line_gap
            return
        font_size -= 0.5

    font_size = target_min
    while font_size > 0.5:
        lines = _wrap_tokens(tokens, max_width, fontname, font_size)
        needed_height = len(lines) * font_size * line_gap
        if needed_height <= max_height:
            y = inner.y0 + font_size
            for line in lines:
                page.insert_text(
                    (inner.x0, y),
                    line,
                    fontsize=font_size,
                    fontname=fontname,
                    color=(0, 0, 0),
                )
                y += font_size * line_gap
            return
        font_size -= 0.5

    lines = _wrap_tokens(tokens, max_width, fontname, 0.5)
    page.insert_textbox(
        inner,
        "\n".join(lines),
        fontsize=0.5,
        fontname=fontname,
        color=(0, 0, 0),
        align=0,
    )


def norm_text(s: str) -> str:
    """Normalize text: uppercase and drop non-alphanumerics."""
    return re.sub(r"[^A-Z0-9]+", "", (s or "").upper())


def normalize_cell_text(s: str) -> str:
    """Trim and collapse whitespace for cell content."""
    return re.sub(r"\s+", " ", str(s or "")).strip()


def is_valid_tag_value(s: str) -> bool:
    if s is None:
        return False
    val = str(s).strip()
    if not val:
        return False
    if len(val) < 4 or len(val) > 50:
        return False
    if val.upper() in {"OK", "NA", "PL"}:
        return False
    return bool(re.fullmatch(r"[A-Za-z0-9\-\._/]+", val))


def extract_tag_candidates_from_text(text: str, regex_pattern: str) -> list[dict]:
    rx = re.compile(regex_pattern, re.IGNORECASE)
    candidates: list[dict] = []
    seen = set()
    for idx, match in enumerate(rx.finditer(text or "")):
        value = match.group(1) if match.groups() else match.group(0)
        value = normalize_cell_text(value)
        if not value or value in seen:
            continue
        seen.add(value)
        line_hint = None
        if text:
            cursor = 0
            for line in text.splitlines():
                next_cursor = cursor + len(line) + 1
                if cursor <= match.start() <= next_cursor:
                    line_hint = line.strip()
                    break
                cursor = next_cursor
        candidates.append({
            "value": value,
            "page_index": 0,
            "span_index": idx,
            "line_hint": line_hint,
        })
    return candidates


def extract_tag_candidates_first_page(doc: fitz.Document, regex_pattern: str) -> list[dict]:
    page = doc[0]
    text = page.get_text("text") or ""
    return extract_tag_candidates_from_text(text, regex_pattern)


def parse_pages_per_itr_regex(doc: fitz.Document, pattern: str, scan_pages: int) -> int | None:
    """Scan initial pages for 'Page x of y' to infer pages per ITR."""
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


def extract_rulings(page: fitz.Page, tol: float = 1.5) -> tuple[list[tuple[float, float, float]], list[tuple[float, float, float]]]:
    """Extract vertical/horizontal ruling lines from drawings."""
    verticals: list[tuple[float, float, float]] = []
    horizontals: list[tuple[float, float, float]] = []

    drawings = page.get_drawings()
    for d in drawings:
        for it in d.get("items", []):
            if not it:
                continue
            kind = it[0]
            if kind == "l":
                (x0, y0) = it[1]
                (x1, y1) = it[2]
                if abs(x0 - x1) <= tol:
                    x = (x0 + x1) / 2.0
                    verticals.append((x, min(y0, y1), max(y0, y1)))
                elif abs(y0 - y1) <= tol:
                    y = (y0 + y1) / 2.0
                    horizontals.append((y, min(x0, x1), max(x0, x1)))
            elif kind == "re":
                r = it[1]
                if isinstance(r, fitz.Rect):
                    x0, y0, x1, y1 = r.x0, r.y0, r.x1, r.y1
                    verticals.extend([(x0, y0, y1), (x1, y0, y1)])
                    horizontals.extend([(y0, x0, x1), (y1, x0, x1)])

    verticals = [(x, y0, y1) for (x, y0, y1) in verticals if (y1 - y0) > 6]
    horizontals = [(y, x0, x1) for (y, x0, x1) in horizontals if (x1 - x0) > 6]
    return verticals, horizontals


def cell_rect_for_word(word_rect: fitz.Rect, verticals, horizontals) -> fitz.Rect | None:
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


def _cell_rect_from_point(
    mid_x: float,
    mid_y: float,
    verticals: list[tuple[float, float, float]],
    horizontals: list[tuple[float, float, float]],
    tol: float = 1.5,
) -> fitz.Rect | None:
    xs = [x for x, y0, y1 in verticals if y0 - tol <= mid_y <= y1 + tol]
    ys = [y for y, x0, x1 in horizontals if x0 - tol <= mid_x <= x1 + tol]
    if not xs or not ys:
        return None
    xs_sorted = sorted(xs)
    ys_sorted = sorted(ys)
    xs_deduped = []
    for x in xs_sorted:
        if not xs_deduped or abs(x - xs_deduped[-1]) > tol:
            xs_deduped.append(x)
    ys_deduped = []
    for y in ys_sorted:
        if not ys_deduped or abs(y - ys_deduped[-1]) > tol:
            ys_deduped.append(y)
    lefts = [x for x in xs_deduped if x <= mid_x + tol]
    rights = [x for x in xs_deduped if x >= mid_x - tol]
    tops = [y for y in ys_deduped if y <= mid_y + tol]
    bottoms = [y for y in ys_deduped if y >= mid_y - tol]
    if not lefts or not rights or not tops or not bottoms:
        return None
    x0 = max(lefts)
    x1 = min(rights)
    y0 = max(tops)
    y1 = min(bottoms)
    if x1 <= x0 or y1 <= y0:
        return None
    return fitz.Rect(x0 + 0.2, y0 + 0.2, x1 - 0.2, y1 - 0.2)


def _anchor_candidates(words: list[tuple], anchor_norms: list[str]) -> list[tuple[fitz.Rect, str]]:
    candidates: list[tuple[fitz.Rect, str]] = []
    if not words:
        return candidates
    by_line: dict[int, list[tuple]] = {}
    for w in words:
        by_line.setdefault(w[6], []).append(w)
    anchor_set = {norm_text(a) for a in anchor_norms if norm_text(a)}
    for line_words in by_line.values():
        line_words.sort(key=lambda w: (w[0], w[1]))
        norms = [norm_text(w[4]) for w in line_words]
        for i in range(len(line_words)):
            combined = ""
            for j in range(i, min(i + 4, len(line_words))):
                combined += norms[j]
                if combined in anchor_set:
                    x0 = min(line_words[k][0] for k in range(i, j + 1))
                    y0 = min(line_words[k][1] for k in range(i, j + 1))
                    x1 = max(line_words[k][2] for k in range(i, j + 1))
                    y1 = max(line_words[k][3] for k in range(i, j + 1))
                    candidates.append((fitz.Rect(x0, y0, x1, y1), combined))
    return candidates


def find_anchor_cell_strict(
    page: fitz.Page,
    anchor_norms: list[str],
    tol: float = 1.5,
) -> tuple[fitz.Rect | None, dict]:
    verticals, horizontals = extract_rulings(page, tol=tol)
    words = page.get_text("words") or []
    candidates = _anchor_candidates(words, anchor_norms)
    debug = {"anchor_norms": anchor_norms}
    for bbox, _ in candidates:
        mid_x = (bbox.x0 + bbox.x1) / 2.0
        mid_y = (bbox.y0 + bbox.y1) / 2.0
        cell = _cell_rect_from_point(mid_x, mid_y, verticals, horizontals, tol=tol)
        if not cell:
            continue
        raw = get_cell_text(page, cell)
        cell_norm = norm_text(raw)
        if cell_norm in {norm_text(a) for a in anchor_norms}:
            debug.update({
                "key_cell_rect": cell,
                "key_cell_text_raw": raw,
                "key_cell_text_norm": cell_norm,
            })
            return cell, debug
    debug["error"] = "anchor_cell_not_found"
    return None, debug


def build_covering_verticals(
    page: fitz.Page,
    key_cell: fitz.Rect,
    eps: float = 0.25,
    cover: float = 0.8,
) -> list[float]:
    verticals, _ = extract_rulings(page)
    kx0, ky0, kx1, ky1 = key_cell
    height = max(ky1 - ky0, 1.0)
    xs: list[float] = []
    for x, y0, y1 in verticals:
        overlap_h = min(y1, ky1) - max(y0, ky0)
        if overlap_h >= cover * height:
            xs.append(x)
    xs.sort()
    deduped: list[float] = []
    for x in xs:
        if not deduped or abs(x - deduped[-1]) > eps:
            deduped.append(x)
    return deduped


def build_covering_horizontals(
    page: fitz.Page,
    key_cell: fitz.Rect,
    eps: float = 0.25,
    cover: float = 0.8,
) -> list[float]:
    _, horizontals = extract_rulings(page)
    kx0, ky0, kx1, ky1 = key_cell
    width = max(kx1 - kx0, 1.0)
    ys: list[float] = []
    for y, x0, x1 in horizontals:
        overlap_w = min(x1, kx1) - max(x0, kx0)
        if overlap_w >= cover * width:
            ys.append(y)
    ys.sort()
    deduped: list[float] = []
    for y in ys:
        if not deduped or abs(y - deduped[-1]) > eps:
            deduped.append(y)
    return deduped


def find_adjacent_cell_by_lines(
    page: fitz.Page,
    key_cell: fitz.Rect,
    direction: str,
    eps: float = 0.25,
    cover: float = 0.8,
    pad: float = 0.4,
) -> tuple[fitz.Rect | None, dict]:
    kx0, ky0, kx1, ky1 = key_cell
    direction = (direction or "RIGHT").upper()
    vcover = build_covering_verticals(page, key_cell, eps=eps, cover=cover)
    hcover = build_covering_horizontals(page, key_cell, eps=eps, cover=cover)
    debug: dict = {
        "direction": direction,
        "vcover_count": len(vcover),
        "hcover_count": len(hcover),
    }
    chosen_line = None
    if direction == "RIGHT":
        chosen_line = min([x for x in vcover if x > kx1 + eps], default=None)
        if chosen_line is None:
            debug["error"] = "adjacent_cell_not_found"
            debug["chosen_line"] = None
            return None, debug
        rect = fitz.Rect(kx1, ky0, chosen_line, ky1)
    elif direction == "LEFT":
        chosen_line = max([x for x in vcover if x < kx0 - eps], default=None)
        if chosen_line is None:
            debug["error"] = "adjacent_cell_not_found"
            debug["chosen_line"] = None
            return None, debug
        rect = fitz.Rect(chosen_line, ky0, kx0, ky1)
    elif direction == "DOWN":
        chosen_line = min([y for y in hcover if y > ky1 + eps], default=None)
        if chosen_line is None:
            debug["error"] = "adjacent_cell_not_found"
            debug["chosen_line"] = None
            return None, debug
        rect = fitz.Rect(kx0, ky1, kx1, chosen_line)
    else:
        chosen_line = max([y for y in hcover if y < ky0 - eps], default=None)
        if chosen_line is None:
            debug["error"] = "adjacent_cell_not_found"
            debug["chosen_line"] = None
            return None, debug
        rect = fitz.Rect(kx0, chosen_line, kx1, ky0)
    debug["chosen_line"] = chosen_line
    pad = min(max(pad, 0.0), 0.8)
    rect = fitz.Rect(rect.x0 + pad, rect.y0 + pad, rect.x1 - pad, rect.y1 - pad)
    return rect, debug


def get_cell_text_cached(
    cell: fitz.Rect,
    words: list[tuple] | None,
    buckets: dict[int, list[tuple]] | None,
    cache: dict[tuple[float, float, float, float], str] | None,
    bucket_size: float = 8.0,
) -> str:
    if cache is None:
        cache = {}
    key = (round(cell.x0, 2), round(cell.y0, 2), round(cell.x1, 2), round(cell.y1, 2))
    if key in cache:
        return cache[key]

    picked = []
    if words:
        if buckets:
            if bucket_size <= 0:
                bucket_size = 8.0
            b0 = int(cell.y0 // bucket_size)
            b1 = int(cell.y1 // bucket_size)
            for b in range(b0, b1 + 1):
                for w in buckets.get(b, []):
                    wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
                    cx = (wx0 + wx1) / 2.0
                    cy = (wy0 + wy1) / 2.0
                    if cell.x0 <= cx <= cell.x1 and cell.y0 <= cy <= cell.y1:
                        picked.append(w)
        else:
            for w in words:
                wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
                cx = (wx0 + wx1) / 2.0
                cy = (wy0 + wy1) / 2.0
                if cell.x0 <= cx <= cell.x1 and cell.y0 <= cy <= cell.y1:
                    picked.append(w)

    text = _norm_join_words(picked).strip()
    cache[key] = text
    return text


def extract_tag_by_cell_adjacency(
    page: fitz.Page,
    matchkey_norm: str,
    direction: str,
    value_regex: str,
) -> tuple[str | None, dict]:
    key_cell, key_debug = find_anchor_cell_strict(page, [matchkey_norm])
    debug = {"direction": direction}
    debug.update(key_debug)
    if not key_cell:
        debug["error"] = "anchor_cell_not_found"
        return None, debug

    debug.update({
        "key_cell_rect": key_cell,
        "key_cell_text_norm": key_debug.get("key_cell_text_norm", ""),
    })
    value_cell, adj_debug = find_adjacent_cell_by_lines(page, key_cell, direction, eps=0.25, cover=0.8, pad=0.4)
    debug.update(adj_debug)
    if not value_cell:
        debug["error"] = "adjacent_cell_not_found"
        return None, debug
    debug["value_cell_rect"] = value_cell
    if value_cell.width <= 2 or value_cell.height <= 2:
        debug["error"] = "adjacent_cell_not_found"
        return None, debug

    raw = get_cell_text(page, value_cell)
    debug["value_cell_text_raw"] = raw
    debug["value_cell_text_raw_preview"] = (raw or "")[:200]
    normalized_raw = normalize_cell_text(raw)
    debug["value_cell_text_normed"] = normalized_raw
    if not normalized_raw:
        debug["error"] = "adjacent_cell_empty"
        return None, debug
    try:
        extracted = extract_value_by_regex(normalized_raw, value_regex)
    except re.error:
        debug["error"] = "regex_no_match"
        return None, debug
    extracted = normalize_cell_text(extracted)
    if not extracted:
        debug["error"] = "regex_no_match"
        return None, debug
    return extracted, debug


def find_cell_by_candidates(
    page: fitz.Page,
    candidates: list[str],
    verticals,
    horizontals,
    search_clip: fitz.Rect | None = None,
) -> tuple[fitz.Rect | None, dict]:
    cand_norms = [norm_text(c) for c in candidates if norm_text(c)]
    debug = {"candidates": cand_norms}
    if not cand_norms:
        debug["error"] = "candidates_empty"
        return None, debug

    words = page.get_text("words", clip=search_clip) if search_clip else page.get_text("words")
    hits = []
    for x0, y0, x1, y1, w, *_ in words:
        wn = norm_text(w)
        if not wn:
            continue
        if not any((wn == cand or cand in wn or wn in cand) for cand in cand_norms):
            continue
        cell = cell_rect_for_word(fitz.Rect(x0, y0, x1, y1), verticals, horizontals)
        if not cell:
            continue
        cell_norm = norm_text(get_cell_text(page, cell))
        matched = None
        for cand in cand_norms:
            if cell_norm == cand or cell_norm.endswith(cand):
                matched = cand
                break
        if matched:
            hits.append((cell, matched, cell_norm))

    if not hits:
        debug["error"] = "match_cell_not_found"
        return None, debug

    hits.sort(key=lambda item: (item[0].y0, item[0].x0))
    cell, matched, cell_norm = hits[0]
    debug.update({"matched_candidate": matched, "cell_norm": cell_norm, "cell_rect": cell})
    return cell, debug


def extract_tag_by_cell_adjacency_candidates(
    page: fitz.Page,
    candidates: list[str],
    direction: str,
    value_regex: str,
) -> tuple[str | None, dict]:
    key_cell, key_debug = find_anchor_cell_strict(page, candidates)
    debug = {"direction": direction}
    debug.update(key_debug)
    if not key_cell:
        debug["error"] = "anchor_cell_not_found"
        return None, debug

    debug.update({
        "key_cell_rect": key_cell,
        "key_cell_text_norm": key_debug.get("key_cell_text_norm", ""),
    })
    value_cell, adj_debug = find_adjacent_cell_by_lines(page, key_cell, direction, eps=0.25, cover=0.8, pad=0.4)
    debug.update(adj_debug)
    if not value_cell:
        debug["error"] = "adjacent_cell_not_found"
        return None, debug
    debug["value_cell_rect"] = value_cell
    if value_cell.width <= 2 or value_cell.height <= 2:
        debug["error"] = "adjacent_cell_not_found"
        return None, debug

    raw = get_cell_text(page, value_cell)
    debug["value_cell_text_raw"] = raw
    debug["value_cell_text_raw_preview"] = (raw or "")[:200]
    normalized_raw = normalize_cell_text(raw)
    debug["value_cell_text_normed"] = normalized_raw
    if not normalized_raw:
        debug["error"] = "adjacent_cell_empty"
        return None, debug
    try:
        extracted = extract_value_by_regex(normalized_raw, value_regex)
    except re.error:
        debug["error"] = "regex_no_match"
        return None, debug
    extracted = normalize_cell_text(extracted)
    if not extracted:
        debug["error"] = "regex_no_match"
        return None, debug
    return extracted, debug


def find_adjacent_cell_with_tolerance(
    page: fitz.Page,
    key_rect: fitz.Rect,
    direction: str,
    tol: float = 4.0,
    overlap_ratio: float = 0.6,
) -> tuple[fitz.Rect | None, dict]:
    words = page.get_text("words") or []
    debug = {"direction": direction, "tol": tol, "overlap_ratio": overlap_ratio}
    if not words:
        debug["error"] = "no_words"
        return None, debug

    direction = (direction or "RIGHT").upper()
    key_h = max(key_rect.height, 1.0)
    key_w = max(key_rect.width, 1.0)
    hits = []
    for x0, y0, x1, y1, w, *_ in words:
        rect = fitz.Rect(x0, y0, x1, y1)
        if direction in {"RIGHT", "LEFT"}:
            overlap = max(0.0, min(key_rect.y1, rect.y1) - max(key_rect.y0, rect.y0))
            if overlap < overlap_ratio * key_h:
                continue
            if direction == "RIGHT":
                if abs(rect.x0 - key_rect.x1) <= tol:
                    hits.append(rect)
            else:
                if abs(rect.x1 - key_rect.x0) <= tol:
                    hits.append(rect)
        else:
            overlap = max(0.0, min(key_rect.x1, rect.x1) - max(key_rect.x0, rect.x0))
            if overlap < overlap_ratio * key_w:
                continue
            if direction == "DOWN":
                if abs(rect.y0 - key_rect.y1) <= tol:
                    hits.append(rect)
            else:
                if abs(rect.y1 - key_rect.y0) <= tol:
                    hits.append(rect)

    if hits:
        hits.sort(key=lambda r: (r.y0, r.x0))
        debug["match"] = "tolerance"
        return hits[0], debug

    fallback = []
    for x0, y0, x1, y1, w, *_ in words:
        rect = fitz.Rect(x0, y0, x1, y1)
        overlap = max(0.0, min(key_rect.y1, rect.y1) - max(key_rect.y0, rect.y0))
        if overlap < overlap_ratio * key_h:
            continue
        if rect.x0 > key_rect.x1:
            fallback.append(rect)
    if fallback:
        fallback.sort(key=lambda r: r.x0)
        debug["match"] = "fallback_right_band"
        return fallback[0], debug

    debug["error"] = "adjacent_cell_not_found"
    return None, debug


def extract_candidates_in_cell_text(text: str, regex_pattern: str) -> list[str]:
    rx = re.compile(regex_pattern, re.IGNORECASE)
    candidates: list[str] = []
    seen = set()
    for match in rx.finditer(text or ""):
        value = match.group(1) if match.groups() else match.group(0)
        value = normalize_cell_text(value)
        if not value or value in seen:
            continue
        seen.add(value)
        candidates.append(value)
    return candidates


def extract_value_by_regex(value_raw: str, pattern: str) -> str:
    if not pattern:
        return value_raw
    rx = re.compile(pattern, re.IGNORECASE)
    match = rx.search(value_raw or "")
    if not match:
        return ""
    if match.groups():
        for group in match.groups():
            if group:
                return group
        return ""
    return match.group(0)


def template_fingerprint(preset_name: str, key_norm: str, direction: str, value_regex: str) -> str:
    base = f"{preset_name}::{key_norm}::{direction}::{value_regex}"
    return hashlib.sha1(base.encode("utf-8")).hexdigest()


def _norm_join_words(words_in_row) -> str:
    """Join words in a row (PyMuPDF words tuples) from left to right."""
    if not words_in_row:
        return ""
    words_in_row = sorted(words_in_row, key=lambda w: (w[0], w[1]))
    return " ".join((w[4] or "").strip() for w in words_in_row if (w[4] or "").strip())


def _cell_text_from_row_words(row_words, x0: float, x1: float) -> str:
    """Get cell text from row words whose centers fall in [x0, x1]."""
    if not row_words:
        return ""
    picked = []
    for w in row_words:
        wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
        cx = (wx0 + wx1) / 2.0
        if x0 <= cx <= x1:
            picked.append(w)
    return _norm_join_words(picked)


def _uniq_sorted(vals, tol: float = 0.8) -> list[float]:
    """Sort + dedupe: values within tol are treated as one."""
    vals = sorted(vals)
    out = []
    for v in vals:
        if not out or abs(v - out[-1]) > tol:
            out.append(v)
    return out


def build_table_row_lines(
    page: fitz.Page,
    horizontals,
    x_left: float,
    x_right: float,
    y_start: float,
    min_span_pad: float = 8.0,
) -> list[float]:
    """Filter horizontal rulings that span the table width and return deduped y's."""
    ys = []
    for y, x0, x1 in horizontals:
        if x0 <= x_left + min_span_pad and x1 >= x_right - min_span_pad and y >= y_start - 2:
            ys.append(y)
    return _uniq_sorted(ys)


def is_pure_int(s: str) -> bool:
    s = (s or "").strip()
    return bool(s) and s.isdigit()


def rect_between_lines(x0: float, x1: float, y0: float, y1: float, pad: float = 0.6) -> fitz.Rect:
    return fitz.Rect(x0 + pad, y0 + pad, x1 - pad, y1 - pad)


def _unique_sorted_x_from_verticals(verticals) -> list[float]:
    """Extract sorted unique x values from vertical segments."""
    xs: list[float] = []
    for v in verticals or []:
        if not v:
            continue
        if len(v) == 3:
            x, _, _ = v
            xs.append(float(x))
        elif len(v) >= 4:
            x0, _, _, _ = v[:4]
            xs.append(float(x0))
    xs = sorted({round(x, 2) for x in xs})
    return xs


def _snap_col_bounds(xs: list[float], x_center: float) -> tuple[float, float] | None:
    """Return adjacent x bounds that contain x_center or nearest segment."""
    if not xs or len(xs) < 2:
        return None
    for i in range(len(xs) - 1):
        if xs[i] - 1.0 <= x_center <= xs[i + 1] + 1.0:
            return (xs[i], xs[i + 1])
    best = None
    best_d = 1e18
    for i in range(len(xs) - 1):
        c = (xs[i] + xs[i + 1]) / 2.0
        d = abs(c - x_center)
        if d < best_d:
            best_d = d
            best = (xs[i], xs[i + 1])
    return best


def find_cell_by_exact_norm(
    page: fitz.Page,
    target: str,
    verticals,
    horizontals,
    search_clip: fitz.Rect | None = None,
) -> fitz.Rect | None:
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


def find_lowest_header_anchor(page: fitz.Page, candidates: list[str], verticals, horizontals) -> fitz.Rect | None:
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


def header_row_band(no_cell: fitz.Rect, pad: float = 3.0) -> fitz.Rect:
    return fitz.Rect(0, no_cell.y0 - pad, 10000, no_cell.y1 + pad)


def collect_ex_header_cells(page: fitz.Page, row_band: fitz.Rect, verticals, horizontals) -> list[tuple[str, fitz.Rect]]:
    """Collect EX* header cells by scanning header band columns."""
    xs = _unique_sorted_x_from_verticals(verticals)
    if len(xs) < 2:
        return []
    cells: list[tuple[str, fitz.Rect]] = []
    for x0, x1 in zip(xs, xs[1:]):
        rr = fitz.Rect(x0, row_band.y0, x1, row_band.y1)
        txt = norm_text(get_cell_text(page, rr))
        if re.fullmatch(r"EX[A-Z0-9]{1,3}", txt or ""):
            cells.append((txt, rr))
    return cells


def find_ok_na_pl_cells(page: fitz.Page, row_band: fitz.Rect, verticals, horizontals) -> dict[str, fitz.Rect]:
    xs = _unique_sorted_x_from_verticals(verticals)
    if len(xs) < 2:
        return {}
    cells = {}
    for x0, x1 in zip(xs, xs[1:]):
        rr = fitz.Rect(x0, row_band.y0, x1, row_band.y1)
        txt = norm_text(get_cell_text(page, rr))
        if txt in ("OK", "NA", "PL"):
            cells[txt] = rr
    return cells


def find_ex_concept_cells(page: fitz.Page, verticals, horizontals) -> tuple[fitz.Rect | None, fitz.Rect | None]:
    ex_label = find_cell_by_exact_norm(page, "Ex Concept", verticals, horizontals)
    if not ex_label:
        return None, None
    xs = _unique_sorted_x_from_verticals(verticals)
    if len(xs) < 2:
        return ex_label, None
    x_center = (ex_label.x0 + ex_label.x1) / 2.0
    bounds = _snap_col_bounds(xs, x_center)
    if not bounds:
        return ex_label, None
    x0, x1 = bounds
    ex_val = fitz.Rect(x0, ex_label.y1 + 1.0, x1, ex_label.y1 + (ex_label.height + 2.0))
    return ex_label, ex_val


def draw_checkmark(page: fitz.Page, rr: fitz.Rect, width: float = 1.6) -> None:
    """Draw a checkmark using lines to avoid font issues."""
    if rr is None:
        return

    side = max(min(rr.width, rr.height), 1.0)
    cx = (rr.x0 + rr.x1) / 2.0
    cy = (rr.y0 + rr.y1) / 2.0
    sq = fitz.Rect(cx - side / 2.0, cy - side / 2.0, cx + side / 2.0, cy + side / 2.0)

    inset = max(side * 0.18, 1.0)
    r = fitz.Rect(sq.x0 + inset, sq.y0 + inset, sq.x1 - inset, sq.y1 - inset)

    w = r.width
    h = r.height

    p1 = (r.x0 + 0.18 * w, r.y0 + 0.55 * h)
    p2 = (r.x0 + 0.42 * w, r.y0 + 0.78 * h)
    p3 = (r.x0 + 0.82 * w, r.y0 + 0.22 * h)

    page.draw_line(p1, p2, width=width)
    page.draw_line(p2, p3, width=width)


def extract_table_grid_lines(
    page: fitz.Page,
    table_bbox: fitz.Rect,
    verticals: list[tuple[float, float, float]],
    horizontals: list[tuple[float, float, float]],
    pad: float = 0.5,
) -> tuple[list[float], list[float]]:
    """Extract table grid x/y coordinates within the table bbox."""
    xs = []
    ys = []
    for x, y0, y1 in verticals:
        if y1 >= table_bbox.y0 - pad and y0 <= table_bbox.y1 + pad:
            if table_bbox.x0 - pad <= x <= table_bbox.x1 + pad:
                xs.append(x)
    for y, x0, x1 in horizontals:
        if x1 >= table_bbox.x0 - pad and x0 <= table_bbox.x1 + pad:
            if table_bbox.y0 - pad <= y <= table_bbox.y1 + pad:
                ys.append(y)

    xs.extend([table_bbox.x0, table_bbox.x1])
    ys.extend([table_bbox.y0, table_bbox.y1])

    xs = _uniq_sorted(xs, tol=0.6)
    ys = _uniq_sorted(ys, tol=0.6)
    return xs, ys


def snap_to_grid_x(cx: float, xs: list[float]) -> tuple[float, float] | None:
    """Find nearest left/right grid lines that bound cx."""
    if not xs or len(xs) < 2:
        return None
    xs = sorted(xs)
    left = None
    right = None
    for x in xs:
        if x <= cx:
            left = x
        if x >= cx and right is None:
            right = x
    if left is None or right is None:
        return None
    if left == right:
        for i in range(len(xs) - 1):
            if xs[i] <= cx <= xs[i + 1]:
                return xs[i], xs[i + 1]
        return None
    return left, right


def row_band_from_ys(index: int, ys: list[float]) -> tuple[float, float] | None:
    if index < 0 or index + 1 >= len(ys):
        return None
    return ys[index], ys[index + 1]


def row_index_from_ys(ys: list[float], y_center: float) -> int:
    for i in range(len(ys) - 1):
        if ys[i] - 1 <= y_center <= ys[i + 1] + 1:
            return i
    return -1


def split_columns_from_grid(xs: list[float], x0: float, x1: float, count: int) -> list[tuple[float, float]]:
    """Split a span into count columns, snapping to grid lines."""
    if count <= 0:
        return []
    xs_sorted = sorted([x for x in xs if x0 - 1 <= x <= x1 + 1])
    if len(xs_sorted) < 2:
        return []
    if count == 1:
        return [(xs_sorted[0], xs_sorted[-1])]
    total_segments = len(xs_sorted) - 1
    step = max(1, total_segments // count)
    bounds = []
    start_idx = 0
    for i in range(count):
        end_idx = start_idx + step
        if i == count - 1 or end_idx >= total_segments:
            end_idx = total_segments
        bounds.append((xs_sorted[start_idx], xs_sorted[end_idx]))
        start_idx = end_idx
    return bounds


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


def _scan_header_cells_by_grid(
    page: fitz.Page,
    xs: list[float],
    band: tuple[float, float] | None,
    header_norms: list[str],
) -> dict[str, fitz.Rect]:
    if not band or not xs or len(xs) < 2:
        return {}
    y0, y1 = band
    header_norms_set = {norm_text(h) for h in header_norms if norm_text(h)}
    cells: dict[str, fitz.Rect] = {}
    for x0, x1 in zip(xs, xs[1:]):
        rect = fitz.Rect(x0, y0, x1, y1)
        txt = norm_text(get_cell_text(page, rect))
        if txt and txt in header_norms_set:
            cells[txt] = rect
    return cells


def _detect_header_row_index(
    page: fitz.Page,
    ys: list[float],
    header_norms: list[str],
    header_anchor: fitz.Rect | None,
) -> int:
    if header_anchor:
        idx = _header_row_index(ys, header_anchor)
        if idx >= 0:
            return idx
    if not ys:
        return -1
    header_set = {norm_text(h) for h in header_norms if norm_text(h)}
    words = page.get_text("words") or []
    counts: dict[int, int] = {}
    for x0, y0, x1, y1, w, *_ in words:
        if norm_text(w) in header_set:
            cy = (y0 + y1) / 2.0
            row_idx = row_index_from_ys(ys, cy)
            if row_idx >= 0:
                counts[row_idx] = counts.get(row_idx, 0) + 1
    if not counts:
        return -1
    return max(counts.items(), key=lambda kv: kv[1])[0]


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


def detect_checkitems_table(
    page: fitz.Page,
    header_norms: list[str],
    index_norm: str,
    state_norms: list[str],
) -> dict:
    """Detect CheckItems table structure for a single page."""
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

    header_row_idx = _detect_header_row_index(page, ys, header_norms, header_anchor)
    header_band_ys = row_band_from_ys(header_row_idx, ys) if header_row_idx >= 0 else None
    if not header_cells and header_band_ys:
        header_cells.update(_scan_header_cells_by_grid(page, xs, header_band_ys, header_norms))

    index_header = header_cells.get(index_norm)
    index_bounds = _find_column_bounds(xs, index_header)
    if index_bounds is None and header_anchor:
        index_bounds = _find_column_bounds(xs, header_anchor)

    expected = 1
    numbered_rows: list[int] = []
    if index_bounds and ys:
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

    state_bounds_map: dict[str, tuple[float, float]] = {}
    if header_band_ys and xs:
        header_cells.update(_scan_header_cells_by_grid(page, xs, header_band_ys, header_norms))

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

    header_texts = {}
    for norm, rect in header_cells.items():
        header_texts[norm] = get_cell_text(page, rect)

    return {
        "table_bbox": table_rect,
        "grid_xs": xs,
        "grid_ys": ys,
        "header_cells": header_cells,
        "header_texts": header_texts,
        "index_bounds": index_bounds,
        "numbered_rows": numbered_rows,
        "state_bounds": state_bounds_map,
    }


__all__ = [
    "detect_checkitems_table",
    "draw_checkmark",
    "extract_tag_by_cell_adjacency",
    "extract_tag_by_cell_adjacency_candidates",
    "extract_tag_candidates_from_text",
    "extract_tag_candidates_first_page",
    "extract_candidates_in_cell_text",
    "extract_value_by_regex",
    "fit_text_to_box",
    "build_covering_horizontals",
    "build_covering_verticals",
    "find_adjacent_cell_by_lines",
    "find_anchor_cell_strict",
    "find_cell_by_candidates",
    "find_adjacent_cell_with_tolerance",
    "get_cell_text_cached",
    "is_valid_tag_value",
    "normalize_cell_text",
    "norm_text",
    "row_band_from_ys",
    "template_fingerprint",
]
