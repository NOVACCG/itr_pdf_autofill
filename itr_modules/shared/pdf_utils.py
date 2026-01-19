"""Shared PDF helpers."""

from __future__ import annotations

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
