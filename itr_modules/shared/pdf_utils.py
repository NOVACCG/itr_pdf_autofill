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


def _wrap_tokens(tokens: List[str], max_width: float, fontname: str, fontsize: float) -> List[str]:
    lines: List[str] = []
    current = ""
    for token in tokens:
        trial = token if current == "" else current + " " + token
        width = fitz.get_text_length(trial, fontname=fontname, fontsize=fontsize)
        if width <= max_width:
            current = trial
        else:
            if current:
                lines.append(current)
                current = token
            else:
                lines.append(token)
                current = ""
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
    while font_size >= min_font:
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

    lines = _wrap_tokens(tokens, max_width, fontname, min_font)
    page.insert_textbox(
        inner,
        "\n".join(lines),
        fontsize=min_font,
        fontname=fontname,
        color=(0, 0, 0),
        align=0,
    )
