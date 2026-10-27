"""Text measuring and fitting with the bundled fonts. No fallback font (docs/04 section 4)."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

from trmnl_calendar.render.fontcheck import FONT_DIR

MIN_FONT_PX = 10
_ELLIPSIS = "…"


@lru_cache(maxsize=32)
def get_font(size: int, bold: bool = False) -> ImageFont.FreeTypeFont:
    if size < MIN_FONT_PX:
        raise ValueError(f"no text under {MIN_FONT_PX} px (got {size})")
    name = "DejaVuSansCondensed-Bold.ttf" if bold else "DejaVuSansCondensed.ttf"
    # basic layout: no dependency on libraqm, so measuring is identical everywhere
    return ImageFont.truetype(
        str(Path(str(FONT_DIR / name))), size, layout_engine=ImageFont.Layout.BASIC
    )


def measure(text: str, size: int, bold: bool = False) -> float:
    return get_font(size, bold).getlength(text)


def ellipsize(text: str, size: int, bold: bool, max_width: float) -> tuple[str, bool]:
    """Truncate by measured pixel width. Returns (text, was_truncated)."""
    if measure(text, size, bold) <= max_width:
        return text, False
    t = text
    while t and measure(t + _ELLIPSIS, size, bold) > max_width:
        t = t[:-1]
    return t.rstrip() + _ELLIPSIS, True


def wrap(
    text: str, size: int, bold: bool, max_width: float, max_lines: int
) -> tuple[list[str], bool]:
    """Wrap on words, at most max_lines; the last line is ellipsized. Returns (lines, truncated)."""
    words = text.split()
    lines: list[str] = []
    cur = ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if measure(trial, size, bold) <= max_width:
            cur = trial
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    truncated = False
    if len(lines) > max_lines:
        rest = " ".join(lines[max_lines - 1 :])
        lines = lines[: max_lines - 1] + [rest]
        truncated = True
    out = []
    for line in lines:
        t, cut = ellipsize(line, size, bold, max_width)
        truncated = truncated or cut
        out.append(t)
    return out, truncated
