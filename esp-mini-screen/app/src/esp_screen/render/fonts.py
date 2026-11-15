"""Bundled Inter fonts (SIL OFL 1.1, see fonts/OFL.txt), so rendering is identical everywhere."""

from __future__ import annotations

from functools import lru_cache
from pathlib import Path

from PIL import ImageFont

FONT_DIR = Path(__file__).parent / "fonts"
WEIGHTS = ("Regular", "Medium", "SemiBold", "Bold")


@lru_cache(maxsize=64)
def font(weight: str, size: int) -> ImageFont.FreeTypeFont:
    if weight not in WEIGHTS:
        raise ValueError(f"unknown weight {weight!r}")
    return ImageFont.truetype(str(FONT_DIR / f"Inter-{weight}.otf"), size)


def verify_fonts() -> None:
    """Fail at startup, not at the first render, if a font file is missing."""
    missing = [w for w in WEIGHTS if not (FONT_DIR / f"Inter-{w}.otf").is_file()]
    if missing:
        raise RuntimeError(f"bundled fonts missing: {missing}")
    font("Regular", 12)
