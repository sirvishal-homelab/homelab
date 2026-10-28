"""Generic rasteriser: draws a Scene with Pillow. Text is crisp (no anti-aliasing), docs/04 s3."""

from __future__ import annotations

from PIL import Image, ImageDraw

from trmnl_calendar.render.scene import Glyph, Line, Rect, Scene, Text
from trmnl_calendar.render.text import get_font

SHAPES = ("square", "triangle", "circle", "diamond", "plus", "cross", "ring", "box")


def _glyph(d: ImageDraw.ImageDraw, g: Glyph) -> None:
    x, y, z = g.x, g.y, g.size
    c = z / 2 - 0.5
    if g.shape == "square":
        d.rectangle((x, y, x + z - 2, y + z - 2), fill=0)
    elif g.shape == "triangle":
        d.polygon([(x + c, y), (x + z - 1, y + z - 1), (x, y + z - 1)], fill=0)
    elif g.shape == "circle":
        d.ellipse((x, y, x + z - 1, y + z - 1), fill=0)
    elif g.shape == "diamond":
        d.polygon([(x + c, y), (x + z - 1, y + c), (x + c, y + z - 1), (x, y + c)], fill=0)
    elif g.shape == "plus":
        d.line((x + c, y, x + c, y + z - 1), fill=0, width=2)
        d.line((x, y + c, x + z - 1, y + c), fill=0, width=2)
    elif g.shape == "cross":
        d.line((x, y, x + z - 1, y + z - 1), fill=0, width=2)
        d.line((x, y + z - 1, x + z - 1, y), fill=0, width=2)
    elif g.shape == "ring":
        d.ellipse((x, y, x + z - 1, y + z - 1), outline=0, width=2)
    else:  # box
        d.rectangle((x, y, x + z - 2, y + z - 2), outline=0, width=2)


def rasterise(scene: Scene) -> bytes:
    """Return width*height bytes of 8-bit grey, row by row."""
    im = Image.new("L", (scene.width, scene.height), scene.background)
    d = ImageDraw.Draw(im)
    d.fontmode = "1"  # crisp glyphs: no anti-aliasing
    for it in scene.items:
        if isinstance(it, Rect):
            d.rectangle(
                (it.x, it.y, it.x + it.w - 1, it.y + it.h - 1), fill=it.fill, outline=it.outline
            )
        elif isinstance(it, Line):
            if it.dotted:
                if it.y1 == it.y2:
                    for x in range(it.x1, it.x2, 6):
                        d.line((x, it.y1, min(x + 2, it.x2), it.y1), fill=it.gray)
                else:
                    for y in range(it.y1, it.y2, 6):
                        d.line((it.x1, y, it.x1, min(y + 2, it.y2)), fill=it.gray)
            else:
                d.line((it.x1, it.y1, it.x2, it.y2), fill=it.gray, width=it.width)
        elif isinstance(it, Text):
            anchor = {"l": "ls", "m": "ms", "r": "rs"}[it.anchor]
            d.text(
                (it.x, it.y), it.text, font=get_font(it.size, it.bold), fill=it.gray, anchor=anchor
            )
        elif isinstance(it, Glyph):
            _glyph(d, it)
        else:  # pragma: no cover - programming error
            raise TypeError(f"unknown scene item {it!r}")
    return im.tobytes()
