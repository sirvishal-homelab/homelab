"""Pure rendering: data in, 240x240 RGB image out. No I/O beyond the bundled fonts."""

from __future__ import annotations

import hashlib
import io

from PIL import Image, ImageDraw

from esp_screen.render.fonts import font

SIZE = (240, 240)
BG, CARD, FG, DIM = (14, 17, 22), (28, 33, 41), (235, 238, 242), (140, 148, 160)
OK, WARN, BAD = (46, 204, 113), (241, 196, 15), (231, 76, 60)


def hello(version: str, started_at: str) -> Image.Image:
    """Start-up frame (S-1): proves upload and display work end to end."""
    im = Image.new("RGB", SIZE, BG)
    d = ImageDraw.Draw(im)
    d.text((10, 8), "HOMELAB", font=font("Bold", 15), fill=DIM)
    d.text((230, 8), started_at, font=font("SemiBold", 15), fill=DIM, anchor="ra")
    d.rounded_rectangle((10, 40, 230, 200), 16, fill=CARD)
    d.ellipse((108, 62, 132, 86), fill=OK)
    d.text((120, 116), "Screen connected", font=font("Bold", 20), fill=FG, anchor="mm")
    d.text((120, 146), "waiting for Beszel", font=font("Medium", 14), fill=DIM, anchor="mm")
    d.text(
        (120, 176), f"esp-mini-screen {version}", font=font("Regular", 11), fill=DIM, anchor="mm"
    )
    d.text((120, 222), f"started {started_at}", font=font("Medium", 12), fill=DIM, anchor="mm")
    return im


def encode_jpeg(im: Image.Image, quality: int = 85) -> bytes:
    """Baseline (non-progressive) JPEG, which the ESP8266 decoder handles."""
    if im.size != SIZE:
        raise ValueError(f"frame must be {SIZE}, got {im.size}")
    buf = io.BytesIO()
    im.convert("RGB").save(buf, "JPEG", quality=quality, optimize=True, progressive=False)
    return buf.getvalue()


def frame_hash(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()[:16]


# --- overview (docs/05 section 5, mock-ups in docs/assets) ------------------------------------

LEVEL_COLOUR = {"ok": OK, "warn": WARN, "bad": BAD}
GREY = (90, 96, 108)
ROW_TEXT = {
    "down": ("DOWN", BAD),
    "paused": ("PAUSED", GREY),
    "pending": ("PENDING", WARN),
    "missing": ("NOT FOUND", GREY),
}
DOT = {"up": OK, "down": BAD, "paused": GREY, "pending": WARN, "missing": GREY}


def _fit(d: ImageDraw.ImageDraw, text: str, weight: str, size: int, width: int) -> str:
    f = font(weight, size)
    if d.textlength(text, font=f) <= width:
        return text
    while text and d.textlength(text + "…", font=f) > width:
        text = text[:-1]
    return text + "…"


def _bar(d: ImageDraw.ImageDraw, x: int, y: int, w: int, pct: float, colour) -> None:
    d.rounded_rectangle((x, y, x + w, y + 8), 4, fill=(50, 56, 66))
    d.rounded_rectangle((x, y, x + max(8, int(w * min(pct, 100) / 100)), y + 8), 4, fill=colour)


def _row(d: ImageDraw.ImageDraw, y: int, row, name_size: int) -> None:
    d.rounded_rectangle((10, y, 230, y + 39), 8, fill=CARD)
    d.ellipse((17, y + 7, 27, y + 17), fill=DOT[row.state.value])
    d.text(
        (33, y + 3),
        _fit(d, row.label, "SemiBold", name_size, 120),
        font=font("SemiBold", name_size),
        fill=FG,
    )
    if row.state.value != "up":
        text, colour = ROW_TEXT[row.state.value]
        d.text((222, y + 5), text, font=font("Bold", 13), fill=colour, anchor="ra")
        return
    for j, (lab, pct, lvl) in enumerate(
        (("CPU", row.cpu_pct, row.cpu_level), ("MEM", row.mem_pct, row.mem_level))
    ):
        bx = 18 + j * 106
        d.text((bx, y + 22), lab, font=font("Medium", 10), fill=DIM)
        _bar(d, bx + 26, y + 25, 46, pct, LEVEL_COLOUR[lvl.value])
        d.text((bx + 100, y + 21), f"{pct:.0f}%", font=font("SemiBold", 11), fill=FG, anchor="ra")


def overview(status, updated: str) -> Image.Image:
    """Up to 3 focus rows: roomy tiles. 4 rows: compact strip (owner decision 2026-11-15)."""
    im = Image.new("RGB", SIZE, BG)
    d = ImageDraw.Draw(im)
    ct, ot = status.containers, status.others
    if len(status.rows) <= 3:
        d.text((10, 8), "HOMELAB", font=font("Bold", 15), fill=DIM)
        d.text((230, 8), updated, font=font("SemiBold", 15), fill=DIM, anchor="ra")
        for i, (label, tile) in enumerate((("Containers", ct), ("Other systems", ot))):
            x = 10 + i * 113
            colour = LEVEL_COLOUR[tile.level.value]
            d.rounded_rectangle((x, 32, x + 107, 92), 10, fill=CARD)
            d.text((x + 10, 38), label, font=font("Medium", 12), fill=DIM)
            d.text((x + 10, 54), str(tile.up), font=font("Bold", 28), fill=colour)
            tx = x + 10 + d.textlength(str(tile.up), font=font("Bold", 28)) + 4
            d.text((tx, 66), f"/ {tile.total}", font=font("Medium", 15), fill=DIM)
            if tile.paused:
                d.text(
                    (x + 100, 74),
                    f"{tile.paused} paused",
                    font=font("Medium", 10),
                    fill=DIM,
                    anchor="ra",
                )
        y, step, name_size = 102, 45, 14
    else:
        d.text((10, 6), "HOMELAB", font=font("Bold", 13), fill=DIM)
        d.text((230, 6), updated, font=font("SemiBold", 13), fill=DIM, anchor="ra")
        d.rounded_rectangle((10, 26, 230, 58), 8, fill=CARD)
        x = 18
        for label, tile in (("Containers", ct), ("Others", ot)):
            colour = LEVEL_COLOUR[tile.level.value]
            d.text((x, 33), label, font=font("Medium", 11), fill=DIM)
            lx = x + d.textlength(label, font=font("Medium", 11)) + 6
            d.text((lx, 29), str(tile.up), font=font("Bold", 19), fill=colour)
            lx += d.textlength(str(tile.up), font=font("Bold", 19)) + 2
            d.text((lx, 34), f"/{tile.total}", font=font("Medium", 12), fill=DIM)
            if tile.paused:
                lx += d.textlength(f"/{tile.total}", font=font("Medium", 12)) + 4
                d.text((lx, 35), f"+{tile.paused}p", font=font("Medium", 9), fill=DIM)
            x = 128
        y, step, name_size = 64, 43, 13
    for row in status.rows:
        _row(d, y, row, name_size)
        y += step
    if not status.rows:
        d.text((120, 160), "No focus systems set", font=font("Medium", 13), fill=DIM, anchor="mm")
    return im
