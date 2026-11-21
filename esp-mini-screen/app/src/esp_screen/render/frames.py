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
    name = _fit(d, row.label, "SemiBold", name_size, 120)
    d.text((33, y + 3), name, font=font("SemiBold", name_size), fill=FG)
    if row.updates is not None:  # pending updates: small dot right after the name
        x = 33 + int(d.textlength(name, font=font("SemiBold", name_size))) + 6
        d.ellipse((x, y + 9, x + 7, y + 16), fill=LEVEL_COLOUR[row.updates.value])
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


# --- alert and warning cards (S-3) ------------------------------------------------------------

RED_BG, RED_TEXT, RED_SOFT = (60, 14, 14), (255, 180, 170), (220, 160, 155)
AMBER_BG, AMBER_TEXT, AMBER_SOFT = (58, 44, 8), (255, 214, 120), (215, 185, 120)


def _card(heading: str, title: str, lines: list[str], *, red: bool) -> Image.Image:
    bg, accent, text, soft = (
        (RED_BG, BAD, RED_TEXT, RED_SOFT) if red else (AMBER_BG, WARN, AMBER_TEXT, AMBER_SOFT)
    )
    im = Image.new("RGB", SIZE, bg)
    d = ImageDraw.Draw(im)
    d.rounded_rectangle((10, 10, 230, 230), 16, outline=accent, width=3)
    d.text((120, 40), heading, font=font("Bold", 16), fill=text, anchor="mm")
    size = 30
    while d.textlength(title, font=font("Bold", size)) > 196 and size > 14:
        size -= 1
    if d.textlength(title, font=font("Bold", size)) > 196:
        title = _fit(d, title, "Bold", size, 196)
    d.text((120, 100), title, font=font("Bold", size), fill=FG, anchor="mm")
    for i, line in enumerate(lines):
        if not line:
            continue
        weight, fsize, colour = ("Medium", 16, text) if i == 0 else ("Medium", 13, soft)
        y = 140 if i == 0 else 175 + (i - 1) * 25
        d.text(
            (120, y),
            _fit(d, line, weight, fsize, 196),
            font=font(weight, fsize),
            fill=colour,
            anchor="mm",
        )
    return im


def _ago(minutes: int) -> str:
    if minutes < 60:
        return f"{minutes} min"
    return f"{minutes // 60} h {minutes % 60:02d} min"


def alert(title: str, detail: str, since: str, minutes: int, also: str) -> Image.Image:
    """Red card for a confirmed problem; alternates with the overview."""
    return _card("ATTENTION", title, [f"since {since} ({_ago(minutes)})", detail, also], red=True)


def unreachable(last_good: str | None, minutes: int, *, red: bool) -> Image.Image:
    """No data from Beszel: amber at first, red after the grace period (docs/05 section 4)."""
    when = f"last good read {last_good}" if last_good else "no good read since start"
    return _card(
        "NO DATA", "Beszel unreachable", [f"for {_ago(minutes)}", when, "check the hub"], red=red
    )


def data_changed(detail: str) -> Image.Image:
    """Beszel answered in a shape this app does not know (F-16): never show wrong numbers."""
    return _card(
        "CHECK APP",
        "Beszel data changed",
        ["after a Beszel update?", detail, "see /healthz"],
        red=False,
    )
