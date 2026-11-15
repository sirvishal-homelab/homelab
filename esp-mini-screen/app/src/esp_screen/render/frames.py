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
