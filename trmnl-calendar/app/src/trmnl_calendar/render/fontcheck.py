"""Startup check for the bundled fonts. No silent fallback to another font (docs/04 section 4)."""

from __future__ import annotations

import hashlib
from importlib import resources

FONT_DIR = resources.files("trmnl_calendar.render") / "fonts"
REQUIRED = ("DejaVuSansCondensed.ttf", "DejaVuSansCondensed-Bold.ttf")


class FontError(RuntimeError):
    pass


def verify_fonts() -> dict[str, str]:
    sums = {}
    for line in (FONT_DIR / "SHA256SUMS").read_text().splitlines():
        digest, _, name = line.partition("  ")
        sums[name.strip()] = digest
    ok = {}
    for name in REQUIRED:
        f = FONT_DIR / name
        if not f.is_file():
            raise FontError(f"bundled font missing: {name}")
        actual = hashlib.sha256(f.read_bytes()).hexdigest()
        if sums.get(name) != actual:
            raise FontError(f"bundled font checksum mismatch: {name}")
        ok[name] = actual
    return ok
