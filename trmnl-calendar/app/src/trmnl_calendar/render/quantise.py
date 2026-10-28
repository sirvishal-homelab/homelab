"""Map 8-bit grey to the profile's levels (docs/04 section 3)."""

from __future__ import annotations

LEVELS = {1: (0, 255), 2: (0, 85, 170, 255)}


def levels(bit_depth: int) -> tuple[int, ...]:
    if bit_depth in LEVELS:
        return LEVELS[bit_depth]
    n = 1 << bit_depth
    return tuple(round(i * 255 / (n - 1)) for i in range(n))


def to_samples(gray: bytes, bit_depth: int) -> bytes:
    """8-bit grey values -> sample indices (0 = black), by nearest level."""
    lv = levels(bit_depth)
    table = bytes(min(range(len(lv)), key=lambda i: (abs(lv[i] - v), i)) for v in range(256))
    return gray.translate(table)
