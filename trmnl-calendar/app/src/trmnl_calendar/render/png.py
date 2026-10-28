"""Minimal grayscale PNG encoder (colour type 0, bit depth 1, 2, 4 or 8).

Pillow cannot write 2-bit grayscale (it writes a palette PNG), and the spike proved Terminus keeps
a true grayscale PNG unchanged (docs/14), so the project carries its own encoder.
"""

from __future__ import annotations

import struct
import zlib

_SIGNATURE = b"\x89PNG\r\n\x1a\n"


def _chunk(tag: bytes, data: bytes) -> bytes:
    body = tag + data
    return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)


def encode_gray_png(samples: bytes, width: int, height: int, bit_depth: int) -> bytes:
    """samples: one byte per pixel holding the sample value 0..(2**bit_depth - 1)."""
    if bit_depth not in (1, 2, 4, 8):
        raise ValueError("bit_depth must be 1, 2, 4 or 8")
    if len(samples) != width * height:
        raise ValueError("samples length does not match width x height")
    maxv = (1 << bit_depth) - 1
    if samples and max(samples) > maxv:
        raise ValueError("sample value too large for the bit depth")
    per_byte = 8 // bit_depth
    rows = []
    for y in range(height):
        row = samples[y * width : (y + 1) * width]
        out = bytearray([0])  # filter type 0 (none)
        for x in range(0, width, per_byte):
            group = row[x : x + per_byte]
            byte = 0
            for i in range(per_byte):
                v = group[i] if i < len(group) else 0
                byte |= v << (8 - bit_depth * (i + 1))
            out.append(byte)
        rows.append(bytes(out))
    ihdr = struct.pack(">IIBBBBB", width, height, bit_depth, 0, 0, 0, 0)
    return (
        _SIGNATURE
        + _chunk(b"IHDR", ihdr)
        + _chunk(b"IDAT", zlib.compress(b"".join(rows), 9))
        + _chunk(b"IEND", b"")
    )
