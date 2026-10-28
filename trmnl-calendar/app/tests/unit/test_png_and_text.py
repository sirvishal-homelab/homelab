import io
import struct
import zlib

import pytest
from PIL import Image

from trmnl_calendar.render import text as T
from trmnl_calendar.render.png import encode_gray_png
from trmnl_calendar.render.quantise import levels, to_samples


def chunks(png: bytes):
    assert png[:8] == b"\x89PNG\r\n\x1a\n"
    i, out = 8, []
    while i < len(png):
        n = struct.unpack(">I", png[i : i + 4])[0]
        tag, data = png[i + 4 : i + 8], png[i + 8 : i + 8 + n]
        crc = struct.unpack(">I", png[i + 8 + n : i + 12 + n])[0]
        assert crc == zlib.crc32(tag + data) & 0xFFFFFFFF
        out.append((tag, data))
        i += 12 + n
    return out


def test_header_is_true_grayscale_two_bit():
    png = encode_gray_png(bytes([0, 1, 2, 3] * (800 * 480 // 4)), 800, 480, 2)
    tag, ihdr = chunks(png)[0]
    assert tag == b"IHDR"
    w, h, depth, ctype, comp, filt, inter = struct.unpack(">IIBBBBB", ihdr)
    assert (w, h, depth, ctype, comp, filt, inter) == (800, 480, 2, 0, 0, 0, 0)
    assert [t for t, _ in chunks(png)] == [b"IHDR", b"IDAT", b"IEND"]


@pytest.mark.parametrize("depth", [1, 2, 4, 8])
@pytest.mark.parametrize("width", [1, 3, 5, 8, 13])
def test_roundtrip_through_pillow_including_padding(depth, width):
    maxv = (1 << depth) - 1
    samples = bytes((x * 7 + y * 3) % (maxv + 1) for y in range(4) for x in range(width))
    im = Image.open(io.BytesIO(encode_gray_png(samples, width, 4, depth)))
    expected = bytes(round(v * 255 / maxv) for v in samples)
    assert im.size == (width, 4) and im.convert("L").tobytes() == expected


def test_rejects_bad_input():
    with pytest.raises(ValueError):
        encode_gray_png(b"\x00", 2, 2, 2)
    with pytest.raises(ValueError):
        encode_gray_png(b"\x04", 1, 1, 2)
    with pytest.raises(ValueError):
        encode_gray_png(b"\x00", 1, 1, 3)


def test_quantise_nearest_level():
    assert levels(2) == (0, 85, 170, 255)
    assert list(to_samples(bytes([0, 42, 43, 127, 128, 212, 213, 255]), 2)) == [
        0,
        0,
        1,
        1,
        2,
        2,
        3,
        3,
    ]
    assert list(to_samples(bytes([0, 127, 128, 255]), 1)) == [0, 0, 1, 1]


def test_ellipsize_respects_measured_width():
    for w in (30, 60, 100):
        out, cut = T.ellipsize("Quarterly planning meeting with everybody", 12, False, w)
        assert cut and T.measure(out, 12, False) <= w and out.endswith("…")
    assert T.ellipsize("Hi", 12, False, 50) == ("Hi", False)


def test_wrap_limits_lines_and_width():
    lines, cut = T.wrap("one two three four five six seven eight nine ten", 12, False, 70, 2)
    assert len(lines) == 2 and cut and all(T.measure(x, 12, False) <= 70 for x in lines)


def test_no_text_under_ten_px():
    with pytest.raises(ValueError):
        T.get_font(9)
