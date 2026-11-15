import io

import pytest
from PIL import Image

from esp_screen.render import fonts, frames


def test_fonts_present():
    fonts.verify_fonts()


def test_hello_frame_is_small_baseline_jpeg():
    data = frames.encode_jpeg(frames.hello("0.1.0", "07:15"))
    im = Image.open(io.BytesIO(data))
    assert im.format == "JPEG"
    assert im.size == (240, 240)
    assert not im.info.get("progressive")  # the ESP8266 decoder needs baseline JPEG
    assert len(data) < 20_000  # small uploads keep pushes fast (docs/03)


def test_rendering_is_deterministic():
    a = frames.encode_jpeg(frames.hello("0.1.0", "07:15"))
    b = frames.encode_jpeg(frames.hello("0.1.0", "07:15"))
    c = frames.encode_jpeg(frames.hello("0.1.0", "07:16"))
    assert frames.frame_hash(a) == frames.frame_hash(b)
    assert frames.frame_hash(a) != frames.frame_hash(c)


def test_wrong_size_rejected():
    with pytest.raises(ValueError):
        frames.encode_jpeg(Image.new("RGB", (100, 100)))
