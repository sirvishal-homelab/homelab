"""Overview rendering. Checks layout by sampling pixels, not by hashes, so a different
FreeType build in CI does not break the tests."""

import io

from PIL import Image

from esp_screen.domain.model import Level, Row, RowState, Status, Tile
from esp_screen.render import frames


def status(n_rows, state=RowState.UP):
    rows = tuple(Row(f"node-{i}", state, 12, 40) for i in range(n_rows))
    return Status(rows=rows, containers=Tile(9, 9), others=Tile(3, 3, 1))


def near(px, colour, tol=40):
    return all(abs(a - b) <= tol for a, b in zip(px, colour, strict=True))


def test_three_rows_layout_dots_and_size():
    im = frames.overview(status(3), "08:30")
    assert im.size == (240, 240)
    for y in (102, 147, 192):  # roomy layout: rows start at 102, step 45
        assert near(im.getpixel((22, y + 12)), frames.OK)


def test_four_rows_use_compact_layout():
    im = frames.overview(status(4), "08:30")
    for y in (64, 107, 150, 193):  # compact layout: rows start at 64, step 43
        assert near(im.getpixel((22, y + 12)), frames.OK)


def test_down_row_has_red_dot():
    im = frames.overview(status(1, state=RowState.DOWN), "08:30")
    assert near(im.getpixel((22, 114)), frames.BAD)


def test_paused_and_missing_rows_are_grey():
    for state in (RowState.PAUSED, RowState.MISSING):
        im = frames.overview(status(1, state=state), "08:30")
        assert near(im.getpixel((22, 114)), frames.GREY)


def test_tile_colour_follows_level():
    assert Tile(8, 9).level is Level.BAD
    assert Tile(0, 0, paused=1).level is Level.OK


def test_long_names_and_no_rows_render():
    st = Status(
        rows=(Row("a-very-long-system-name-indeed", RowState.UP, 1, 1),),
        containers=Tile(0, 0),
        others=Tile(0, 0),
    )
    frames.overview(st, "08:30")
    frames.overview(Status(rows=(), containers=Tile(0, 0), others=Tile(0, 0)), "08:30")


def test_overview_jpeg_is_small():
    data = frames.encode_jpeg(frames.overview(status(4), "23:59"))
    assert Image.open(io.BytesIO(data)).size == (240, 240)
    assert len(data) < 25_000


def test_time_change_changes_frame():
    a = frames.encode_jpeg(frames.overview(status(3), "08:30"))
    b = frames.encode_jpeg(frames.overview(status(3), "08:31"))
    assert frames.frame_hash(a) != frames.frame_hash(b)
