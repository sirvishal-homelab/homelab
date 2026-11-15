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


def test_cards_red_and_amber():
    red = frames.alert("node-b down", "Beszel reports it down", "08:17", 13, "also: a stopped")
    amber = frames.unreachable("08:30", 3, red=False)
    late = frames.unreachable(None, 75, red=True)
    data = frames.data_changed("systems/x: field 'mp' is missing " * 3)
    assert red.getpixel((40, 120)) == frames.RED_BG
    assert amber.getpixel((40, 120)) == frames.AMBER_BG
    assert late.getpixel((40, 120)) == frames.RED_BG
    assert data.getpixel((40, 120)) == frames.AMBER_BG
    for im in (red, amber, late, data):
        assert len(frames.encode_jpeg(im)) < 25_000


def test_long_alert_title_fits():
    frames.alert(
        "a-container-with-a-really-long-name stopped", "container on host", "08:00", 125, ""
    )
