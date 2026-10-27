import pytest

from trmnl_calendar.render import fontcheck


def test_bundled_fonts_verify():
    assert len(fontcheck.verify_fonts()) == 2


def test_missing_font_fails(monkeypatch, tmp_path):
    monkeypatch.setattr(fontcheck, "FONT_DIR", tmp_path)
    (tmp_path / "SHA256SUMS").write_text("")
    with pytest.raises(fontcheck.FontError):
        fontcheck.verify_fonts()
