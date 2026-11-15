from pathlib import Path

from esp_screen.config import load_settings


def test_defaults():
    s = load_settings({})
    assert s.data_dir == Path("/data")
    assert s.db_path == Path("/data/screen.db")
    assert s.backup_keep == 7
    assert s.timezone == "America/Halifax"
    assert not s.beszel_configured


def test_urls_trimmed_and_secret_hidden():
    s = load_settings(
        {
            "SCREEN_URL": " http://screen.test/ ",
            "BESZEL_URL": "http://beszel:8090/",
            "BESZEL_USER": "app@example.com",
            "BESZEL_PASSWORD": "s3cret",
        }
    )
    assert s.screen_url == "http://screen.test"
    assert s.beszel_url == "http://beszel:8090"
    assert s.beszel_configured
    assert "s3cret" not in repr(s)
