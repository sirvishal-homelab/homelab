import contextlib

import pytest

from esp_screen.config import load_settings
from esp_screen.storage import db, migrate


@pytest.fixture
def settings(tmp_path):
    return load_settings({"DATA_DIR": str(tmp_path), "SCREEN_URL": "http://screen.test"})


@pytest.fixture
def migrated(settings):
    """A database at the latest schema, with the screen address seeded."""
    from esp_screen.storage import repo

    with contextlib.closing(db.connect(settings.db_path)) as conn:
        migrate.migrate(conn, settings.db_path, settings.backup_dir)
        repo.seed_device_url(conn, settings.screen_url)
    return settings
