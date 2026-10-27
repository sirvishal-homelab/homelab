import pytest

from trmnl_calendar.config import load_settings


@pytest.fixture
def settings(tmp_path):
    return load_settings({"DATA_DIR": str(tmp_path)})
