import dataclasses

import pytest
from fastapi.testclient import TestClient

from trmnl_calendar.main import create_app
from trmnl_calendar.storage import migrate


def test_healthz_ok_and_first_backup(settings):
    with TestClient(create_app(settings)) as client:
        r = client.get("/healthz")
        assert r.status_code == 200
        body = r.json()
        assert body["status"] == "ok" and body["schema_version"] == len(migrate.available())
        assert body["scheduler"] == "off"  # Terminus is not configured in tests
    assert list(settings.backup_dir.glob("calendar-*-auto.db"))


def test_refuses_multiple_workers(settings):
    bad = dataclasses.replace(settings, web_concurrency=2)
    with pytest.raises(RuntimeError, match="WEB_CONCURRENCY"):
        with TestClient(create_app(bad)):
            pass
