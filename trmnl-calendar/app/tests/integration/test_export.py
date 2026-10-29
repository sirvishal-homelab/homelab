from datetime import UTC, datetime

import pytest

from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as repo
from trmnl_calendar.storage.repositories.export import TABLES, export_all


@pytest.fixture
def conn(settings):
    c = db.connect(settings.db_path)
    migrate.migrate(c, settings.db_path, settings.backup_dir)
    yield c
    c.close()


def test_export_has_every_table_and_the_events(conn):
    repo.create_event(
        conn, dict(title="Dentist", section_id=1, start_date="2026-03-10", freq="weekly", count=3)
    )
    out = export_all(conn, "9.9.9", datetime(2026, 10, 28, 12, 0, tzinfo=UTC))
    assert out["format"] == "trmnl-calendar-export" and out["app_version"] == "9.9.9"
    assert out["exported_at"] == "2026-10-28T12:00:00+00:00"
    assert set(out["tables"]) == set(TABLES)
    assert [e["title"] for e in out["tables"]["events"]] == ["Dentist"]
    assert out["tables"]["events"][0]["repeat_count"] == 3
    assert [s["name"] for s in out["tables"]["sections"]][0] == "Events"


def test_export_leaves_out_internal_state(conn):
    out = export_all(conn, "1")
    assert not {"push_state", "idempotency_keys", "app_meta"} & set(out["tables"])
