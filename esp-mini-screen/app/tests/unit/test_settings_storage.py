import contextlib

from esp_screen.storage import db, repo


def connect(s):
    return contextlib.closing(db.connect(s.db_path))


SYSTEMS = [
    ("id-a", "node-a"),
    ("id-b", "node-b"),
    ("id-n", "nas"),
    ("id-d", "dns"),
    ("id-v", "vm-1"),
]


def test_sync_adds_and_renames(migrated):
    with connect(migrated) as c:
        repo.sync_systems(c, SYSTEMS)
        repo.sync_systems(c, [("id-a", "node-a-renamed")])
        names = {p.beszel_id: p.beszel_name for p in repo.list_systems(c)}
    assert names["id-a"] == "node-a-renamed"
    assert len(names) == 5


def test_seed_focus_from_env_once(migrated):
    with connect(migrated) as c:
        assert repo.seed_focus_by_name(c, ("NODE-A", "nas")) is False  # systems not known yet
        repo.sync_systems(c, SYSTEMS)
        assert repo.seed_focus_by_name(c, ("NODE-A", "nas", "ghost")) is True
        assert repo.focus_ids(c) == ["id-a", "id-n"]
        repo.save_system_prefs(c, {"id-d": (1, None, False)})
        assert repo.seed_focus_by_name(c, ("node-a",)) is False  # never again
        assert repo.focus_ids(c) == ["id-d"]


def test_save_prefs_labels_and_ignore(migrated):
    with connect(migrated) as c:
        repo.sync_systems(c, SYSTEMS)
        repo.save_system_prefs(
            c,
            {
                "id-n": (1, "NAS", False),
                "id-a": (2, None, False),
                "id-v": (None, None, True),
            },
        )
        assert repo.focus_ids(c) == ["id-n", "id-a"]
        assert repo.labels(c)["id-n"] == "NAS"
        assert repo.labels(c)["id-a"] == "node-a"
        assert repo.ignored_systems(c) == frozenset({"id-v"})


def test_swapping_positions_in_one_save(migrated):
    with connect(migrated) as c:
        repo.sync_systems(c, SYSTEMS)
        repo.save_system_prefs(c, {"id-a": (1, None, False), "id-b": (2, None, False)})
        repo.save_system_prefs(c, {"id-a": (2, None, False), "id-b": (1, None, False)})
        assert repo.focus_ids(c) == ["id-b", "id-a"]


def test_containers_allowed_and_forget(migrated):
    with connect(migrated) as c:
        repo.sync_systems(c, SYSTEMS)
        repo.save_container_memory(c, {("id-d", "web"): 0, ("id-d", "job"): 3})
        repo.save_allowed_to_stop(c, {("id-d", "job")})
        listed = {k.name: k for k in repo.list_containers(c)}
        assert listed["job"].allowed_to_stop and not listed["web"].allowed_to_stop
        assert listed["job"].system_name == "dns"
        assert repo.forget_container(c, "id-d", "job") is True
        assert [k.name for k in repo.list_containers(c)] == ["web"]


def test_clear_push_hash(migrated):
    with connect(migrated) as c:
        repo.record_push(c, ok=True, frame_hash="abc")
        repo.clear_push_hash(c)
        assert repo.get_push_state(c).last_hash is None
