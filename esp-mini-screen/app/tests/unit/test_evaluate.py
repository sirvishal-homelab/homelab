from datetime import UTC, datetime

import pytest

from esp_screen.domain.evaluate import (
    ContainerMemory,
    Thresholds,
    evaluate,
    find_focus,
    level,
    remember,
)
from esp_screen.domain.model import (
    Container,
    Level,
    RowState,
    Snapshot,
    System,
    SystemState,
)

T0 = datetime(2026, 11, 15, 12, 0, tzinfo=UTC)


def sys(name, state="up", cpu=1.0, mem=10.0, updates=None):
    return System(
        id=f"id-{name}",
        name=name,
        state=SystemState(state),
        cpu_pct=cpu,
        mem_pct=mem,
        updates=updates,
    )


def snapshot(systems, containers=()):
    return Snapshot(tuple(systems), tuple(containers), T0)


@pytest.mark.parametrize(
    "pct, expected", [(0, Level.OK), (74.9, Level.OK), (75, Level.WARN), (90, Level.BAD)]
)
def test_levels(pct, expected):
    assert level(pct, Thresholds(75, 90)) is expected


def test_focus_matches_by_name_or_id_and_keeps_order():
    systems = [sys("node-a"), sys("node-b"), sys("nas")]
    found = find_focus(systems, ["NAS", "id-node-a", "ghost"])
    assert [getattr(f, "name", f) for f in found] == ["nas", "node-a", "ghost"]


def test_focus_limited_to_four():
    systems = [sys(f"n{i}") for i in range(6)]
    assert len(find_focus(systems, [f"n{i}" for i in range(6)])) == 4


def test_rows_by_state():
    snap = snapshot(
        [sys("a", cpu=80, mem=95), sys("b", "down"), sys("c", "paused"), sys("d", "pending")]
    )
    st = evaluate(snap, ["a", "b", "c", "d"], {})
    states = [(r.label, r.state) for r in st.rows]
    assert states == [
        ("a", RowState.UP),
        ("b", RowState.DOWN),
        ("c", RowState.PAUSED),
        ("d", RowState.PENDING),
    ]
    assert st.rows[0].cpu_level is Level.WARN
    assert st.rows[0].mem_level is Level.BAD


def test_missing_focus_system_row():
    st = evaluate(snapshot([sys("a")]), ["ghost"], {})
    assert st.rows[0].state is RowState.MISSING


def test_other_systems_exclude_focus_and_paused():
    snap = snapshot(
        [sys("node-a"), sys("dns"), sys("vm-1", "down"), sys("backup", "paused"), sys("host")]
    )
    st = evaluate(snap, ["node-a"], {})
    assert (st.others.up, st.others.total, st.others.paused) == (2, 3, 1)
    assert st.others.level is Level.BAD


def test_paused_alone_is_never_a_problem():
    st = evaluate(snapshot([sys("node-a"), sys("backup", "paused")]), ["node-a"], {})
    assert (st.others.up, st.others.total, st.others.paused) == (0, 0, 1)
    assert st.others.level is Level.OK


def c(system, name):
    return Container(f"id-{system}", name)


def test_container_memory_learns_and_counts_missing():
    mem = remember(ContainerMemory({}), [c("host", "web"), c("host", "db")], set())
    assert mem == {("id-host", "web"): 0, ("id-host", "db"): 0}
    mem = remember(ContainerMemory(mem), [c("host", "web")], set())
    assert mem[("id-host", "db")] == 1
    mem = remember(ContainerMemory(mem), [c("host", "web")], set())
    assert mem[("id-host", "db")] == 2
    mem = remember(ContainerMemory(mem), [c("host", "web"), c("host", "db")], set())
    assert mem[("id-host", "db")] == 0


def test_containers_on_a_system_that_is_not_up_are_not_counted_missing():
    mem = {("id-nas", "agent"): 0}
    assert remember(ContainerMemory(mem), [], {"id-nas"}) == mem


def test_container_tile_grace_and_allowed_to_stop():
    snap = snapshot([sys("host")], [c("host", "web")])
    one_miss = {("id-host", "web"): 0, ("id-host", "db"): 1}
    st = evaluate(snap, [], one_miss, missing_limit=2)
    assert (st.containers.up, st.containers.total) == (2, 2)  # one slow read: still fine
    two_miss = {("id-host", "web"): 0, ("id-host", "db"): 2}
    st = evaluate(snap, [], two_miss, missing_limit=2)
    assert (st.containers.up, st.containers.total) == (1, 2)
    assert st.down_containers == ("host/db",)
    st = evaluate(
        snap,
        [],
        two_miss,
        missing_limit=2,
        allowed_to_stop=frozenset(two_miss) - {("id-host", "web")},
    )
    assert (st.containers.up, st.containers.total) == (1, 1)
    assert st.down_containers == ()


def test_display_labels():
    st = evaluate(snapshot([sys("srv")]), ["srv"], {}, labels={"id-srv": "NAS"})
    assert st.rows[0].label == "NAS"


@pytest.mark.parametrize(
    "updates, expected",
    [((3, 1), Level.BAD), ((3, 0), Level.WARN), ((0, 0), None), (None, None)],
)
def test_update_dot_level(updates, expected):
    st = evaluate(snapshot([sys("a", updates=updates)]), ["a"], {})
    assert st.rows[0].updates is expected


def test_update_dot_only_on_up_rows_and_can_be_turned_off():
    snap = snapshot([sys("a", updates=(2, 1)), sys("b", "down", updates=(2, 1))])
    st = evaluate(snap, ["a", "b"], {})
    assert [r.updates for r in st.rows] == [Level.BAD, None]
    assert evaluate(snap, ["a"], {}, show_updates=False).rows[0].updates is None


def test_pending_updates_never_raise_an_alert():
    st = evaluate(snapshot([sys("a", updates=(9, 9)), sys("b", updates=(9, 9))]), ["a"], {})
    assert st.down_systems == () and st.others.level is Level.OK
