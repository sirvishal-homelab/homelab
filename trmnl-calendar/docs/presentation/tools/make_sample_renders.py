"""Makes the two e-paper renders used in the talk, from invented sample data.

Run from trmnl-calendar/app:  python3 ../docs/presentation/tools/make_sample_renders.py <out-dir>
The phone screenshots (ui-*.png) were taken in a browser from the pages this script also writes
to <out-dir>/site. Names and events are made up; nothing here is real household data.
"""
import sys, tempfile, pathlib
from datetime import UTC, datetime
sys.path.insert(0, "src")
from trmnl_calendar.storage import db, migrate
from trmnl_calendar.storage.repositories import events as ev, people as pp, sections as sec, views as vw, calendar as cal
from trmnl_calendar.web import render_service

OUT = pathlib.Path(sys.argv[1])
d = pathlib.Path(tempfile.mkdtemp()); p = d / "c.db"
conn = db.connect(p); migrate.migrate(conn, p, d / "b")
cal.update_calendar(conn, {"week_start": "sun"}) if hasattr(cal, "update_calendar") else None
alex = pp.create_person(conn, "Alex", "A"); sam = pp.create_person(conn, "Sam", "S"); kid = pp.create_person(conn, "Robin", "R")
hol = sec.create_section(conn, "Holidays", placement="banner")
def add(**k):
    base = dict(section_id=1); base.update(k)
    try:
        return ev.create_event(conn, base)
    except Exception as e:
        print(k["title"], getattr(e, "fields", e))
        raise
S = lambda n: next(s.id for s in cal.list_sections(conn) if s.name == n)
# Tue 27 Oct 2026 is "today"
add(title="Dentist", section_id=S("Events"), start_date="2026-10-27", start_time="09:30", people_ids=[alex.id], owner_id=alex.id)
add(title="Piano lesson", section_id=S("Events"), start_date="2026-10-29", start_time="16:00", people_ids=[kid.id], owner_id=kid.id)
add(title="Book club", section_id=S("Events"), start_date="2026-10-28", start_time="19:00", people_ids=[sam.id, alex.id], owner_id=sam.id)
add(title="Bin day", section_id=S("Events"), start_date="2026-10-30", freq="weekly", by_weekday=["FR"])
add(title="Team stand-up", section_id=S("Work"), start_date="2026-10-27", start_time="10:00", people_ids=[alex.id], owner_id=alex.id)
add(title="Project review", section_id=S("Work"), start_date="2026-10-28", start_time="14:00", people_ids=[alex.id], owner_id=alex.id)
add(title="Conference", section_id=S("Work"), start_date="2026-10-29", end_date="2026-10-31", people_ids=[sam.id], owner_id=sam.id)
add(title="Lunch with Jo", section_id=S("Lunch"), start_date="2026-10-29", start_time="12:30", people_ids=[alex.id], owner_id=alex.id)
add(title="Pasta night", section_id=S("Dinner"), start_date="2026-10-27")
add(title="Tacos", section_id=S("Dinner"), start_date="2026-10-28")
add(title="Halloween", section_id=hol.id, start_date="2026-10-31")
add(title="Clocks back", section_id=hol.id, start_date="2026-11-01")
add(title="Remembrance Day", section_id=hol.id, start_date="2026-11-11")
add(title="Work", section_id=S("Work"), start_date="2026-11-02", people_ids=[sam.id], owner_id=sam.id,
    rotation=[["WE", "TH"], ["MO", "TU", "FR", "SA", "SU"]])
add(title="Book club", section_id=S("Events"), start_date="2026-11-11", start_time="19:00", people_ids=[sam.id], owner_id=sam.id)
now = datetime(2026, 10, 27, 15, 0, tzinfo=UTC)
for tpl, name in (("week_lanes", "week"), ("twoweek_list", "twoweek")):
    v = vw.create_view(conn, tpl, {"name": "Deck " + name}); vw.activate_view(conn, v.id)
    (OUT / f"render-{name}.png").write_bytes(render_service.render_view(conn, v, now).png)
print("ok")

# ---- web pages
from trmnl_calendar.web import views, forms
site = OUT / "site"; (site / "static").mkdir(parents=True, exist_ok=True)
import shutil
for f in pathlib.Path("src/trmnl_calendar/web/static").iterdir():
    shutil.copy(f, site / "static" / f.name)
chrome = {"me": None, "request_path": "/"}
mctx = views.month_context(conn, now, 2026, 11, None, "2026-11-04")
(site / "month.html").write_text(views.render_page("month.html", nav="month", flash="", flash_deleted=None, back="/month", **chrome, **mctx))
work = next(e for e in ev.list_events(conn, __import__("datetime").date(2026, 11, 1), __import__("datetime").date(2026, 11, 30)) if e.title == "Work")
vals = forms.event_to_values(work)
fctx = views.form_context(conn, vals, event=work, skips=[])
(site / "form.html").write_text(views.render_page("event_form.html", back="", **chrome, **fctx))
print("pages ok")
