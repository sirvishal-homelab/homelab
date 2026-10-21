"""Optional: fill the database with sample events relative to today.  python seed_demo.py"""
from datetime import date, timedelta
import store

store.init_db()
sec = {s["name"]: s["id"] for s in store.list_sections()}
t = date.today()
d = lambda n: (t + timedelta(days=n)).isoformat()
samples = [
    dict(title="Dentist", date=d(0), start_time="09:30", end_time="10:15", location="Main St Dental"),
    dict(title="Team stand-up", date=d(0), start_time="10:00", repeat="weekly", section="Work"),
    dict(title="Project review", date=d(1), start_time="14:00", section="Work"),
    dict(title="Soup and sandwich", date=d(0), section="Lunch", all_day=True),
    dict(title="Lunch with Sam", date=d(2), start_time="12:30", location="Cafe Nord", section="Lunch"),
    dict(title="Pasta night", date=d(0), section="Dinner", all_day=True),
    dict(title="Tacos", date=d(1), section="Dinner", all_day=True),
    dict(title="Pizza", date=d(3), section="Dinner", all_day=True),
    dict(title="Mum's birthday", date=d(1), repeat="yearly"),
    dict(title="Piano lesson", date=d(2), start_time="16:00", end_time="16:45", location="Studio 4"),
    dict(title="Conference", date=d(4), end_date=d(6), location="Convention Centre", section="Work", all_day=True),
    dict(title="Bin day", date=d(3), repeat="weekly"),
    dict(title="Book club", date=d(8), start_time="19:00", location="Sam's place"),
    dict(title="Car service", date=d(12), start_time="08:00"),
    dict(title="Rent due", date=d(14), repeat="monthly"),
]
for s in samples:
    name = s.pop("section", "Events")
    s["section_id"] = sec[name]
    clean, errors = store.validate(s)
    assert not errors, (s, errors)
    store.create_event(clean)
print(f"Added {len(samples)} sample events.")
