# Calendar for TRMNL

A small self-hosted app: add events in a web page, they're stored in SQLite, and
`/calendar.png` renders them as a 1-bit 800×480 image for a TRMNL e-paper display.
Every time the image is requested it is drawn fresh, so the screen always reflects
the current date and your latest events.

## Run it

**Docker (recommended for a home server / Raspberry Pi)**

```bash
# edit TZ in docker-compose.yml first (IANA name, e.g. Europe/London)
docker compose up -d --build
```

**Without Docker**

```bash
python3 -m venv .venv && . .venv/bin/activate
pip install -r requirements.txt
TZ=America/New_York python app.py        # http://<this-machine>:8080
```

Open `http://<machine-ip>:8080`. Events live in `data/events.db`; back up that one file.
`python seed_demo.py` adds sample events if you want to see a populated screen.

## Settings (environment variables)

| Variable      | Default | Meaning                                              |
|---------------|---------|------------------------------------------------------|
| `TZ`          | `UTC`   | Timezone used for "today" and the "Updated" stamp    |
| `TIME_FORMAT` | `12`    | `12` (9:30 AM) or `24` (09:30)                        |
| `WEEK_START`  | `sun`   | `sun` or `mon`                                       |
| `DATA_DIR`    | `./data`| Where `events.db` is stored                          |
| `PORT`        | `8080`  | Port when running `python app.py`                    |
| `FONT_REGULAR`, `FONT_BOLD` | auto | Paths to `.ttf` files if you want other fonts |

## Endpoints

| URL | Purpose |
|-----|---------|
| `/` | Editor + live preview |
| `/calendar.png?view=split\|month\|agenda` | The image. Optional `w` and `h` (default 800×480), and `date=YYYY-MM-DD` to preview another day |
| `/api/events` | CRUD (GET/POST, PUT/DELETE on `/api/events/<id>`) |
| `/api/upcoming?days=30` | Upcoming occurrences with repeats expanded, as JSON |
| `/healthz` | Liveness check |

## Getting the image onto the TRMNL

The app only has to be reachable from the Terminus server at a stable address.
Give this machine a DHCP reservation, then point Terminus at
`http://<machine-ip>:8080/calendar.png`. Check Terminus's `doc/extensions.adoc`
and `doc/api.adoc` for the exact way to pull or push an image, and set the
refresh schedule there.

## Notes

- **No login.** It is meant for your home network. Don't expose port 8080 to the internet.
- Repeating events: daily, weekly, monthly, yearly. A monthly event on the 31st is skipped in shorter months.
- All-day events are marked with a solid bar in the month view. Multi-day events show "Day 1/3" in the agenda.
- Output is pure black and white, drawn with DejaVu Sans Condensed (installed in the Docker image).
