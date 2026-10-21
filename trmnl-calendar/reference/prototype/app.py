"""Calendar editor + image endpoint for a self-hosted TRMNL (Terminus) setup."""
import os
from datetime import date, datetime, timedelta
from io import BytesIO
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from flask import Flask, jsonify, render_template, request, send_file

import render
import store

try:
    TZ = ZoneInfo(os.environ.get("TZ") or "UTC")
except ZoneInfoNotFoundError:
    TZ = ZoneInfo("UTC")

app = Flask(__name__)
store.init_db()


def _now():
    now = datetime.now(TZ).replace(tzinfo=None)
    override = request.args.get("date")        # handy for testing: ?date=2026-12-24
    if override:
        try:
            now = datetime.combine(date.fromisoformat(override), now.time())
        except ValueError:
            pass
    return now


def _size():
    def clamp(name, default):
        try:
            return max(200, min(2000, int(request.args.get(name, default))))
        except ValueError:
            return default
    return clamp("w", 800), clamp("h", 480)


@app.get("/")
def index():
    return render_template("index.html", views=render.VIEWS, time_format=render.TIME_FORMAT)


@app.get("/calendar.png")
def calendar_png():
    view = request.args.get("view", "split")
    if view not in render.VIEWS:
        return jsonify(error=f"view must be one of {', '.join(render.VIEWS)}"), 400
    w, h = _size()
    png = render.render_png(view, w, h, _now())
    resp = send_file(BytesIO(png), mimetype="image/png")
    resp.headers["Cache-Control"] = "no-store"
    return resp


# ------------------------------------------------------------------ JSON API

@app.get("/api/events")
def api_list():
    return jsonify(store.list_events())


@app.post("/api/events")
def api_create():
    clean, errors = store.validate(request.get_json(silent=True))
    if errors:
        return jsonify(errors=errors), 400
    return jsonify(store.get_event(store.create_event(clean))), 201


@app.put("/api/events/<int:event_id>")
def api_update(event_id):
    clean, errors = store.validate(request.get_json(silent=True))
    if errors:
        return jsonify(errors=errors), 400
    if not store.update_event(event_id, clean):
        return jsonify(error="Event not found."), 404
    return jsonify(store.get_event(event_id))


@app.delete("/api/events/<int:event_id>")
def api_delete(event_id):
    if not store.delete_event(event_id):
        return jsonify(error="Event not found."), 404
    return "", 204


@app.get("/api/upcoming")
def api_upcoming():
    """Expanded occurrences (recurrences applied) as JSON, e.g. for a Terminus extension."""
    try:
        days = max(1, min(366, int(request.args.get("days", 30))))
    except ValueError:
        days = 30
    today = _now().date()
    by_day = store.events_by_day(today, today + timedelta(days=days))
    return jsonify([
        {"date": d.isoformat(), **e} for d in sorted(by_day) for e in by_day[d]
    ])


@app.get("/healthz")
def healthz():
    return jsonify(status="ok")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
