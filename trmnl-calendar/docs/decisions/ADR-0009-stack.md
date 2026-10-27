# ADR-0009: Python 3.12 with FastAPI, server-rendered pages and no build step
Status: **Accepted 2026-10-27 by the owner** (design sign-off)
Date: 2026-10-27

## Context
Python is confirmed. The app is small, runs as one process on a 2 GiB container, serves a phone-first UI to
two people, renders images with Pillow, and must be easy to test and rebuild.

## Decision
- **Python 3.12**, one **uvicorn** worker, **FastAPI** with Pydantic v2 (typed, partial-update models,
  generated OpenAPI used by contract tests). Blocking work (rendering, the Terminus call) runs in a thread
  pool so the image route stays responsive during a push (`02`, `04`).
- **SQLite via the standard library `sqlite3`**, numbered SQL migrations (ADR-0007, `03`). No ORM.
- **Server-rendered pages (Jinja2) with HTMX**, vendored locally, plus a small script for drafts and the
  preview. **No Node toolchain, no front-end build step.**
- **Pillow** for drawing; **httpx** for the Terminus client; **uv** with a committed lock file; **ruff** for
  lint and format; **pytest**; **Playwright** for browser tests.
- One container image from a multi-stage Dockerfile, non-root user, bundled fonts, `HEALTHCHECK`.

## Consequences
Good: few moving parts; no JavaScript build; typed boundaries; same language as the renderer.
Bad: HTMX plus a little script limits very rich interactions; pinned FreeType and Pillow are needed for
stable image snapshots; Playwright is a heavy CI-only dependency.

## Alternatives considered
Flask (simpler, weaker validation, no generated contract); a single-page JavaScript app (build toolchain,
more to maintain); SQLAlchemy and Alembic (more machinery than one small schema needs).
