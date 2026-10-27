# syntax=docker/dockerfile:1
# Multi-stage, non-root, bundled fonts, HEALTHCHECK. Python 3.12 (ADR-0009).
# TODO(slice 0 follow-up): pin both base images by digest once CI has built once (docs/12 "Reproducible").

FROM python:3.12-slim AS build
ENV UV_COMPILE_BYTECODE=1 UV_LINK_MODE=copy UV_PROJECT_ENVIRONMENT=/opt/venv
RUN pip install --no-cache-dir uv
WORKDIR /src
COPY pyproject.toml README.md uv.loc[k] ./
COPY src ./src
# --locked when a lock file is committed; otherwise resolve at build time (see docs/19)
RUN if [ -f uv.lock ]; then uv sync --no-dev --no-editable --locked; else uv sync --no-dev --no-editable; fi

FROM python:3.12-slim
ENV PATH="/opt/venv/bin:$PATH" PYTHONUNBUFFERED=1 PYTHONDONTWRITEBYTECODE=1 \
    DATA_DIR=/data WEB_CONCURRENCY=1
RUN useradd --system --uid 10001 --home-dir /nonexistent --shell /usr/sbin/nologin app \
 && mkdir /data && chown app:app /data
COPY --from=build /opt/venv /opt/venv
USER app
VOLUME /data
EXPOSE 8080
HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8080/healthz', timeout=4).status == 200 else 1)"
# One process, one worker (ADR-0007). Not --reload, not several workers.
CMD ["uvicorn", "--factory", "trmnl_calendar.main:app_factory", "--host", "0.0.0.0", "--port", "8080", "--workers", "1"]
