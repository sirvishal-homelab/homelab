"""Turn every failure into the single error shape for the API, and a plain page for browsers."""

from __future__ import annotations

from fastapi import FastAPI, Request
from fastapi.exceptions import RequestValidationError
from fastapi.responses import HTMLResponse, JSONResponse
from starlette.exceptions import HTTPException as StarletteHTTPException

from trmnl_calendar.errors import AppError
from trmnl_calendar.web import views

_CODES = {404: "not_found", 405: "bad_request"}


def _is_api(request: Request) -> bool:
    return request.url.path.startswith("/api/")


def _page(status: int, title: str, message: str) -> HTMLResponse:
    return HTMLResponse(
        views.render_page("error.html", title=title, message=message), status_code=status
    )


def install(app: FastAPI) -> None:
    @app.exception_handler(AppError)
    async def app_error(request: Request, exc: AppError):
        if _is_api(request):
            return JSONResponse(exc.body(), status_code=exc.status)
        titles = {404: "Not found", 409: "Please reload", 422: "Please check the form"}
        return _page(exc.status, titles.get(exc.status, "Something went wrong"), exc.message)

    @app.exception_handler(StarletteHTTPException)
    async def http_error(request: Request, exc: StarletteHTTPException):
        code = _CODES.get(exc.status_code, "bad_request")
        message = (
            "That page or resource does not exist." if exc.status_code == 404 else str(exc.detail)
        )
        if _is_api(request):
            return JSONResponse(
                {"error": {"code": code, "message": message}}, status_code=exc.status_code
            )
        return _page(exc.status_code, "Not found" if exc.status_code == 404 else "Error", message)

    @app.exception_handler(RequestValidationError)
    async def validation_error(request: Request, exc: RequestValidationError):
        fields = {".".join(str(p) for p in e["loc"][1:]) or "body": e["msg"] for e in exc.errors()}
        malformed = any(
            e["type"] in ("json_invalid", "missing") and e["loc"][:1] == ("body",)
            for e in exc.errors()
        )
        status, code = (400, "bad_request") if malformed else (422, "validation_failed")
        msg = (
            "The request body is not valid JSON."
            if malformed
            else "Please check the highlighted fields."
        )
        return JSONResponse(
            {"error": {"code": code, "message": msg, "fields": fields}}, status_code=status
        )
