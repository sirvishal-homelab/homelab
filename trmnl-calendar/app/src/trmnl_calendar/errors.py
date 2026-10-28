"""One error type for every failure.

The web layer turns it into the single error shape (docs/06 section 5).
"""

from __future__ import annotations


class AppError(Exception):
    status = 400
    code = "bad_request"

    def __init__(
        self, message: str, *, fields: dict[str, str] | None = None, extra: dict | None = None
    ):
        super().__init__(message)
        self.message = message
        self.fields = fields or {}
        self.extra = extra or {}

    def body(self) -> dict:
        err: dict = {"code": self.code, "message": self.message}
        if self.fields:
            err["fields"] = self.fields
        err.update(self.extra)
        return {"error": err}


class NotFound(AppError):
    status, code = 404, "not_found"


class ValidationFailed(AppError):
    status, code = 422, "validation_failed"

    def __init__(self, fields: dict[str, str]):
        super().__init__("Please check the highlighted fields.", fields=fields)


class EditConflict(AppError):
    status, code = 409, "edit_conflict"


class PreconditionRequired(AppError):
    status, code = 428, "precondition_required"
