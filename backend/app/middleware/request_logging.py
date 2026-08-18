import logging
import time
from uuid import uuid4

from fastapi import Request
from starlette.middleware.base import BaseHTTPMiddleware
from starlette.responses import Response

from app.core.config import settings


logger = logging.getLogger(__name__)


class RequestLoggingMiddleware(BaseHTTPMiddleware):
    async def dispatch(
        self,
        request: Request,
        call_next,
    ) -> Response:
        request_id = request.headers.get(
            "X-Request-ID",
            str(uuid4()),
        )

        request.state.request_id = request_id

        started_at = time.perf_counter()

        try:
            response = await call_next(request)
        except Exception:
            logger.exception(
                "Request failed | request_id=%s | method=%s | path=%s",
                request_id,
                request.method,
                request.url.path,
            )
            raise

        duration_ms = (
            time.perf_counter() - started_at
        ) * 1000

        response.headers["X-Request-ID"] = request_id
        response.headers["Server-Timing"] = f"app;dur={duration_ms:.2f}"

        log = logger.warning if duration_ms >= settings.slow_request_ms else logger.info
        log(
            "Request completed | request_id=%s | "
            "method=%s | path=%s | status=%s | duration_ms=%.2f | slow=%s",
            request_id,
            request.method,
            request.url.path,
            response.status_code,
            duration_ms,
            duration_ms >= settings.slow_request_ms,
        )

        return response