from __future__ import annotations

from time import perf_counter

from fastapi import APIRouter, Response, status
from sqlalchemy import text

from app.core.config import settings
from app.database.session import engine

router = APIRouter(tags=["Health"])


def _grade_latency(milliseconds: float) -> str:
    if milliseconds < 250:
        return "healthy"
    if milliseconds < 1000:
        return "degraded"
    return "unhealthy"


@router.get("/health")
async def health_check() -> dict[str, str]:
    return {"status": "healthy", "application": settings.app_name, "environment": settings.environment, "version": settings.app_version}


@router.get("/health/database")
async def database_health_check(response: Response) -> dict[str, str | float]:
    if engine is None:
        response.status_code = status.HTTP_503_SERVICE_UNAVAILABLE
        return {"status": "not_configured", "database": "not_configured", "latency_ms": 0.0}
    started = perf_counter()
    try:
        async with engine.connect() as connection:
            await connection.execute(text("SELECT 1"))
        latency_ms = round((perf_counter() - started) * 1000, 2)
        health = _grade_latency(latency_ms)
        if health == "unhealthy":
            response.status_code = status.HTTP_503_SERVICE_UNAVAILABLE
        return {"status": health, "database": "connected", "latency_ms": latency_ms}
    except Exception:
        response.status_code = status.HTTP_503_SERVICE_UNAVAILABLE
        # Never expose connection strings, driver errors or infrastructure details publicly.
        return {"status": "unhealthy", "database": "connection_failed", "latency_ms": round((perf_counter() - started) * 1000, 2)}


@router.get("/health/readiness")
async def readiness_check(response: Response) -> dict[str, object]:
    checks: dict[str, object] = {"api": {"status": "healthy"}}
    overall = "healthy"
    if engine is None:
        checks["database"] = {"status": "not_configured", "latency_ms": 0.0}
        overall = "unhealthy"
    else:
        started = perf_counter()
        try:
            async with engine.connect() as connection:
                await connection.execute(text("SELECT 1"))
            latency_ms = round((perf_counter() - started) * 1000, 2)
            db_status = _grade_latency(latency_ms)
            checks["database"] = {"status": db_status, "latency_ms": latency_ms}
            if db_status != "healthy": overall = db_status
        except Exception:
            checks["database"] = {"status": "unhealthy", "latency_ms": round((perf_counter() - started) * 1000, 2)}
            overall = "unhealthy"
    if overall == "unhealthy": response.status_code = status.HTTP_503_SERVICE_UNAVAILABLE
    return {"status": overall, "version": settings.app_version, "environment": settings.environment, "checks": checks}
