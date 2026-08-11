from fastapi import APIRouter
from sqlalchemy import text

from app.core.config import settings
from app.database.session import engine


router = APIRouter(tags=["Health"])


@router.get("/health")
async def health_check() -> dict[str, str]:
    return {
        "status": "healthy",
        "application": settings.app_name,
        "environment": settings.environment,
        "version": settings.app_version,
    }


@router.get("/health/database")
async def database_health_check() -> dict[str, str]:
    if engine is None:
        return {
            "status": "not_configured",
            "database": "DATABASE_URL is missing",
        }

    try:
        async with engine.connect() as connection:
            await connection.execute(text("SELECT 1"))

        return {
            "status": "healthy",
            "database": "connected",
        }

    except Exception as exc:
        return {
            "status": "unhealthy",
            "database": "connection_failed",
            "detail": str(exc),
        }