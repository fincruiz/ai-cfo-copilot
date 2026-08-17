from contextlib import asynccontextmanager
import asyncio

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles

from app.api.v1.router import api_router
from app.core.config import settings
from app.core.exception_handlers import (
    register_exception_handlers,
)
from app.core.logging import configure_logging
from app.database.session import engine
from app.middleware.request_logging import (
    RequestLoggingMiddleware,
)
from app.services.finance.ingestion_job_service import worker_loop


configure_logging()


@asynccontextmanager
async def lifespan(app: FastAPI):
    stop_event = asyncio.Event()
    worker_task = asyncio.create_task(worker_loop(stop_event))
    try:
        yield
    finally:
        stop_event.set()
        worker_task.cancel()
        try:
            await worker_task
        except asyncio.CancelledError:
            pass
        if engine is not None:
            await engine.dispose()


app = FastAPI(
    title=settings.app_name,
    version=settings.app_version,
    debug=settings.debug,
    lifespan=lifespan,
    docs_url=None if settings.is_production else "/docs",
    redoc_url=None if settings.is_production else "/redoc",
    openapi_url=None if settings.is_production else "/openapi.json",
)


app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


app.add_middleware(
    RequestLoggingMiddleware,
)


register_exception_handlers(app)


app.mount("/uploads", StaticFiles(directory="uploads", check_dir=False), name="uploads")

app.include_router(
    api_router,
    prefix=settings.api_v1_prefix,
)


@app.get("/", tags=["Root"])
async def root() -> dict[str, str]:
    return {
        "message": "AI CFO Copilot API",
        "docs": "/docs",
        "health": f"{settings.api_v1_prefix}/health",
    }
