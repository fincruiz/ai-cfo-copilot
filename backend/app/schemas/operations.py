from datetime import datetime
from pydantic import BaseModel


class OperationalCheck(BaseModel):
    key: str
    label: str
    status: str
    detail: str
    action: str | None = None


class OperationalReadiness(BaseModel):
    status: str
    score: int
    checks: list[OperationalCheck]
    database_latency_ms: float
    ingestion_open_jobs: int
    ingestion_stale_jobs: int
    ingestion_recent_failures: int
    active_gl_datasets: int
    latest_ingestion_update_at: datetime | None = None
    checked_at: datetime
