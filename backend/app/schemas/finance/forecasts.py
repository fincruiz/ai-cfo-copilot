from decimal import Decimal
from uuid import UUID

from pydantic import BaseModel, Field


class ForecastRequest(BaseModel):
    reporting_group: str = Field(default="Revenue", min_length=1)
    future_months: int = Field(default=12, ge=1, le=60)
    method: str = "run_rate"
    branch_id: UUID | None = None
    downside_factor: Decimal = Decimal("0.90")
    upside_factor: Decimal = Decimal("1.10")
    recent_months: int = Field(default=3, ge=1, le=24)


class ForecastPointResponse(BaseModel):
    period: str
    base: Decimal
    downside: Decimal
    upside: Decimal


class ForecastResponse(BaseModel):
    reporting_group: str
    method: str
    branch_id: UUID | None
    history_periods: int
    confidence: str
    warning: str | None
    points: list[ForecastPointResponse]
