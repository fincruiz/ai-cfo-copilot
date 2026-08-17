from datetime import datetime
from typing import Any, Literal

from pydantic import BaseModel, Field

PlanName = Literal["trial", "founding", "growth", "enterprise"]
SubscriptionStatusName = Literal["trialing", "active", "past_due", "cancelled", "expired"]


class SubscriptionStatusOut(BaseModel):
    plan: PlanName
    status: SubscriptionStatusName
    trial_started_at: datetime | None = None
    trial_ends_at: datetime | None = None
    current_period_ends_at: datetime | None = None
    days_remaining: int | None = Field(default=None, ge=0)
    entitlements: dict[str, Any] = Field(default_factory=dict)
    is_access_active: bool = True
    billing_managed_externally: bool = True


class BetaReadinessCheck(BaseModel):
    key: str
    label: str
    status: Literal["ready", "attention", "blocked"]
    detail: str


class BetaReadinessOut(BaseModel):
    score: int = Field(ge=0, le=100)
    status: Literal["ready", "attention", "blocked"]
    checks: list[BetaReadinessCheck]
