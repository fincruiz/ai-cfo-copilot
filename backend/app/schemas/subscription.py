from datetime import datetime
from typing import Any, Literal
from pydantic import BaseModel, Field

PlanName = Literal['trial','founding','growth','enterprise']
SubscriptionStatusName = Literal['trialing','active','past_due','cancelled','expired']
BillingInterval = Literal['monthly','annual']

class SubscriptionStatusOut(BaseModel):
    plan: PlanName
    display_name: str
    status: SubscriptionStatusName
    trial_started_at: datetime | None = None
    trial_ends_at: datetime | None = None
    current_period_ends_at: datetime | None = None
    days_remaining: int | None = Field(default=None, ge=0)
    entitlements: dict[str, Any] = Field(default_factory=dict)
    is_access_active: bool = True
    billing_managed_externally: bool = True
    billing_country_code: str | None = None
    billing_interval: BillingInterval = 'monthly'
    requested_plan: str | None = None
    requested_interval: BillingInterval | None = None
    change_requested_at: datetime | None = None
    cancellation_requested_at: datetime | None = None

class SubscriptionChangeRequest(BaseModel):
    plan: Literal['founding','growth','enterprise']
    billing_interval: BillingInterval = 'monthly'

class BillingMarketRequest(BaseModel):
    country_code: str = Field(min_length=2, max_length=2)

class SubscriptionChangeOut(BaseModel):
    requested_plan: str
    requested_interval: BillingInterval
    change_requested_at: datetime
    message: str

class BetaReadinessCheck(BaseModel):
    key: str
    label: str
    status: Literal['ready','attention','blocked']
    detail: str

class BetaReadinessOut(BaseModel):
    score: int = Field(ge=0,le=100)
    status: Literal['ready','attention','blocked']
    checks: list[BetaReadinessCheck]
