from datetime import date
from decimal import Decimal
from pydantic import BaseModel, Field


class PlanImportResponse(BaseModel):
    plan_type: str
    version_name: str
    total_rows: int
    inserted_rows: int
    invalid_rows: int
    issues: list[dict] = Field(default_factory=list)


class PlanLineResponse(BaseModel):
    period: date
    plan_type: str
    version_name: str
    reporting_group: str
    reporting_subgroup: str | None = None
    source_account_code: str | None = None
    branch_id: str | None = None
    amount: Decimal


class VarianceLineResponse(BaseModel):
    period: date
    reporting_group: str
    actual: Decimal
    budget: Decimal
    forecast: Decimal
    budget_variance: Decimal
    budget_variance_percent: Decimal | None
    forecast_variance: Decimal
    forecast_variance_percent: Decimal | None
