from datetime import date
from decimal import Decimal
from typing import Any

from pydantic import BaseModel, Field


class ImportIssue(BaseModel):
    row_number: int | None = None
    column: str | None = None
    message: str
    severity: str = "error"


class FinanceImportResponse(BaseModel):
    import_type: str
    original_file_name: str
    total_rows: int
    valid_rows: int
    invalid_rows: int
    inserted_rows: int
    issues: list[ImportIssue] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)


class AgeingBucketResponse(BaseModel):
    bucket: str
    amount: Decimal
    document_count: int


class PartyExposureResponse(BaseModel):
    party_name: str
    outstanding_amount: Decimal
    overdue_amount: Decimal
    document_count: int
    oldest_due_date: date | None = None
    weighted_days_overdue: Decimal | None = None


class WorkingCapitalSummaryResponse(BaseModel):
    ageing_type: str
    total_outstanding: Decimal
    overdue_amount: Decimal
    overdue_percent: Decimal
    current_amount: Decimal
    document_count: int
    party_count: int
    weighted_average_days_overdue: Decimal | None = None
    buckets: list[AgeingBucketResponse]
    top_parties: list[PartyExposureResponse]


class AnalyticsOverviewResponse(BaseModel):
    monthly_actuals: list[dict]
    branch_comparison: list[dict]
    ar_summary: WorkingCapitalSummaryResponse | None = None
    ap_summary: WorkingCapitalSummaryResponse | None = None
    insights: list[str]
