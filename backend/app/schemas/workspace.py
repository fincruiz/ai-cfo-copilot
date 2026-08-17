from typing import Any

from pydantic import BaseModel


class DestructiveActionRequest(BaseModel):
    confirmed: bool = False


class DemoDataRequest(BaseModel):
    replace_existing: bool = False


class WorkspaceStatusResponse(BaseModel):
    has_financial_data: bool
    demo_data_active: bool
    upload_count: int
    transaction_count: int
    mapping_count: int


class WorkspaceResetResponse(BaseModel):
    deleted_rows: dict[str, int]


class DemoDataResponse(BaseModel):
    upload_id: str
    months: int
    transactions_created: int
    mappings_created: int


class AccountDeletionResponse(BaseModel):
    auth_user_deleted: bool
    companies_deleted: int
    memberships_deleted: int
    profile_deleted: bool


class ScopedResetRequest(BaseModel):
    confirmed: bool = False

class ScopedResetResponse(BaseModel):
    scope: str
    deleted_rows: dict[str, int]


class LaunchReadinessCheck(BaseModel):
    key: str
    label: str
    ready: bool
    detail: str
    path: str


class LaunchReadinessResponse(BaseModel):
    score: int
    completed_steps: int
    total_steps: int
    checks: list[LaunchReadinessCheck]
    next_path: str
    next_label: str
    connected_sources: int
    healthy_sources: int
    ready_for_management_use: bool


class CommercialOnboardingStep(BaseModel):
    key: str
    label: str
    complete: bool


class CommercialOnboardingSummaryResponse(BaseModel):
    stage: str
    ready_for_intelligence: bool
    progress_percent: int
    completed_steps: int
    total_steps: int
    steps: list[CommercialOnboardingStep]
    transaction_count: int
    account_count: int
    mapping_count: int
    unmapped_account_count: int
    branch_count: int
    pending_branch_count: int
    months_history: int
    period_start: str | None = None
    period_end: str | None = None
    financial_confidence_score: float | None = None
    financial_confidence_grade: str | None = None
    financial_checks: list[dict[str, Any]] = []
    latest_ingestion: dict[str, Any] | None = None
    briefing: dict[str, Any] | None = None
    briefing_error: str | None = None
    next_path: str
    next_label: str
