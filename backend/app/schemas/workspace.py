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
