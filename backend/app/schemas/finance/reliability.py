from datetime import date, datetime
from pydantic import BaseModel


class FinanceReliabilityCheck(BaseModel):
    key: str
    label: str
    status: str
    detail: str
    action: str | None = None
    category: str
    blocking: bool = False


class FinanceReliabilityResponse(BaseModel):
    status: str
    score: int
    pass_count: int
    warning_count: int
    fail_count: int
    checks: list[FinanceReliabilityCheck]
    active_upload_id: str | None = None
    first_transaction_date: date | None = None
    last_transaction_date: date | None = None
    assurance_score: int
    assurance_grade: str
    certified_at: datetime
    caveat: str
