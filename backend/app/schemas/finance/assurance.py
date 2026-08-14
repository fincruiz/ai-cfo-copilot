from pydantic import BaseModel


class AssuranceCheck(BaseModel):
    key: str
    label: str
    status: str
    score: int
    detail: str
    action: str | None = None


class FinancialAssuranceResponse(BaseModel):
    score: int
    grade: str
    status: str
    checks: list[AssuranceCheck]
    caveat: str
