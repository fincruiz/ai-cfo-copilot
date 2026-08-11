from datetime import datetime
from uuid import UUID
from pydantic import BaseModel, ConfigDict, Field


class BranchCreate(BaseModel):
    branch_code: str = Field(min_length=1, max_length=50)
    branch_name: str = Field(min_length=1, max_length=200)
    region: str | None = Field(default=None, max_length=200)


class BranchUpdate(BaseModel):
    branch_code: str | None = Field(default=None, min_length=1, max_length=50)
    branch_name: str | None = Field(default=None, min_length=1, max_length=200)
    region: str | None = Field(default=None, max_length=200)
    review_status: str | None = None
    is_active: bool | None = None


class BranchResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: UUID
    company_id: UUID
    branch_code: str
    branch_name: str
    region: str | None
    review_status: str
    source_value: str | None
    discovered_from_upload_id: UUID | None
    is_active: bool
    created_at: datetime
    updated_at: datetime
