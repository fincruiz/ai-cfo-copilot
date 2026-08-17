from datetime import datetime
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field


class ValidationIssue(BaseModel):
    row_number: int | None = None
    column: str | None = None
    message: str
    severity: str = "error"


class GLValidationSummary(BaseModel):
    required_columns: list[str]
    detected_columns: list[str]
    missing_columns: list[str]
    total_rows: int = Field(ge=0)
    valid_rows: int = Field(ge=0)
    invalid_rows: int = Field(ge=0)
    issues: list[ValidationIssue]


class FileUploadResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    company_id: UUID
    reporting_period_id: UUID | None = None
    file_name: str
    original_file_name: str | None = None
    storage_bucket: str
    storage_path: str
    mime_type: str | None = None
    file_size_bytes: int | None = None
    document_type: str
    source_system: str | None = None
    processing_status: str
    row_count: int | None = None
    valid_row_count: int | None = None
    invalid_row_count: int | None = None
    validation_summary: dict
    column_mapping: dict
    processing_metadata: dict
    uploaded_by: UUID | None = None
    processed_at: datetime | None = None
    created_at: datetime
    updated_at: datetime


class GLUploadValidationResponse(BaseModel):
    upload: FileUploadResponse
    validation: GLValidationSummary

class IngestionJobResponse(BaseModel):
    id: UUID
    company_id: UUID
    job_type: str
    original_file_name: str
    file_size_bytes: int
    source_system: str | None = None
    status: str
    progress_percent: int = Field(ge=0, le=100)
    phase: str
    total_rows: int | None = None
    valid_rows: int | None = None
    invalid_rows: int | None = None
    inserted_rows: int = 0
    file_upload_id: UUID | None = None
    error_message: str | None = None
    attempts: int = 0
    created_at: datetime
    started_at: datetime | None = None
    completed_at: datetime | None = None
    updated_at: datetime
