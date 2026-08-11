from datetime import UTC, datetime
from pathlib import Path
from uuid import UUID, uuid4

from fastapi import UploadFile
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.models.finance.file_upload import FileUpload
from app.domain.finance.gl_csv_validator import GLCSVValidationResult, validate_gl_csv
from app.domain.finance.ingestion.gl_parser import parse_validated_gl_csv
from app.repositories.finance.file_upload_repository import FileUploadRepository
from app.repositories.core.branch_repository import BranchRepository
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository

MAX_UPLOAD_SIZE_BYTES = 10 * 1024 * 1024
ALLOWED_CSV_CONTENT_TYPES = {"text/csv", "application/csv", "application/vnd.ms-excel", "text/plain", "application/octet-stream"}


class GLUploadService:
    def __init__(self, repository: FileUploadRepository, transaction_repository: GLTransactionRepository, session: AsyncSession) -> None:
        self.repository = repository
        self.transaction_repository = transaction_repository
        self.session = session

    async def validate_and_record_upload(self, *, company: Company, uploaded_by: UUID, file: UploadFile, source_system: str | None = None, reporting_period_id: UUID | None = None) -> tuple[FileUpload, GLCSVValidationResult, int]:
        original_file_name = Path(file.filename or "general-ledger.csv").name
        if Path(original_file_name).suffix.lower() != ".csv":
            raise ApplicationError(message="Only CSV files are supported at this stage.", error_code="UNSUPPORTED_FILE_TYPE", status_code=415, details={"allowed_extensions": [".csv"]})
        if file.content_type and file.content_type not in ALLOWED_CSV_CONTENT_TYPES:
            raise ApplicationError(message="The uploaded file is not a supported CSV file.", error_code="UNSUPPORTED_CONTENT_TYPE", status_code=415, details={"content_type": file.content_type})
        file_bytes = await file.read()
        if not file_bytes:
            raise ApplicationError(message="The uploaded file is empty.", error_code="EMPTY_UPLOAD", status_code=422)
        if len(file_bytes) > MAX_UPLOAD_SIZE_BYTES:
            raise ApplicationError(message="The uploaded file exceeds the 10 MB limit.", error_code="FILE_TOO_LARGE", status_code=413, details={"maximum_size_bytes": MAX_UPLOAD_SIZE_BYTES, "received_size_bytes": len(file_bytes)})
        try:
            validation = validate_gl_csv(file_bytes)
        except ValueError as exc:
            raise ApplicationError(message=str(exc), error_code="INVALID_CSV_FILE", status_code=422) from exc

        upload_id = uuid4()
        safe_file_name = f"{upload_id}_{original_file_name}"
        storage_path = f"{company.id}/general-ledger/{datetime.now(UTC).strftime('%Y/%m/%d')}/{safe_file_name}"
        upload = await self.repository.create({
            "id": upload_id, "company_id": company.id, "reporting_period_id": reporting_period_id,
            "file_name": safe_file_name, "original_file_name": original_file_name,
            "storage_bucket": "company-uploads", "storage_path": storage_path,
            "mime_type": file.content_type or "text/csv", "file_size_bytes": len(file_bytes),
            "document_type": "general_ledger", "source_system": source_system.strip() if source_system else None,
            "processing_status": "validated" if validation.is_valid else "validation_failed",
            "is_active": False,
            "row_count": validation.total_rows, "valid_row_count": validation.valid_rows, "invalid_row_count": validation.invalid_rows,
            "validation_summary": validation.to_dict(), "column_mapping": validation.column_mapping,
            "processing_metadata": {"validation_version": "2.0", "storage_status": "reserved_not_uploaded", "gl_transactions_inserted": False},
            "uploaded_by": uploaded_by, "processed_at": datetime.now(UTC),
        })

        inserted = 0
        if validation.is_valid:
            try:
                rows = parse_validated_gl_csv(
                    file_bytes,
                    default_currency=company.currency_code,
                    company_id=company.id,
                    file_upload_id=upload.id,
                    reporting_period_id=reporting_period_id,
                )
                branch_repository = BranchRepository(self.session)
                branch_mapping = await branch_repository.mapping_by_code_and_name(company.id)

                def branch_code_from_value(value: str) -> str:
                    import re
                    base = re.sub(r"[^A-Z0-9]+", "", value.upper())[:12] or "BRANCH"
                    return base

                for row in rows:
                    branch_reference = row.pop("_branch_reference", None)
                    if not branch_reference:
                        row["branch_id"] = None
                        continue

                    key = branch_reference.strip().lower()
                    branch = branch_mapping.get(key)

                    if branch is None:
                        base_code = branch_code_from_value(branch_reference)
                        code = base_code
                        suffix = 2
                        while await branch_repository.find_by_code_or_name(company.id, code):
                            code = f"{base_code[:9]}{suffix}"
                            suffix += 1

                        branch = await branch_repository.create(
                            {
                                "company_id": company.id,
                                "branch_code": code,
                                "branch_name": branch_reference.strip(),
                                "region": None,
                                "review_status": "pending",
                                "source_value": branch_reference.strip(),
                                "discovered_from_upload_id": upload.id,
                                "is_active": True,
                            }
                        )
                        branch_mapping[key] = branch
                        branch_mapping[code.lower()] = branch

                    row["branch_id"] = branch.id

                inserted = await self.transaction_repository.bulk_create(rows)

                # Preserve every upload, but make only the latest successful dataset
                # active for live reporting. Earlier uploads remain auditable and
                # can be reactivated/versioned later.
                await self.repository.deactivate_active_datasets(
                    company_id=company.id,
                    document_type="general_ledger",
                    reporting_period_id=reporting_period_id,
                    exclude_upload_id=upload.id,
                )

                await self.repository.update(
                    upload,
                    {
                        "is_active": True,
                        "superseded_at": None,
                        "processing_metadata": {
                            **upload.processing_metadata,
                            "gl_transactions_inserted": True,
                            "inserted_transaction_count": inserted,
                            "dataset_status": "active",
                        },
                    },
                )
            except ValueError as exc:
                await self.session.rollback()
                raise ApplicationError(message=str(exc), error_code="GL_INGESTION_FAILED", status_code=422) from exc
        await self.session.commit()
        await self.session.refresh(upload)
        return upload, validation, inserted
