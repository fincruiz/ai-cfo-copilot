from typing import Annotated
from uuid import UUID

from fastapi import (
    APIRouter,
    Depends,
    File,
    Form,
    UploadFile,
    status,
)
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company
from app.repositories.finance.file_upload_repository import (
    FileUploadRepository,
)
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.schemas.auth import CurrentUser
from app.schemas.finance.uploads import (
    FileUploadResponse,
    GLUploadValidationResponse,
    GLValidationSummary,
    ValidationIssue,
)
from app.schemas.responses import APIResponse
from app.services.audit_service import AuditService
from app.services.finance.gl_upload_service import (
    GLUploadService,
)


router = APIRouter(
    prefix="/uploads",
    tags=["Finance Uploads"],
)


def get_gl_upload_service(
    session: Annotated[
        AsyncSession,
        Depends(get_db_session),
    ],
) -> GLUploadService:
    repository = FileUploadRepository(session)
    transaction_repository = GLTransactionRepository(session)

    return GLUploadService(repository, transaction_repository, session)


@router.post(
    "/general-ledger",
    response_model=APIResponse[GLUploadValidationResponse],
    status_code=status.HTTP_201_CREATED,
)
async def upload_general_ledger(
    file: Annotated[
        UploadFile,
        File(description="General ledger CSV file."),
    ],
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
    current_company: Annotated[
        Company,
        Depends(get_current_company),
    ],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    service: Annotated[
        GLUploadService,
        Depends(get_gl_upload_service),
    ],
    source_system: Annotated[
        str | None,
        Form(),
    ] = None,
    reporting_period_id: Annotated[
        UUID | None,
        Form(),
    ] = None,
) -> APIResponse[GLUploadValidationResponse]:
    upload, validation, inserted_count = (
        await service.validate_and_record_upload(
            company=current_company,
            uploaded_by=current_user.id,
            file=file,
            source_system=source_system,
            reporting_period_id=reporting_period_id,
        )
    )

    await AuditService(session).record(company_id=current_company.id, user_id=current_user.id, action="upload", module="general_ledger", summary=f"Uploaded general ledger: {file.filename or 'general_ledger.csv'}", metadata={"inserted_transactions": inserted_count, "invalid_rows": validation.invalid_rows}, commit=True)

    validation_response = GLValidationSummary(
        required_columns=validation.required_columns,
        detected_columns=validation.detected_columns,
        missing_columns=validation.missing_columns,
        total_rows=validation.total_rows,
        valid_rows=validation.valid_rows,
        invalid_rows=validation.invalid_rows,
        issues=[
            ValidationIssue(
                row_number=issue.row_number,
                column=issue.column,
                message=issue.message,
                severity=issue.severity,
            )
            for issue in validation.issues
        ],
    )

    return APIResponse[GLUploadValidationResponse](
        message=(
            "General ledger validated successfully."
            if validation.is_valid
            else "General ledger uploaded with validation issues."
        ),
        data=GLUploadValidationResponse(
            upload=FileUploadResponse.model_validate(upload),
            validation=validation_response,
            inserted_transaction_count=inserted_count,
        ),
    )