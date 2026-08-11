from typing import Annotated

from fastapi import APIRouter, Depends, File, Form, UploadFile
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.finance.imports import FinanceImportResponse
from app.schemas.responses import APIResponse
from app.services.finance.import_service import FinanceImportService

router = APIRouter(prefix="/imports", tags=["Finance Imports"])


def get_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> FinanceImportService:
    return FinanceImportService(session)


async def _read_csv(file: UploadFile) -> bytes:
    if not (file.filename or "").lower().endswith(".csv"):
        raise ValueError("This import currently accepts CSV files.")
    content = await file.read()
    if not content:
        raise ValueError("The uploaded file is empty.")
    if len(content) > 15 * 1024 * 1024:
        raise ValueError("The uploaded file exceeds the 15 MB limit.")
    return content


@router.post("/coa", response_model=APIResponse[FinanceImportResponse])
async def upload_coa(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[FinanceImportService, Depends(get_service)],
    source_system: Annotated[str | None, Form()] = None,
):
    result = await service.import_coa(
        company_id=current_company.id,
        file_name=file.filename or "coa.csv",
        content=await _read_csv(file),
        source_system=source_system,
    )
    return APIResponse(
        message="Chart of accounts imported successfully.",
        data=result,
    )


@router.post("/ar-ageing", response_model=APIResponse[FinanceImportResponse])
async def upload_ar_ageing(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[FinanceImportService, Depends(get_service)],
    source_system: Annotated[str | None, Form()] = None,
    replace_existing: Annotated[bool, Form()] = True,
):
    result = await service.import_ageing(
        company_id=current_company.id,
        ageing_type="AR",
        file_name=file.filename or "ar_ageing.csv",
        content=await _read_csv(file),
        source_system=source_system,
        replace_existing=replace_existing,
    )
    return APIResponse(
        message="Accounts receivable ageing imported successfully.",
        data=result,
    )


@router.post("/ap-ageing", response_model=APIResponse[FinanceImportResponse])
async def upload_ap_ageing(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[FinanceImportService, Depends(get_service)],
    source_system: Annotated[str | None, Form()] = None,
    replace_existing: Annotated[bool, Form()] = True,
):
    result = await service.import_ageing(
        company_id=current_company.id,
        ageing_type="AP",
        file_name=file.filename or "ap_ageing.csv",
        content=await _read_csv(file),
        source_system=source_system,
        replace_existing=replace_existing,
    )
    return APIResponse(
        message="Accounts payable ageing imported successfully.",
        data=result,
    )
