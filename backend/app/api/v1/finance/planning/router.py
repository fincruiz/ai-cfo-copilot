from typing import Annotated
from uuid import UUID
from fastapi import APIRouter, Depends, File, Form, UploadFile
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.finance.planning import PlanImportResponse, VarianceLineResponse
from app.schemas.responses import APIResponse
from app.services.finance.planning_service import PlanningService

router = APIRouter(prefix="/planning", tags=["Finance Planning"])

def service(session: Annotated[AsyncSession, Depends(get_db_session)]):
    return PlanningService(session)

async def upload_plan(kind, file, version_name, replace_existing, company, svc):
    if not (file.filename or "").lower().endswith(".csv"):
        raise ValueError("Budget and forecast imports currently accept CSV files.")
    content = await file.read()
    result = await svc.import_plan(
        company_id=company.id,
        plan_type=kind,
        version_name=version_name,
        content=content,
        replace_existing=replace_existing,
    )
    return APIResponse(message=f"{kind.title()} imported.", data=PlanImportResponse(**result))

@router.post("/budget", response_model=APIResponse[PlanImportResponse])
async def budget(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[PlanningService, Depends(service)],
    version_name: Annotated[str, Form()] = "Default",
    replace_existing: Annotated[bool, Form()] = True,
):
    return await upload_plan("budget", file, version_name, replace_existing, current_company, svc)

@router.post("/forecast", response_model=APIResponse[PlanImportResponse])
async def forecast(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[PlanningService, Depends(service)],
    version_name: Annotated[str, Form()] = "Default",
    replace_existing: Annotated[bool, Form()] = True,
):
    return await upload_plan("forecast", file, version_name, replace_existing, current_company, svc)

@router.get("/variance", response_model=APIResponse[list[VarianceLineResponse]])
async def variance(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[PlanningService, Depends(service)],
    branch_id: UUID | None = None,
):
    return APIResponse(
        message="Actual versus budget and forecast retrieved.",
        data=[VarianceLineResponse(**row) for row in await svc.variance(current_company.id, branch_id)],
    )
