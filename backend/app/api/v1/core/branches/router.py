from typing import Annotated
from uuid import UUID

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company, require_finance_write
from app.dependencies.auth import get_current_user
from app.schemas.auth import CurrentUser
from app.services.audit_service import AuditService
from app.repositories.core.branch_repository import BranchRepository
from app.schemas.core.branch import BranchCreate, BranchResponse, BranchUpdate
from app.schemas.responses import APIResponse
from app.services.core.branch_service import BranchService

router = APIRouter(prefix="/branches", tags=["Branches"])


def get_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> BranchService:
    return BranchService(BranchRepository(session))


@router.get("", response_model=APIResponse[list[BranchResponse]])
async def list_branches(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[BranchService, Depends(get_service)],
):
    branches = await service.list_branches(current_company.id)
    return APIResponse(
        message="Branches retrieved.",
        data=[BranchResponse.model_validate(item) for item in branches],
    )


@router.post("", response_model=APIResponse[BranchResponse], status_code=201)
async def create_branch(
    payload: BranchCreate,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    _membership: Annotated[object, Depends(require_finance_write)],
    service: Annotated[BranchService, Depends(get_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    branch = await service.create_branch(current_company.id, payload)
    await AuditService(session).record(company_id=current_company.id, user_id=current_user.id, action="create", module="branches", summary=f"Created branch {branch.branch_code}", commit=True)
    return APIResponse(
        message="Branch created.",
        data=BranchResponse.model_validate(branch),
    )


@router.put("/{branch_id}", response_model=APIResponse[BranchResponse])
async def update_branch(
    branch_id: UUID,
    payload: BranchUpdate,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    _membership: Annotated[object, Depends(require_finance_write)],
    service: Annotated[BranchService, Depends(get_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    branch = await service.update_branch(
        current_company.id,
        branch_id,
        payload,
    )
    await AuditService(session).record(company_id=current_company.id, user_id=current_user.id, action="update", module="branches", summary=f"Updated branch {branch.branch_code}", commit=True)
    return APIResponse(
        message="Branch updated.",
        data=BranchResponse.model_validate(branch),
    )
