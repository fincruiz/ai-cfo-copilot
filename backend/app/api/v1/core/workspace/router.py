from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company, require_finance_write, require_company_admin
from app.schemas.auth import CurrentUser
from app.schemas.responses import APIResponse
from app.schemas.workspace import (
    DemoDataRequest,
    DemoDataResponse,
    DestructiveActionRequest,
    WorkspaceResetResponse,
    WorkspaceStatusResponse,
    ScopedResetRequest,
    ScopedResetResponse,
    LaunchReadinessResponse,
)
from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService
from app.services.audit_service import AuditService
from app.services.integrations.base import IntegrationStore
from app.services.launch_readiness_service import build_launch_readiness


router = APIRouter(prefix="/workspace", tags=["Workspace & Privacy"])


def get_workspace_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> WorkspaceLifecycleService:
    return WorkspaceLifecycleService(session)


@router.get("/status", response_model=APIResponse[WorkspaceStatusResponse])
async def workspace_status(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[WorkspaceLifecycleService, Depends(get_workspace_service)],
) -> APIResponse[WorkspaceStatusResponse]:
    data = await service.status(company_id=current_company.id)
    return APIResponse(message="Workspace status retrieved.", data=WorkspaceStatusResponse(**data))


@router.get("/launch-readiness", response_model=APIResponse[LaunchReadinessResponse])
async def launch_readiness(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[WorkspaceLifecycleService, Depends(get_workspace_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> APIResponse[LaunchReadinessResponse]:
    workspace = await service.status(company_id=current_company.id)
    connections = await IntegrationStore(session).list_connections(current_company.id)
    data = build_launch_readiness(company=current_company, workspace=workspace, connections=connections)
    return APIResponse(message="Workspace launch readiness retrieved.", data=LaunchReadinessResponse(**data))


@router.post("/demo", response_model=APIResponse[DemoDataResponse])
async def load_demo_data(
    payload: DemoDataRequest,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[WorkspaceLifecycleService, Depends(get_workspace_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> APIResponse[DemoDataResponse]:
    data = await service.seed_demo_data(
        company=current_company,
        user_id=current_user.id,
        replace_existing=payload.replace_existing,
    )
    await AuditService(session).record(
        company_id=current_company.id, user_id=current_user.id, action="seed_demo",
        module="workspace", summary="Loaded synthetic demo finance data",
        metadata={"transactions_created": data.get("transactions_created", 0)}, commit=True,
    )
    return APIResponse(
        message="Demo financial data loaded successfully.",
        data=DemoDataResponse(**data),
    )


@router.delete("/data", response_model=APIResponse[WorkspaceResetResponse])
async def reset_workspace_data(
    payload: DestructiveActionRequest,
    _membership: Annotated[object, Depends(require_company_admin)],
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[WorkspaceLifecycleService, Depends(get_workspace_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> APIResponse[WorkspaceResetResponse]:
    if not payload.confirmed:
        from app.core.exceptions import ApplicationError

        raise ApplicationError(
            message="Please confirm deletion of workspace financial data.",
            error_code="RESET_CONFIRMATION_REQUIRED",
            status_code=422,
        )

    deleted = await service.reset_financial_data(company_id=current_company.id)
    await AuditService(session).record(
        company_id=current_company.id, user_id=current_user.id, action="reset_all",
        module="workspace", summary="Reset all loaded financial data",
        metadata={"deleted_counts": deleted}, commit=True,
    )
    return APIResponse(
        message="All loaded financial data has been deleted. Company profile and account settings were preserved.",
        data=WorkspaceResetResponse(deleted_rows=deleted),
    )


@router.delete("/data/{scope}", response_model=APIResponse[ScopedResetResponse])
async def reset_workspace_scope(
    scope: str,
    payload: ScopedResetRequest,
    _membership: Annotated[object, Depends(require_finance_write)],
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[WorkspaceLifecycleService, Depends(get_workspace_service)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    if not payload.confirmed:
        from app.core.exceptions import ApplicationError
        raise ApplicationError(message="Please confirm this reset.", error_code="RESET_CONFIRMATION_REQUIRED", status_code=422)
    deleted = await service.reset_scope(company_id=current_company.id, scope=scope)
    await AuditService(session).record(company_id=current_company.id, user_id=current_user.id, action="reset", module=scope, summary=f"Reset {scope.replace('_', ' ')} data", metadata={"deleted_counts": deleted}, commit=True)
    return APIResponse(message=f"{scope.replace('_', ' ').title()} data reset successfully.", data=ScopedResetResponse(scope=scope, deleted_rows=deleted))
