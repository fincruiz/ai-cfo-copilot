from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.schemas.auth import CurrentUser
from app.schemas.responses import APIResponse
from app.schemas.workspace import AccountDeletionResponse, DestructiveActionRequest
from app.services.account_deletion_service import AccountDeletionService


router = APIRouter(prefix="/account", tags=["Account & Privacy"])


def get_account_deletion_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> AccountDeletionService:
    return AccountDeletionService(session)


@router.delete("/me", response_model=APIResponse[AccountDeletionResponse])
async def delete_my_account(
    payload: DestructiveActionRequest,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    service: Annotated[AccountDeletionService, Depends(get_account_deletion_service)],
) -> APIResponse[AccountDeletionResponse]:
    if not payload.confirmed:
        from app.core.exceptions import ApplicationError

        raise ApplicationError(
            message="Please confirm permanent profile deletion.",
            error_code="DELETE_CONFIRMATION_REQUIRED",
            status_code=422,
        )

    data = await service.delete_account(user_id=current_user.id)
    return APIResponse(
        message="Your FinCruiz profile and associated single-user workspace data have been permanently deleted.",
        data=AccountDeletionResponse(**data),
    )
