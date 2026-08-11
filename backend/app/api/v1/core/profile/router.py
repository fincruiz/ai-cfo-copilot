from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.repositories.core.profile_repository import ProfileRepository
from app.schemas.auth import CurrentUser
from app.schemas.core.profile import ProfileResponse, UpdateProfileRequest
from app.schemas.responses import APIResponse
from app.services.core.profile_service import ProfileService


router = APIRouter(
    prefix="/profile",
    tags=["Profile"],
)


def get_profile_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> ProfileService:
    repository = ProfileRepository(session=session)
    return ProfileService(repository=repository)


@router.get(
    "/me",
    response_model=APIResponse[ProfileResponse],
)
async def get_my_profile(
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    service: Annotated[ProfileService, Depends(get_profile_service)],
) -> APIResponse[ProfileResponse]:
    profile = await service.get_or_create_profile(current_user.id)

    return APIResponse(
        success=True,
        message="Profile retrieved successfully.",
        data=ProfileResponse.model_validate(profile),
    )


@router.put(
    "/me",
    response_model=APIResponse[ProfileResponse],
)
async def update_my_profile(
    payload: UpdateProfileRequest,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    service: Annotated[ProfileService, Depends(get_profile_service)],
) -> APIResponse[ProfileResponse]:
    values = payload.model_dump(exclude_unset=True)

    profile = await service.update_profile(
        user_id=current_user.id,
        values=values,
    )

    return APIResponse(
        success=True,
        message="Profile updated successfully.",
        data=ProfileResponse.model_validate(profile),
    )