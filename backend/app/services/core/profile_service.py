from collections.abc import Mapping
from typing import Any
from uuid import UUID

from app.database.models.core.profile import Profile
from app.repositories.core.profile_repository import ProfileRepository
from app.services.base import BaseService


class ProfileService(BaseService[Profile]):
    def __init__(self, repository: ProfileRepository) -> None:
        super().__init__(
            repository=repository,
            resource_name="Profile",
        )
        self.profile_repository = repository

    async def get_or_create_profile(self, user_id: UUID) -> Profile:
        profile = await self.profile_repository.get_by_id(user_id)

        if profile is not None:
            return profile

        return await self.profile_repository.create(
            {
                "id": user_id,
            }
        )

    async def update_profile(
        self,
        user_id: UUID,
        values: Mapping[str, Any],
    ) -> Profile:
        profile = await self.get_or_create_profile(user_id)

        return await self.profile_repository.update(
            record=profile,
            values=values,
        )