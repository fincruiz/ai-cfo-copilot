from uuid import UUID

from app.database.models.core.company_member import CompanyMember
from app.repositories.core.company_member_repository import (
    CompanyMemberRepository,
)
from app.services.base import BaseService


class CompanyMemberService(BaseService[CompanyMember]):
    def __init__(
        self,
        repository: CompanyMemberRepository,
    ) -> None:
        super().__init__(
            repository=repository,
            resource_name="Company member",
        )

        self.company_member_repository = repository

    async def get_active_membership_by_user(
        self,
        *,
        user_id: UUID,
    ) -> CompanyMember | None:
        return (
            await self.company_member_repository
            .get_active_membership_by_user(user_id)
        )

    async def get_active_owner_membership(
        self,
        *,
        user_id: UUID,
    ) -> CompanyMember | None:
        return (
            await self.company_member_repository
            .get_active_owner_membership(user_id)
        )

    async def create_owner_membership(
        self,
        *,
        company_id: UUID,
        user_id: UUID,
    ) -> CompanyMember:
        membership = await self.company_member_repository.create(
            {
                "company_id": company_id,
                "user_id": user_id,
                "role": "owner",
                "is_active": True,
                "invited_by": None,
            }
        )

        return membership