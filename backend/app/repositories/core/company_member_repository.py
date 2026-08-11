from uuid import UUID

from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company_member import CompanyMember
from app.repositories.base import BaseRepository


class CompanyMemberRepository(BaseRepository[CompanyMember]):
    def __init__(self, session: AsyncSession) -> None:
        super().__init__(
            session=session,
            model=CompanyMember,
        )

    async def get_active_membership_by_user(
        self,
        user_id: UUID,
    ) -> CompanyMember | None:
        statement = (
            select(CompanyMember)
            .where(
                CompanyMember.user_id == user_id,
                CompanyMember.is_active.is_(True),
            )
            .order_by(CompanyMember.created_at.asc())
            .limit(1)
        )

        result = await self.session.execute(statement)
        return result.scalar_one_or_none()

    async def get_active_owner_membership(
        self,
        user_id: UUID,
    ) -> CompanyMember | None:
        statement = (
            select(CompanyMember)
            .where(
                CompanyMember.user_id == user_id,
                CompanyMember.role == "owner",
                CompanyMember.is_active.is_(True),
            )
            .limit(1)
        )

        result = await self.session.execute(statement)
        return result.scalar_one_or_none()