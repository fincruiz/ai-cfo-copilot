from uuid import UUID

from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.repositories.base import BaseRepository


class CompanyRepository(BaseRepository[Company]):
    def __init__(self, session: AsyncSession) -> None:
        super().__init__(session=session, model=Company)

    async def list_companies(
        self,
        *,
        active_only: bool = True,
        limit: int = 100,
        offset: int = 0,
    ) -> tuple[list[Company], int]:
        filters = {"is_active": True} if active_only else {}
        companies = await self.list_records(
            limit=limit,
            offset=offset,
            order_by=Company.legal_name.asc(),
            filters=filters,
        )
        count = await self.count_records(filters=filters)
        return companies, count

    async def list_companies_by_ids(
        self,
        *,
        company_ids: list[UUID],
        active_only: bool = True,
        limit: int = 100,
        offset: int = 0,
    ) -> tuple[list[Company], int]:
        if not company_ids:
            return [], 0

        base = select(Company).where(Company.id.in_(company_ids))
        count_statement = (
            select(func.count())
            .select_from(Company)
            .where(Company.id.in_(company_ids))
        )

        if active_only:
            base = base.where(Company.is_active.is_(True))
            count_statement = count_statement.where(Company.is_active.is_(True))

        result = await self.session.execute(
            base.order_by(Company.legal_name.asc()).limit(limit).offset(offset)
        )
        count_result = await self.session.execute(count_statement)
        return list(result.scalars().all()), int(count_result.scalar_one())
