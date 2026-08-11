from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.repositories.base import BaseRepository


class CompanyRepository(BaseRepository[Company]):
    def __init__(
        self,
        session: AsyncSession,
    ) -> None:
        super().__init__(
            session=session,
            model=Company,
        )

    async def list_companies(
        self,
        *,
        active_only: bool = True,
        limit: int = 100,
        offset: int = 0,
    ) -> tuple[list[Company], int]:
        filters = {}

        if active_only:
            filters["is_active"] = True

        companies = await self.list_records(
            limit=limit,
            offset=offset,
            order_by=Company.legal_name.asc(),
            filters=filters,
        )

        count = await self.count_records(
            filters=filters,
        )

        return companies, count