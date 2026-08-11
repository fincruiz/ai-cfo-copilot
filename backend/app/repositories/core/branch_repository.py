from uuid import UUID
from sqlalchemy import func, or_, select
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.branch import Branch
from app.repositories.base import BaseRepository


class BranchRepository(BaseRepository[Branch]):
    def __init__(self, session: AsyncSession) -> None:
        super().__init__(session=session, model=Branch)

    async def list_company_branches(self, company_id: UUID, *, active_only: bool = False) -> list[Branch]:
        statement = select(Branch).where(Branch.company_id == company_id).order_by(Branch.review_status.desc(), Branch.branch_code.asc())
        if active_only:
            statement = statement.where(Branch.is_active.is_(True))
        return list((await self.session.execute(statement)).scalars().all())

    async def get_company_branch(self, company_id: UUID, branch_id: UUID) -> Branch | None:
        return (await self.session.execute(select(Branch).where(Branch.id == branch_id, Branch.company_id == company_id))).scalar_one_or_none()

    async def find_by_code_or_name(self, company_id: UUID, value: str) -> Branch | None:
        clean = value.strip()
        return (await self.session.execute(
            select(Branch).where(
                Branch.company_id == company_id,
                Branch.is_active.is_(True),
                or_(func.lower(Branch.branch_code) == clean.lower(), func.lower(Branch.branch_name) == clean.lower(), func.lower(Branch.source_value) == clean.lower()),
            )
        )).scalar_one_or_none()

    async def mapping_by_code_and_name(self, company_id: UUID) -> dict[str, Branch]:
        mapping: dict[str, Branch] = {}
        for branch in await self.list_company_branches(company_id, active_only=True):
            for value in (branch.branch_code, branch.branch_name, branch.source_value):
                if value:
                    mapping[value.strip().lower()] = branch
        return mapping
