from uuid import UUID

from app.core.exceptions import ApplicationError
from app.database.models.core.branch import Branch
from app.repositories.core.branch_repository import BranchRepository
from app.schemas.core.branch import BranchCreate, BranchUpdate


class BranchService:
    def __init__(self, repository: BranchRepository) -> None:
        self.repository = repository
        self.session = repository.session

    async def list_branches(self, company_id: UUID) -> list[Branch]:
        return await self.repository.list_company_branches(company_id)

    async def create_branch(
        self,
        company_id: UUID,
        payload: BranchCreate,
    ) -> Branch:
        existing = await self.repository.find_by_code_or_name(
            company_id,
            payload.branch_code,
        )
        if existing:
            raise ApplicationError(
                message="A branch with this code or name already exists.",
                error_code="BRANCH_ALREADY_EXISTS",
                status_code=409,
            )
        branch = await self.repository.create(
            {
                "company_id": company_id,
                "branch_code": payload.branch_code.strip(),
                "branch_name": payload.branch_name.strip(),
                "region": payload.region.strip() if payload.region else None,
            }
        )
        await self.session.commit()
        await self.session.refresh(branch)
        return branch

    async def update_branch(
        self,
        company_id: UUID,
        branch_id: UUID,
        payload: BranchUpdate,
    ) -> Branch:
        branch = await self.repository.get_company_branch(company_id, branch_id)
        if branch is None:
            raise ApplicationError(
                message="Branch not found.",
                error_code="BRANCH_NOT_FOUND",
                status_code=404,
            )
        values = payload.model_dump(exclude_unset=True)
        if "branch_code" in values and values["branch_code"]:
            values["branch_code"] = values["branch_code"].strip().upper()
        if "branch_name" in values and values["branch_name"]:
            values["branch_name"] = values["branch_name"].strip()
        if values.get("review_status") not in {None, "pending", "accepted", "rejected"}:
            raise ApplicationError(message="Invalid branch review status.", error_code="INVALID_BRANCH_STATUS", status_code=422)
        if "region" in values and values["region"]:
            values["region"] = values["region"].strip()
        branch = await self.repository.update(branch, values)
        await self.session.commit()
        await self.session.refresh(branch)
        return branch
