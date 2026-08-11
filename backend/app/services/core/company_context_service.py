from uuid import UUID

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.services.core.company_member_service import (
    CompanyMemberService,
)
from app.services.core.company_service import CompanyService


class CompanyContextService:
    def __init__(
        self,
        *,
        company_service: CompanyService,
        company_member_service: CompanyMemberService,
    ) -> None:
        self.company_service = company_service
        self.company_member_service = company_member_service

    async def get_current_company(
        self,
        *,
        user_id: UUID,
    ) -> Company:
        membership = (
            await self.company_member_service
            .get_active_membership_by_user(
                user_id=user_id,
            )
        )

        if membership is None:
            raise ApplicationError(
                message=(
                    "No active company membership was found "
                    "for this user."
                ),
                error_code="COMPANY_MEMBERSHIP_NOT_FOUND",
                status_code=404,
            )

        company = await self.company_service.get_company(
            membership.company_id
        )

        if not company.is_active:
            raise ApplicationError(
                message="The current company is inactive.",
                error_code="COMPANY_INACTIVE",
                status_code=403,
            )

        return company