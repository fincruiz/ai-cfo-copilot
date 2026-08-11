from uuid import UUID

from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.schemas.core.company import CreateCompanyRequest
from app.services.core.company_member_service import (
    CompanyMemberService,
)
from app.services.core.company_service import CompanyService


class OnboardingService:
    def __init__(
        self,
        *,
        session: AsyncSession,
        company_service: CompanyService,
        company_member_service: CompanyMemberService,
    ) -> None:
        self.session = session
        self.company_service = company_service
        self.company_member_service = company_member_service

    async def onboard_company(
        self,
        *,
        request: CreateCompanyRequest,
        user_id: UUID,
    ) -> Company:
        existing_membership = (
            await self.company_member_service
            .get_active_membership_by_user(
                user_id=user_id,
            )
        )

        if existing_membership is not None:
            raise ApplicationError(
                message=(
                    "Company onboarding has already been completed "
                    "for this user."
                ),
                error_code="COMPANY_ALREADY_EXISTS",
                status_code=409,
            )

        try:
            company = await self.company_service.create_company(
                request=request,
                created_by=user_id,
            )

            # The database trigger on public.companies automatically
            # creates the owner's company_members record.
            # Do not create another membership here.

            await self.session.commit()
            await self.session.refresh(company)

            return company

        except Exception:
            await self.session.rollback()
            raise