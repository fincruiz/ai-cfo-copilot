from typing import Annotated

from fastapi import Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.repositories.core.company_member_repository import (
    CompanyMemberRepository,
)
from app.repositories.core.company_repository import (
    CompanyRepository,
)
from app.schemas.auth import CurrentUser
from app.services.core.company_context_service import (
    CompanyContextService,
)
from app.services.core.company_member_service import (
    CompanyMemberService,
)
from app.services.core.company_service import CompanyService


async def get_current_company(
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
    session: Annotated[
        AsyncSession,
        Depends(get_db_session),
    ],
) -> Company:
    company_repository = CompanyRepository(session)
    company_member_repository = CompanyMemberRepository(session)

    company_service = CompanyService(
        company_repository
    )

    company_member_service = CompanyMemberService(
        company_member_repository
    )

    company_context_service = CompanyContextService(
        company_service=company_service,
        company_member_service=company_member_service,
    )

    return await company_context_service.get_current_company(
        user_id=current_user.id,
    )

async def get_current_company_membership(
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    """Return the active membership behind the current workspace."""
    from app.core.exceptions import ApplicationError
    from app.repositories.core.company_member_repository import CompanyMemberRepository

    membership = await CompanyMemberRepository(session).get_active_membership_by_user(current_user.id)
    if membership is None:
        raise ApplicationError(
            message="No active company membership was found for this user.",
            error_code="COMPANY_MEMBERSHIP_NOT_FOUND",
            status_code=404,
        )
    return membership


def require_company_roles(*allowed_roles: str):
    async def dependency(
        membership=Depends(get_current_company_membership),
    ):
        from app.core.exceptions import ApplicationError

        role = membership.role.value if hasattr(membership.role, "value") else str(membership.role)
        if role not in allowed_roles:
            raise ApplicationError(
                message="Your company role does not allow this action.",
                error_code="INSUFFICIENT_COMPANY_ROLE",
                status_code=403,
            )
        return membership

    return dependency


require_finance_write = require_company_roles(
    "owner", "admin", "cfo", "finance_manager", "accountant"
)
require_company_admin = require_company_roles("owner", "admin")
