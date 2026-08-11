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