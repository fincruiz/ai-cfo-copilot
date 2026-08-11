from pathlib import Path
from typing import Annotated
from uuid import UUID, uuid4

from fastapi import APIRouter, Depends, File, Query, UploadFile, status
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company
from app.repositories.core.company_member_repository import (
    CompanyMemberRepository,
)
from app.repositories.core.company_repository import (
    CompanyRepository,
)
from app.schemas.auth import CurrentUser
from app.schemas.core.company import (
    CompanyResponse,
    CreateCompanyRequest,
    UpdateCompanyRequest,
    CompanyPreferencesRequest,
    CompanyPreferencesResponse,
)
from app.schemas.responses import (
    APIResponse,
    PaginatedResponse,
)
from app.services.core.company_member_service import (
    CompanyMemberService,
)
from app.services.core.company_service import CompanyService
from app.services.core.onboarding_service import (
    OnboardingService,
)


router = APIRouter(
    prefix="/companies",
    tags=["Companies"],
)


def get_company_service(
    session: Annotated[
        AsyncSession,
        Depends(get_db_session),
    ],
) -> CompanyService:
    repository = CompanyRepository(session)

    return CompanyService(repository)


def get_onboarding_service(
    session: Annotated[
        AsyncSession,
        Depends(get_db_session),
    ],
) -> OnboardingService:
    company_repository = CompanyRepository(session)
    company_member_repository = CompanyMemberRepository(session)

    company_service = CompanyService(
        company_repository
    )

    company_member_service = CompanyMemberService(
        company_member_repository
    )

    return OnboardingService(
        session=session,
        company_service=company_service,
        company_member_service=company_member_service,
    )


@router.post(
    "",
    response_model=APIResponse[CompanyResponse],
    status_code=status.HTTP_201_CREATED,
)
async def create_company(
    request: CreateCompanyRequest,
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
    service: Annotated[
        OnboardingService,
        Depends(get_onboarding_service),
    ],
) -> APIResponse[CompanyResponse]:
    company = await service.onboard_company(
        request=request,
        user_id=current_user.id,
    )

    return APIResponse[CompanyResponse](
        message="Company created successfully.",
        data=CompanyResponse.model_validate(company),
    )


@router.get(
    "/me",
    response_model=APIResponse[CompanyResponse],
)
async def get_my_company(
    current_company: Annotated[
        Company,
        Depends(get_current_company),
    ],
) -> APIResponse[CompanyResponse]:
    return APIResponse[CompanyResponse](
        message="Current company retrieved successfully.",
        data=CompanyResponse.model_validate(
            current_company
        ),
    )


@router.get(
    "",
    response_model=PaginatedResponse[CompanyResponse],
)
async def list_companies(
    service: Annotated[
        CompanyService,
        Depends(get_company_service),
    ],
    active_only: bool = True,
    limit: Annotated[
        int,
        Query(ge=1, le=200),
    ] = 100,
    offset: Annotated[
        int,
        Query(ge=0),
    ] = 0,
) -> PaginatedResponse[CompanyResponse]:
    companies, count = await service.list_companies(
        active_only=active_only,
        limit=limit,
        offset=offset,
    )

    return PaginatedResponse[CompanyResponse](
        message="Companies retrieved successfully.",
        count=count,
        limit=limit,
        offset=offset,
        data=[
            CompanyResponse.model_validate(company)
            for company in companies
        ],
    )


@router.get(
    "/{company_id}",
    response_model=APIResponse[CompanyResponse],
)
async def get_company(
    company_id: UUID,
    service: Annotated[
        CompanyService,
        Depends(get_company_service),
    ],
) -> APIResponse[CompanyResponse]:
    company = await service.get_company(
        company_id
    )

    return APIResponse[CompanyResponse](
        message="Company retrieved successfully.",
        data=CompanyResponse.model_validate(company),
    )

ALLOWED_LOGO_TYPES = {
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/webp": ".webp",
}
MAX_LOGO_BYTES = 2 * 1024 * 1024


@router.post(
    "/me/logo",
    response_model=APIResponse[CompanyResponse],
)
async def upload_company_logo(
    file: Annotated[UploadFile, File(...)],
    current_company: Annotated[
        Company,
        Depends(get_current_company),
    ],
    service: Annotated[
        CompanyService,
        Depends(get_company_service),
    ],
) -> APIResponse[CompanyResponse]:
    extension = ALLOWED_LOGO_TYPES.get(
        file.content_type or ""
    )

    if extension is None:
        from app.core.exceptions import ApplicationError

        raise ApplicationError(
            message="Logo must be PNG, JPG or WebP.",
            error_code="INVALID_LOGO_TYPE",
            status_code=415,
        )

    content = await file.read()

    if not content:
        from app.core.exceptions import ApplicationError

        raise ApplicationError(
            message="The logo file is empty.",
            error_code="EMPTY_LOGO",
            status_code=422,
        )

    if len(content) > MAX_LOGO_BYTES:
        from app.core.exceptions import ApplicationError

        raise ApplicationError(
            message="Logo must be smaller than 2 MB.",
            error_code="LOGO_TOO_LARGE",
            status_code=413,
        )

    logo_directory = Path("uploads/logos")
    logo_directory.mkdir(parents=True, exist_ok=True)

    file_name = f"{current_company.id}_{uuid4().hex}{extension}"
    file_path = logo_directory / file_name
    file_path.write_bytes(content)

    company = await service.update_logo(
        current_company,
        f"/uploads/logos/{file_name}",
    )

    return APIResponse[CompanyResponse](
        message="Company logo uploaded successfully.",
        data=CompanyResponse.model_validate(company),
    )


@router.put("/me", response_model=APIResponse[CompanyResponse])
async def update_my_company(
    request: UpdateCompanyRequest,
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[CompanyService, Depends(get_company_service)],
):
    company = await service.update_company(
        current_company,
        request.model_dump(exclude_unset=True),
    )
    return APIResponse(message="Company profile updated.", data=CompanyResponse.model_validate(company))


@router.get("/me/preferences", response_model=APIResponse[CompanyPreferencesResponse])
async def get_preferences(
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    from sqlalchemy import text
    row = (await session.execute(text("""
        SELECT company_id, theme_preference, number_format, reporting_frequency,
               default_report_view, show_ai_assistant, email_notifications,
               variance_warning_percent
        FROM public.company_preferences
        WHERE company_id=:company_id
    """), {"company_id": current_company.id})).mappings().first()
    if not row:
        await session.execute(text("INSERT INTO public.company_preferences (company_id) VALUES (:company_id) ON CONFLICT DO NOTHING"), {"company_id": current_company.id})
        await session.commit()
        row = {
            "company_id": current_company.id,
            "theme_preference": "system",
            "number_format": "international",
            "reporting_frequency": "monthly",
            "default_report_view": "consolidated",
            "show_ai_assistant": True,
            "email_notifications": True,
            "variance_warning_percent": 10,
        }
    return APIResponse(message="Preferences retrieved.", data=CompanyPreferencesResponse(**row))


@router.put("/me/preferences", response_model=APIResponse[CompanyPreferencesResponse])
async def update_preferences(
    request: CompanyPreferencesRequest,
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    from sqlalchemy import text
    values = request.model_dump()
    await session.execute(text("""
        INSERT INTO public.company_preferences
        (company_id, theme_preference, number_format, reporting_frequency,
         default_report_view, show_ai_assistant, email_notifications,
         variance_warning_percent, updated_at)
        VALUES
        (:company_id, :theme_preference, :number_format, :reporting_frequency,
         :default_report_view, :show_ai_assistant, :email_notifications,
         :variance_warning_percent, now())
        ON CONFLICT (company_id) DO UPDATE SET
          theme_preference=EXCLUDED.theme_preference,
          number_format=EXCLUDED.number_format,
          reporting_frequency=EXCLUDED.reporting_frequency,
          default_report_view=EXCLUDED.default_report_view,
          show_ai_assistant=EXCLUDED.show_ai_assistant,
          email_notifications=EXCLUDED.email_notifications,
          variance_warning_percent=EXCLUDED.variance_warning_percent,
          updated_at=now()
    """), {"company_id": current_company.id, **values})
    await session.commit()
    return APIResponse(message="Preferences updated.", data=CompanyPreferencesResponse(company_id=current_company.id, **values))
