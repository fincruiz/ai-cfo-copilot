from pathlib import Path
from typing import Annotated
from uuid import UUID, uuid4

from fastapi import APIRouter, Depends, File, Query, UploadFile, status
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company, get_current_company_membership, require_company_admin
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
    CompanyMemberRoleUpdate,
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


def get_company_member_service(
    session: Annotated[
        AsyncSession,
        Depends(get_db_session),
    ],
) -> CompanyMemberService:
    return CompanyMemberService(CompanyMemberRepository(session))


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
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
    service: Annotated[
        CompanyService,
        Depends(get_company_service),
    ],
    member_service: Annotated[
        CompanyMemberService,
        Depends(get_company_member_service),
    ],
    active_only: bool = True,
    limit: Annotated[int, Query(ge=1, le=200)] = 100,
    offset: Annotated[int, Query(ge=0)] = 0,
) -> PaginatedResponse[CompanyResponse]:
    memberships = await member_service.list_active_memberships_by_user(
        user_id=current_user.id
    )
    company_ids = [membership.company_id for membership in memberships]

    companies, count = await service.list_companies_by_ids(
        company_ids=company_ids,
        active_only=active_only,
        limit=limit,
        offset=offset,
    )

    return PaginatedResponse[CompanyResponse](
        message="Companies retrieved successfully.",
        count=count,
        limit=limit,
        offset=offset,
        data=[CompanyResponse.model_validate(company) for company in companies],
    )


@router.get(
    "/{company_id}",
    response_model=APIResponse[CompanyResponse],
)
async def get_company(
    company_id: UUID,
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
    service: Annotated[
        CompanyService,
        Depends(get_company_service),
    ],
    member_service: Annotated[
        CompanyMemberService,
        Depends(get_company_member_service),
    ],
) -> APIResponse[CompanyResponse]:
    from app.core.exceptions import ApplicationError

    membership = await member_service.get_active_membership(
        user_id=current_user.id,
        company_id=company_id,
    )
    if membership is None:
        # Return 404 rather than revealing whether another tenant's company exists.
        raise ApplicationError(
            message="Company not found.",
            error_code="COMPANY_NOT_FOUND",
            status_code=404,
        )

    company = await service.get_company(company_id)
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
    _admin_membership=Depends(require_company_admin),
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
    _admin_membership=Depends(require_company_admin),
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


@router.get("/me/access", response_model=APIResponse[dict])
async def get_my_company_access(
    membership=Depends(get_current_company_membership),
):
    role = membership.role.value if hasattr(membership.role, "value") else str(membership.role)
    return APIResponse(
        message="Company access retrieved.",
        data={
            "role": role,
            "can_write_finance": role in {"owner", "admin", "cfo", "finance_manager", "accountant"},
            "can_reset_all": role in {"owner", "admin"},
            "can_manage_members": role in {"owner", "admin"},
        },
    )


@router.get("/me/members", response_model=APIResponse[list[dict]])
async def list_company_members(
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    rows = (await session.execute(text("""
        SELECT cm.id, cm.user_id, cm.role::text AS role, cm.is_active, cm.joined_at,
               COALESCE(p.full_name, 'Workspace user') AS full_name, p.job_title
        FROM public.company_members cm
        LEFT JOIN public.profiles p ON p.id=cm.user_id
        WHERE cm.company_id=:company_id AND cm.is_active=true
        ORDER BY CASE WHEN cm.role::text='owner' THEN 0 WHEN cm.role::text='admin' THEN 1 ELSE 2 END, cm.joined_at
    """), {"company_id": current_company.id})).mappings().all()
    return APIResponse(message="Company members retrieved.", data=[dict(row) for row in rows])


@router.patch("/me/members/{member_id}/role", response_model=APIResponse[dict])
async def update_company_member_role(
    member_id: UUID,
    request: CompanyMemberRoleUpdate,
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    admin_membership=Depends(require_company_admin),
):
    from app.core.exceptions import ApplicationError
    allowed = {"admin", "cfo", "finance_manager", "accountant", "board_member", "viewer"}
    if request.role not in allowed:
        raise ApplicationError(message="Unsupported company role.", error_code="INVALID_COMPANY_ROLE", status_code=422)
    row = (await session.execute(text("""
        SELECT id, user_id, role::text AS role FROM public.company_members
        WHERE id=:member_id AND company_id=:company_id AND is_active=true
    """), {"member_id": member_id, "company_id": current_company.id})).mappings().first()
    if not row:
        raise ApplicationError(message="Company member was not found.", error_code="COMPANY_MEMBER_NOT_FOUND", status_code=404)
    if row["role"] == "owner":
        raise ApplicationError(message="Owner role cannot be changed here. Transfer ownership explicitly first.", error_code="OWNER_ROLE_PROTECTED", status_code=409)
    await session.execute(text("""
        UPDATE public.company_members SET role=CAST(:role AS public.company_role), updated_at=now()
        WHERE id=:member_id AND company_id=:company_id
    """), {"role": request.role, "member_id": member_id, "company_id": current_company.id})
    await session.commit()
    return APIResponse(message="Member role updated.", data={"member_id": str(member_id), "role": request.role})
