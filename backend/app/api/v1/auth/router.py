from typing import Annotated

from fastapi import APIRouter, Depends

from app.dependencies.auth import get_current_user
from app.schemas.auth import (
    CurrentUser,
    LoginRequest,
    SignupRequest,
    SignupResponse,
    TokenResponse,
)
from app.schemas.responses import APIResponse
from app.services.auth_service import AuthService


router = APIRouter(
    prefix="/auth",
    tags=["Authentication"],
)




@router.post("/signup", response_model=APIResponse[SignupResponse], status_code=201)
async def signup(signup_data: SignupRequest) -> APIResponse[SignupResponse]:
    auth_service = AuthService()
    result = await auth_service.signup(
        email=signup_data.email,
        password=signup_data.password,
        full_name=signup_data.full_name,
        company_details=signup_data.company_details,
        reporting_preferences=signup_data.reporting_preferences,
        enabled_modules=signup_data.enabled_modules,
        preferred_data_source=signup_data.preferred_data_source,
    )
    access_token = result.get("access_token")
    return APIResponse(
        message="Account created. Confirm your email before signing in." if not access_token else "Account created successfully.",
        data=SignupResponse(
            confirmation_required=not bool(access_token),
            email=signup_data.email,
            access_token=access_token,
            refresh_token=result.get("refresh_token"),
            expires_in=result.get("expires_in"),
        ),
    )


@router.post(
    "/login",
    response_model=APIResponse[TokenResponse],
)
async def login(
    login_data: LoginRequest,
) -> APIResponse[TokenResponse]:
    auth_service = AuthService()

    token_data = await auth_service.login(
        email=login_data.email,
        password=login_data.password,
    )

    token = TokenResponse(
        access_token=token_data["access_token"],
        token_type=token_data.get("token_type", "bearer"),
        expires_in=token_data.get("expires_in"),
        refresh_token=token_data.get("refresh_token"),
    )

    return APIResponse[TokenResponse](
        message="Login successful.",
        data=token,
    )


@router.get(
    "/me",
    response_model=APIResponse[CurrentUser],
)
async def get_me(
    current_user: Annotated[
        CurrentUser,
        Depends(get_current_user),
    ],
) -> APIResponse[CurrentUser]:
    return APIResponse[CurrentUser](
        message="Authenticated user retrieved successfully.",
        data=current_user,
    )