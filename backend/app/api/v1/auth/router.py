from typing import Annotated

from fastapi import APIRouter, Depends
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer

from app.dependencies.auth import get_current_user
from app.schemas.auth import (
    CurrentUser,
    LoginRequest,
    RefreshTokenRequest,
    LogoutRequest,
    SignupRequest,
    SignupResponse,
    TokenResponse,
    ResendConfirmationRequest, PasswordRecoveryRequest, PasswordUpdateRequest,
)
from app.schemas.responses import APIResponse
from app.services.auth_service import AuthService
from app.core.exceptions import ApplicationError


router = APIRouter(
    prefix="/auth",
    tags=["Authentication"],
)

logout_bearer = HTTPBearer(auto_error=False)




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


@router.post(
    "/refresh",
    response_model=APIResponse[TokenResponse],
)
async def refresh_session(
    request: RefreshTokenRequest,
) -> APIResponse[TokenResponse]:
    auth_service = AuthService()
    token_data = await auth_service.refresh_session(
        refresh_token=request.refresh_token,
    )

    token = TokenResponse(
        access_token=token_data["access_token"],
        token_type=token_data.get("token_type", "bearer"),
        expires_in=token_data.get("expires_in"),
        refresh_token=token_data.get("refresh_token") or request.refresh_token,
    )

    return APIResponse[TokenResponse](
        message="Session refreshed successfully.",
        data=token,
    )


@router.post("/logout", response_model=APIResponse[dict])
async def logout(
    request: LogoutRequest,
    credentials: Annotated[
        HTTPAuthorizationCredentials | None,
        Depends(logout_bearer),
    ],
) -> APIResponse[dict]:
    if credentials is None:
        raise ApplicationError(
            message="Authentication is required.",
            error_code="AUTHENTICATION_REQUIRED",
            status_code=401,
        )

    await AuthService().logout(
        access_token=credentials.credentials,
        scope=request.scope,
    )
    return APIResponse(
        message="Signed out successfully.",
        data={"signed_out": True, "scope": request.scope},
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

@router.post("/resend-confirmation",response_model=APIResponse[dict])
async def resend_confirmation(request: ResendConfirmationRequest):
    await AuthService().resend_confirmation(email=request.email)
    return APIResponse(message="If the account is awaiting confirmation, a new confirmation email has been sent.",data={"sent":True})

@router.post("/forgot-password",response_model=APIResponse[dict])
async def forgot_password(request: PasswordRecoveryRequest):
    await AuthService().request_password_recovery(email=request.email)
    return APIResponse(message="If an account exists for that email, password reset instructions have been sent.",data={"sent":True})

@router.post("/reset-password",response_model=APIResponse[dict])
async def reset_password(request: PasswordUpdateRequest):
    await AuthService().update_password(access_token=request.access_token,password=request.password)
    return APIResponse(message="Password updated successfully.",data={"updated":True})
