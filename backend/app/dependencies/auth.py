from typing import Annotated, Any
from uuid import UUID

from fastapi import Depends
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from jose import JWTError, jwt

from app.core.exceptions import ApplicationError
from app.schemas.auth import CurrentUser
from app.services.auth_service import AuthService


bearer_scheme = HTTPBearer(
    auto_error=False,
    description="Enter the Supabase access token.",
)


def decode_unverified_claims(access_token: str) -> dict[str, Any]:
    """
    Reads JWT claims only as a fallback for display fields.

    Security validation is still performed by Supabase through
    AuthService.get_user(). We do not trust this decoded payload
    for authorization decisions.
    """
    try:
        return jwt.get_unverified_claims(access_token)
    except JWTError:
        return {}


async def get_current_user(
    credentials: Annotated[
        HTTPAuthorizationCredentials | None,
        Depends(bearer_scheme),
    ],
) -> CurrentUser:
    if credentials is None:
        raise ApplicationError(
            message="Authentication is required.",
            error_code="AUTHENTICATION_REQUIRED",
            status_code=401,
        )

    access_token = credentials.credentials

    auth_service = AuthService()

    user_data = await auth_service.get_user(
        access_token=access_token,
    )

    claims = decode_unverified_claims(access_token)

    user_id = (
        user_data.get("id")
        or claims.get("sub")
    )

    if not user_id:
        raise ApplicationError(
            message="Authenticated user information is invalid.",
            error_code="INVALID_USER_DATA",
            status_code=401,
        )

    return CurrentUser(
        id=UUID(str(user_id)),
        email=(
            user_data.get("email")
            or claims.get("email")
        ),
        phone=(
            user_data.get("phone")
            or claims.get("phone")
            or None
        ),
        role=(
            user_data.get("role")
            or claims.get("role")
        ),
        aud=(
            user_data.get("aud")
            or claims.get("aud")
        ),
        user_metadata=(
            user_data.get("user_metadata")
            or claims.get("user_metadata")
            or {}
        ),
    )