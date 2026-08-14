from typing import Any
from uuid import UUID
from pydantic import BaseModel, EmailStr, Field


class LoginRequest(BaseModel):
    email: EmailStr
    password: str


class RefreshTokenRequest(BaseModel):
    refresh_token: str = Field(min_length=1)


class SignupRequest(BaseModel):
    email: EmailStr
    password: str = Field(min_length=8)
    full_name: str = Field(min_length=2, max_length=120)
    company_details: dict[str, Any]
    reporting_preferences: dict[str, Any] = {}
    enabled_modules: list[str] = []
    preferred_data_source: str | None = None


class SignupResponse(BaseModel):
    confirmation_required: bool
    email: EmailStr
    access_token: str | None = None
    refresh_token: str | None = None
    expires_in: int | None = None


class TokenResponse(BaseModel):
    access_token: str
    token_type: str = "bearer"
    expires_in: int | None = None
    refresh_token: str | None = None


class CurrentUser(BaseModel):
    id: UUID
    email: EmailStr | None = None
    phone: str | None = None
    role: str | None = None
    aud: str | None = None
    user_metadata: dict[str, Any] = {}
