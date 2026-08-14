from typing import Any

import httpx

from app.core.config import settings
from app.core.exceptions import ApplicationError


class AuthService:
    def __init__(self) -> None:
        if not settings.supabase_url:
            raise RuntimeError("SUPABASE_URL is not configured.")

        if not settings.supabase_publishable_key:
            raise RuntimeError(
                "SUPABASE_PUBLISHABLE_KEY is not configured."
            )

        self.base_url = settings.supabase_url.rstrip("/")
        self.api_key = settings.supabase_publishable_key


    async def signup(
        self,
        *,
        email: str,
        password: str,
        full_name: str,
        company_details: dict[str, Any],
        reporting_preferences: dict[str, Any],
        enabled_modules: list[str],
        preferred_data_source: str | None,
    ) -> dict[str, Any]:
        url = f"{self.base_url}/auth/v1/signup"
        headers = {"apikey": self.api_key, "Content-Type": "application/json"}
        payload = {
            "email": email,
            "password": password,
            "data": {
                "full_name": full_name,
                "company_details": company_details,
                "reporting_preferences": reporting_preferences,
                "enabled_modules": enabled_modules,
                "preferred_data_source": preferred_data_source,
            },
        }
        try:
            async with httpx.AsyncClient(timeout=20.0) as client:
                response = await client.post(url, headers=headers, json=payload)
        except httpx.RequestError as exc:
            raise ApplicationError(
                message="Authentication service is unavailable.",
                error_code="AUTH_SERVICE_UNAVAILABLE",
                status_code=503,
            ) from exc
        if response.status_code not in {200, 201}:
            detail = response.json() if response.headers.get("content-type", "").startswith("application/json") else {}
            raise ApplicationError(
                message=detail.get("msg") or detail.get("message") or "Unable to create account.",
                error_code="SIGNUP_FAILED",
                status_code=422,
            )
        return response.json()

    async def login(
        self,
        *,
        email: str,
        password: str,
    ) -> dict[str, Any]:
        url = f"{self.base_url}/auth/v1/token"

        headers = {
            "apikey": self.api_key,
            "Content-Type": "application/json",
        }

        payload = {
            "email": email,
            "password": password,
        }

        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                response = await client.post(
                    url,
                    params={"grant_type": "password"},
                    headers=headers,
                    json=payload,
                )
        except httpx.RequestError as exc:
            raise ApplicationError(
                message="Authentication service is unavailable.",
                error_code="AUTH_SERVICE_UNAVAILABLE",
                status_code=503,
            ) from exc

        if response.status_code != 200:
            raise ApplicationError(
                message="Invalid email or password.",
                error_code="INVALID_CREDENTIALS",
                status_code=401,
            )

        return response.json()


    async def refresh_session(
        self,
        *,
        refresh_token: str,
    ) -> dict[str, Any]:
        url = f"{self.base_url}/auth/v1/token"
        headers = {
            "apikey": self.api_key,
            "Content-Type": "application/json",
        }
        payload = {"refresh_token": refresh_token}

        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                response = await client.post(
                    url,
                    params={"grant_type": "refresh_token"},
                    headers=headers,
                    json=payload,
                )
        except httpx.RequestError as exc:
            raise ApplicationError(
                message="Authentication service is unavailable.",
                error_code="AUTH_SERVICE_UNAVAILABLE",
                status_code=503,
            ) from exc

        if response.status_code != 200:
            raise ApplicationError(
                message="Your session has expired. Please sign in again.",
                error_code="INVALID_REFRESH_TOKEN",
                status_code=401,
            )

        return response.json()

    async def get_user(
        self,
        *,
        access_token: str,
    ) -> dict[str, Any]:
        url = f"{self.base_url}/auth/v1/user"

        headers = {
            "apikey": self.api_key,
            "Authorization": f"Bearer {access_token}",
        }

        try:
            async with httpx.AsyncClient(timeout=15.0) as client:
                response = await client.get(
                    url,
                    headers=headers,
                )
        except httpx.RequestError as exc:
            raise ApplicationError(
                message="Authentication service is unavailable.",
                error_code="AUTH_SERVICE_UNAVAILABLE",
                status_code=503,
            ) from exc

        if response.status_code != 200:
            raise ApplicationError(
                message="Authentication credentials are invalid or expired.",
                error_code="INVALID_AUTH_TOKEN",
                status_code=401,
            )

        return response.json()