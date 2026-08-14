from __future__ import annotations

from uuid import UUID
from pathlib import Path

import httpx
from sqlalchemy import func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.database.models.core.company_member import CompanyMember
from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService


class AccountDeletionService:
    """Permanently removes a user's FinCruiz account and owned single-user workspaces."""

    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def delete_account(self, *, user_id: UUID) -> dict:
        if not settings.supabase_service_role_key:
            raise ApplicationError(
                message=(
                    "Account deletion is temporarily unavailable because the server-side "
                    "Supabase admin credential is not configured."
                ),
                error_code="AUTH_ADMIN_KEY_NOT_CONFIGURED",
                status_code=503,
            )

        memberships = list(
            (
                await self.session.execute(
                    select(CompanyMember).where(
                        CompanyMember.user_id == user_id,
                        CompanyMember.is_active.is_(True),
                    )
                )
            ).scalars().all()
        )

        # Avoid deleting a shared company or leaving it ownerless. Multi-user
        # workspaces must transfer ownership before the owner deletes the account.
        for membership in memberships:
            if str(membership.role.value if hasattr(membership.role, "value") else membership.role) != "owner":
                continue
            other_members = int(
                (
                    await self.session.execute(
                        select(func.count())
                        .select_from(CompanyMember)
                        .where(
                            CompanyMember.company_id == membership.company_id,
                            CompanyMember.user_id != user_id,
                            CompanyMember.is_active.is_(True),
                        )
                    )
                ).scalar_one()
            )
            if other_members > 0:
                raise ApplicationError(
                    message=(
                        "This account owns a workspace with other active members. "
                        "Transfer workspace ownership before deleting your profile."
                    ),
                    error_code="OWNERSHIP_TRANSFER_REQUIRED",
                    status_code=409,
                )

        companies_to_delete = [
            membership.company_id
            for membership in memberships
            if str(membership.role.value if hasattr(membership.role, "value") else membership.role) == "owner"
        ]

        logo_paths = []
        for company_id in companies_to_delete:
            logo_path = (
                await self.session.execute(
                    text("SELECT logo_path FROM public.companies WHERE id=:company_id"),
                    {"company_id": company_id},
                )
            ).scalar_one_or_none()
            if logo_path:
                logo_paths.append(str(logo_path))

        lifecycle = WorkspaceLifecycleService(self.session)
        for company_id in companies_to_delete:
            await lifecycle.reset_financial_data(company_id=company_id)
            await self.session.execute(
                text("DELETE FROM public.branches WHERE company_id=:company_id"),
                {"company_id": company_id},
            )
            if await lifecycle._table_exists("company_preferences"):
                await self.session.execute(
                    text("DELETE FROM public.company_preferences WHERE company_id=:company_id"),
                    {"company_id": company_id},
                )

        # Delete memberships before owned companies so this remains compatible
        # with deployments whose core foreign keys were created without cascade.
        membership_result = await self.session.execute(
            text("DELETE FROM public.company_members WHERE user_id=:user_id"),
            {"user_id": user_id},
        )
        memberships_deleted = int(membership_result.rowcount or 0)

        companies_deleted = 0
        for company_id in companies_to_delete:
            result = await self.session.execute(
                text("DELETE FROM public.companies WHERE id=:company_id"),
                {"company_id": company_id},
            )
            companies_deleted += int(result.rowcount or 0)

        profile_result = await self.session.execute(
            text("DELETE FROM public.profiles WHERE id=:user_id"),
            {"user_id": user_id},
        )
        profile_deleted = bool(profile_result.rowcount)
        await self.session.commit()

        for logo_path in logo_paths:
            if logo_path.startswith("/uploads/logos/"):
                try:
                    Path(logo_path.lstrip("/")).unlink(missing_ok=True)
                except OSError:
                    pass

        # Supabase auth.users is removed through the server-only Admin API.
        url = f"{settings.supabase_url.rstrip('/')}/auth/v1/admin/users/{user_id}"
        headers = {
            "apikey": settings.supabase_service_role_key,
            "Authorization": f"Bearer {settings.supabase_service_role_key}",
        }
        try:
            async with httpx.AsyncClient(timeout=20.0) as client:
                response = await client.delete(url, headers=headers)
        except httpx.RequestError as exc:
            raise ApplicationError(
                message="Local account data was deleted, but the authentication service could not be reached.",
                error_code="AUTH_ACCOUNT_DELETE_UNAVAILABLE",
                status_code=503,
                details={"local_data_deleted": True},
            ) from exc

        if response.status_code not in {200, 204}:
            raise ApplicationError(
                message="Local account data was deleted, but the authentication identity could not be removed.",
                error_code="AUTH_ACCOUNT_DELETE_FAILED",
                status_code=502,
                details={"local_data_deleted": True, "auth_status": response.status_code},
            )

        return {
            "auth_user_deleted": True,
            "companies_deleted": companies_deleted,
            "memberships_deleted": memberships_deleted,
            "profile_deleted": profile_deleted,
        }
