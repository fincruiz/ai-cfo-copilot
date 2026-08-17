from __future__ import annotations

from datetime import datetime, timezone
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


PLAN_ENTITLEMENTS: dict[str, dict[str, object]] = {
    "trial": {
        "ai_queries_monthly": 200,
        "users": 3,
        "integrations": 1,
        "forecasting": True,
        "decision_simulator": True,
        "board_packs": True,
        "benchmarking": True,
        "audit_history_days": 30,
    },
    "founding": {
        "ai_queries_monthly": 2000,
        "users": 10,
        "integrations": 3,
        "forecasting": True,
        "decision_simulator": True,
        "board_packs": True,
        "benchmarking": True,
        "audit_history_days": 365,
    },
    "growth": {
        "ai_queries_monthly": 5000,
        "users": 25,
        "integrations": 5,
        "forecasting": True,
        "decision_simulator": True,
        "board_packs": True,
        "benchmarking": True,
        "audit_history_days": 730,
    },
    "enterprise": {
        "ai_queries_monthly": -1,
        "users": -1,
        "integrations": -1,
        "forecasting": True,
        "decision_simulator": True,
        "board_packs": True,
        "benchmarking": True,
        "audit_history_days": -1,
    },
}

ACTIVE_STATUSES = {"trialing", "active"}


def entitlements_for_plan(plan: str, overrides: dict | None = None) -> dict:
    base = dict(PLAN_ENTITLEMENTS.get(plan, PLAN_ENTITLEMENTS["trial"]))
    for key, value in (overrides or {}).items():
        if key in base:
            base[key] = value
    return base


def days_remaining(trial_ends_at: datetime | None, now: datetime | None = None) -> int | None:
    if trial_ends_at is None:
        return None
    now = now or datetime.now(timezone.utc)
    if trial_ends_at.tzinfo is None:
        trial_ends_at = trial_ends_at.replace(tzinfo=timezone.utc)
    return max(0, (trial_ends_at - now).days)


class SubscriptionService:
    def __init__(self, session: AsyncSession):
        self.session = session

    async def status(self, *, company_id: UUID) -> dict:
        row = (
            await self.session.execute(
                text(
                    """
                    SELECT plan, status, trial_started_at, trial_ends_at,
                           current_period_ends_at, entitlements
                    FROM public.company_subscriptions
                    WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().first()

        if not row:
            # Safe fallback for a deployment where migration ran after a company was
            # created but before its trigger existed.
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.company_subscriptions(company_id, plan, status)
                    VALUES (:company_id, 'trial', 'trialing')
                    ON CONFLICT (company_id) DO NOTHING
                    """
                ),
                {"company_id": company_id},
            )
            await self.session.commit()
            row = (
                await self.session.execute(
                    text(
                        """
                        SELECT plan, status, trial_started_at, trial_ends_at,
                               current_period_ends_at, entitlements
                        FROM public.company_subscriptions
                        WHERE company_id=:company_id
                        """
                    ),
                    {"company_id": company_id},
                )
            ).mappings().first()

        data = dict(row or {})
        plan = str(data.get("plan") or "trial")
        status = str(data.get("status") or "trialing")
        trial_end = data.get("trial_ends_at")

        if status == "trialing" and trial_end and days_remaining(trial_end) == 0:
            status = "expired"

        return {
            "plan": plan,
            "status": status,
            "trial_started_at": data.get("trial_started_at"),
            "trial_ends_at": trial_end,
            "current_period_ends_at": data.get("current_period_ends_at"),
            "days_remaining": days_remaining(trial_end),
            "entitlements": entitlements_for_plan(plan, data.get("entitlements") or {}),
            "is_access_active": status in ACTIVE_STATUSES,
            "billing_managed_externally": True,
        }
