from __future__ import annotations
import httpx
from uuid import UUID

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.services.billing.base import CheckoutResult


class RazorpayBillingProvider:
    name = "razorpay"
    api_base = "https://api.razorpay.com/v1"

    def _plan_id(self, plan: str, interval: str) -> str:
        value = {
            ("founding", "monthly"): settings.razorpay_starter_monthly_plan_id,
            ("founding", "annual"): settings.razorpay_starter_annual_plan_id,
            ("growth", "monthly"): settings.razorpay_growth_monthly_plan_id,
            ("growth", "annual"): settings.razorpay_growth_annual_plan_id,
        }.get((plan, interval))
        if not value:
            raise ApplicationError(
                message="Razorpay billing is not configured for this plan and billing interval.",
                error_code="BILLING_PLAN_NOT_CONFIGURED",
                status_code=503,
            )
        return value

    async def create_checkout(self, *, company_id: UUID, customer_email: str | None, plan: str, interval: str, success_url: str, cancel_url: str) -> CheckoutResult:
        if not settings.razorpay_key_id or not settings.razorpay_key_secret:
            raise ApplicationError(message="Razorpay billing is not configured.", error_code="BILLING_PROVIDER_NOT_CONFIGURED", status_code=503)
        total_count = 12 if interval == "monthly" else 5
        body = {
            "plan_id": self._plan_id(plan, interval),
            "total_count": total_count,
            "quantity": 1,
            "customer_notify": 1,
            "notes": {
                "company_id": str(company_id),
                "plan": plan,
                "billing_interval": interval,
            },
        }
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                f"{self.api_base}/subscriptions",
                json=body,
                auth=(settings.razorpay_key_id, settings.razorpay_key_secret),
            )
        if response.status_code >= 400:
            raise ApplicationError(message="Razorpay could not start subscription checkout.", error_code="BILLING_CHECKOUT_FAILED", status_code=502)
        payload = response.json()
        subscription_id = str(payload["id"])
        return CheckoutResult(
            provider=self.name,
            provider_session_id=subscription_id,
            public_key=settings.razorpay_key_id,
            subscription_id=subscription_id,
            plan=plan,
            billing_interval=interval,
        )

    async def create_portal(self, *, provider_customer_id: str, return_url: str) -> str:
        raise ApplicationError(
            message="Self-service Razorpay billing management is not enabled yet. Use the FinCruiz subscription page or contact support.",
            error_code="BILLING_PORTAL_UNAVAILABLE",
            status_code=409,
        )
