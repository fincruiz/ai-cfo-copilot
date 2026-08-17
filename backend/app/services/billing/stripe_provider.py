from __future__ import annotations
import httpx
from uuid import UUID

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.services.billing.base import CheckoutResult


class StripeBillingProvider:
    name = "stripe"
    api_base = "https://api.stripe.com/v1"

    def _price_id(self, plan: str, interval: str) -> str:
        key = {
            ("founding", "monthly"): settings.stripe_starter_monthly_price_id,
            ("founding", "annual"): settings.stripe_starter_annual_price_id,
            ("growth", "monthly"): settings.stripe_growth_monthly_price_id,
            ("growth", "annual"): settings.stripe_growth_annual_price_id,
        }.get((plan, interval))
        if not key:
            raise ApplicationError(
                message="Stripe billing is not configured for this plan and billing interval.",
                error_code="BILLING_PRICE_NOT_CONFIGURED",
                status_code=503,
            )
        return key

    async def create_checkout(self, *, company_id: UUID, customer_email: str | None, plan: str, interval: str, success_url: str, cancel_url: str) -> CheckoutResult:
        if not settings.stripe_secret_key:
            raise ApplicationError(message="Stripe billing is not configured.", error_code="BILLING_PROVIDER_NOT_CONFIGURED", status_code=503)
        data = {
            "mode": "subscription",
            "success_url": success_url,
            "cancel_url": cancel_url,
            "client_reference_id": str(company_id),
            "line_items[0][price]": self._price_id(plan, interval),
            "line_items[0][quantity]": "1",
            "metadata[company_id]": str(company_id),
            "metadata[plan]": plan,
            "metadata[billing_interval]": interval,
            "subscription_data[metadata][company_id]": str(company_id),
            "subscription_data[metadata][plan]": plan,
            "subscription_data[metadata][billing_interval]": interval,
        }
        if customer_email:
            data["customer_email"] = customer_email
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                f"{self.api_base}/checkout/sessions",
                data=data,
                auth=(settings.stripe_secret_key, ""),
            )
        if response.status_code >= 400:
            raise ApplicationError(message="Stripe could not start checkout.", error_code="BILLING_CHECKOUT_FAILED", status_code=502)
        payload = response.json()
        return CheckoutResult(
            provider=self.name,
            provider_session_id=str(payload["id"]),
            checkout_url=payload.get("url"),
            plan=plan,
            billing_interval=interval,
        )

    async def create_portal(self, *, provider_customer_id: str, return_url: str) -> str:
        if not settings.stripe_secret_key:
            raise ApplicationError(message="Stripe billing is not configured.", error_code="BILLING_PROVIDER_NOT_CONFIGURED", status_code=503)
        async with httpx.AsyncClient(timeout=30) as client:
            response = await client.post(
                f"{self.api_base}/billing_portal/sessions",
                data={"customer": provider_customer_id, "return_url": return_url},
                auth=(settings.stripe_secret_key, ""),
            )
        if response.status_code >= 400:
            raise ApplicationError(message="Stripe billing portal is unavailable.", error_code="BILLING_PORTAL_FAILED", status_code=502)
        return str(response.json()["url"])
