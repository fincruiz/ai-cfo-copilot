from __future__ import annotations
from dataclasses import dataclass
from typing import Protocol
from uuid import UUID


@dataclass(frozen=True)
class CheckoutResult:
    provider: str
    provider_session_id: str
    plan: str
    billing_interval: str
    checkout_url: str | None = None
    public_key: str | None = None
    subscription_id: str | None = None


class BillingProvider(Protocol):
    name: str

    async def create_checkout(
        self,
        *,
        company_id: UUID,
        customer_email: str | None,
        plan: str,
        interval: str,
        success_url: str,
        cancel_url: str,
    ) -> CheckoutResult: ...

    async def create_portal(self, *, provider_customer_id: str, return_url: str) -> str: ...
