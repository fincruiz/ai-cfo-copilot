from __future__ import annotations
from datetime import datetime, timezone
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.core.exceptions import ApplicationError
from app.services.billing.razorpay_provider import RazorpayBillingProvider
from app.services.billing.certification import (
    assert_checkout_allowed,
    provider_checks,
    razorpay_subscription_status,
    stripe_subscription_status,
)
from app.services.billing.stripe_provider import StripeBillingProvider
from app.services.market_service import resolve_market
from app.services.subscription_service import SubscriptionService


class BillingService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def _subscription_row(self, company_id: UUID):
        await SubscriptionService(self.session).status(company_id=company_id)
        return (
            await self.session.execute(
                text("SELECT * FROM public.company_subscriptions WHERE company_id=:company_id"),
                {"company_id": company_id},
            )
        ).mappings().one()

    def _provider(self, billing_country_code: str):
        market = resolve_market(billing_country_code)
        if market.market_code == "IN":
            return RazorpayBillingProvider()
        return StripeBillingProvider()

    async def create_checkout(self, *, company_id: UUID, email: str | None, plan: str, interval: str) -> dict:
        row = await self._subscription_row(company_id)
        billing_country = str(row.get("billing_country_code") or "GLOBAL")
        assert_checkout_allowed(billing_country)
        provider = self._provider(billing_country)
        root = settings.billing_frontend_url.rstrip("/")
        result = await provider.create_checkout(
            company_id=company_id,
            customer_email=email,
            plan=plan,
            interval=interval,
            success_url=f"{root}/dashboard/subscription?checkout=success",
            cancel_url=f"{root}/dashboard/subscription?checkout=cancelled",
        )
        await self.session.execute(
            text("""
                UPDATE public.company_subscriptions
                SET provider=:provider,
                    last_checkout_id=:checkout_id,
                    requested_plan=:plan,
                    requested_interval=:interval,
                    change_requested_at=now(),
                    updated_at=now()
                WHERE company_id=:company_id
            """),
            {
                "provider": result.provider,
                "checkout_id": result.provider_session_id,
                "plan": plan,
                "interval": interval,
                "company_id": company_id,
            },
        )
        await self.session.commit()
        return result.__dict__

    async def create_portal(self, *, company_id: UUID) -> dict:
        row = await self._subscription_row(company_id)
        provider_name = str(row.get("provider") or "")
        customer_id = row.get("provider_customer_id")
        if provider_name != "stripe" or not customer_id:
            raise ApplicationError(message="A self-service billing portal is not available for this subscription.", error_code="BILLING_PORTAL_UNAVAILABLE", status_code=409)
        root = settings.billing_frontend_url.rstrip("/")
        url = await StripeBillingProvider().create_portal(provider_customer_id=str(customer_id), return_url=f"{root}/dashboard/subscription")
        return {"provider": "stripe", "url": url}

    async def _event_seen(self, provider: str, event_id: str) -> bool:
        exists = await self.session.scalar(
            text("SELECT 1 FROM public.billing_events WHERE provider=:provider AND provider_event_id=:event_id"),
            {"provider": provider, "event_id": event_id},
        )
        return bool(exists)

    async def _record_event(self, *, provider: str, event_id: str, event_type: str, company_id: UUID | None, payload: dict) -> None:
        await self.session.execute(
            text("""
                INSERT INTO public.billing_events(provider,provider_event_id,event_type,company_id,payload)
                VALUES (:provider,:event_id,:event_type,:company_id,CAST(:payload AS jsonb))
                ON CONFLICT (provider,provider_event_id) DO NOTHING
            """),
            {"provider": provider, "event_id": event_id, "event_type": event_type, "company_id": company_id, "payload": __import__("json").dumps(payload)},
        )
        if company_id:
            await self.session.execute(
                text("UPDATE public.company_subscriptions SET last_billing_event_at=now(), updated_at=now() WHERE company_id=:company_id"),
                {"company_id": company_id},
            )

    async def _company_from_provider_refs(self, *, provider: str, subscription_id: str | None = None, customer_id: str | None = None, metadata: dict | None = None) -> UUID | None:
        raw = (metadata or {}).get("company_id")
        if raw:
            try:
                return UUID(str(raw))
            except ValueError:
                pass
        row = (
            await self.session.execute(
                text("""
                    SELECT company_id FROM public.company_subscriptions
                    WHERE provider=:provider
                      AND (
                        (:subscription_id IS NOT NULL AND provider_subscription_id=:subscription_id)
                        OR (:customer_id IS NOT NULL AND provider_customer_id=:customer_id)
                        OR (:subscription_id IS NOT NULL AND last_checkout_id=:subscription_id)
                      )
                    LIMIT 1
                """),
                {"provider": provider, "subscription_id": subscription_id, "customer_id": customer_id},
            )
        ).scalar_one_or_none()
        return row

    async def readiness(self, *, company_id: UUID) -> dict:
        row = await self._subscription_row(company_id)
        country = str(row.get("billing_country_code") or "GLOBAL")
        provider, mode, checks = provider_checks(country)
        stats = (
            await self.session.execute(
                text("""
                    SELECT count(*) AS event_count, max(created_at) AS last_event_at
                    FROM public.billing_events
                    WHERE company_id=:company_id AND provider=:provider
                """),
                {"company_id": company_id, "provider": provider},
            )
        ).mappings().one()
        statuses = {item.status for item in checks}
        overall = "blocked" if "blocked" in statuses else "attention" if "attention" in statuses else "ready"
        return {
            "provider": provider,
            "mode": mode,
            "status": overall,
            "checks": [item.__dict__ for item in checks],
            "recent_verified_events": int(stats["event_count"] or 0),
            "last_verified_event_at": stats["last_event_at"].isoformat() if stats["last_event_at"] else None,
        }

    async def recent_events(self, *, company_id: UUID, limit: int = 10) -> list[dict]:
        rows = (
            await self.session.execute(
                text("""
                    SELECT provider,event_type,created_at
                    FROM public.billing_events
                    WHERE company_id=:company_id
                    ORDER BY created_at DESC
                    LIMIT :limit
                """),
                {"company_id": company_id, "limit": max(1, min(limit, 25))},
            )
        ).mappings().all()
        return [
            {"provider": str(row["provider"]), "event_type": str(row["event_type"]), "created_at": row["created_at"].isoformat()}
            for row in rows
        ]

    async def handle_stripe_event(self, event: dict) -> bool:
        event_id = str(event.get("id") or "")
        event_type = str(event.get("type") or "")
        if not event_id or await self._event_seen("stripe", event_id):
            return True
        obj = ((event.get("data") or {}).get("object") or {})
        metadata = obj.get("metadata") or {}
        subscription_id = obj.get("subscription") or (obj.get("id") if event_type.startswith("customer.subscription.") else None)
        customer_id = obj.get("customer")
        company_id = await self._company_from_provider_refs(provider="stripe", subscription_id=subscription_id, customer_id=customer_id, metadata=metadata)
        await self._record_event(provider="stripe", event_id=event_id, event_type=event_type, company_id=company_id, payload=event)
        if company_id:
            plan = metadata.get("plan")
            interval = metadata.get("billing_interval")
            if event_type == "checkout.session.completed":
                await self.session.execute(
                    text("""UPDATE public.company_subscriptions SET provider='stripe',provider_customer_id=COALESCE(:customer,provider_customer_id),provider_subscription_id=COALESCE(:subscription,provider_subscription_id),updated_at=now() WHERE company_id=:company_id"""),
                    {"customer": customer_id, "subscription": subscription_id, "company_id": company_id},
                )
            elif event_type == "invoice.paid":
                await self.session.execute(
                    text("""UPDATE public.company_subscriptions SET status='active',plan=COALESCE(requested_plan,plan),billing_interval=COALESCE(requested_interval,billing_interval),requested_plan=NULL,requested_interval=NULL,current_period_ends_at=COALESCE(to_timestamp(:period_end),current_period_ends_at),updated_at=now() WHERE company_id=:company_id"""),
                    {"period_end": int(obj.get("period_end") or 0) or None, "company_id": company_id},
                )
            elif event_type in {"invoice.payment_failed", "invoice.payment_action_required"}:
                await self.session.execute(text("UPDATE public.company_subscriptions SET status='past_due',payment_failure_at=now(),updated_at=now() WHERE company_id=:company_id"), {"company_id": company_id})
            elif event_type == "customer.subscription.deleted":
                await self.session.execute(text("UPDATE public.company_subscriptions SET status='cancelled',cancellation_requested_at=COALESCE(cancellation_requested_at,now()),updated_at=now() WHERE company_id=:company_id"), {"company_id": company_id})
            elif event_type == "customer.subscription.paused":
                await self.session.execute(text("UPDATE public.company_subscriptions SET status='past_due',updated_at=now() WHERE company_id=:company_id"), {"company_id": company_id})
            elif event_type in {"customer.subscription.updated", "customer.subscription.created"}:
                status = str(obj.get("status") or "")
                mapped = stripe_subscription_status(event_type, status)
                if mapped:
                    await self.session.execute(
                        text("UPDATE public.company_subscriptions SET status=:status,provider_customer_id=COALESCE(:customer,provider_customer_id),provider_subscription_id=COALESCE(:subscription,provider_subscription_id),updated_at=now() WHERE company_id=:company_id"),
                        {"status": mapped, "customer": customer_id, "subscription": subscription_id, "company_id": company_id},
                    )
        await self.session.commit()
        return False

    async def handle_razorpay_event(self, event: dict) -> bool:
        event_id = str(event.get("id") or event.get("event_id") or "")
        event_type = str(event.get("event") or "")
        if not event_id:
            event_id = __import__("hashlib").sha256(__import__("json").dumps(event, sort_keys=True).encode()).hexdigest()
        if await self._event_seen("razorpay", event_id):
            return True
        subscription = (((event.get("payload") or {}).get("subscription") or {}).get("entity") or {})
        subscription_id = subscription.get("id")
        metadata = subscription.get("notes") or {}
        company_id = await self._company_from_provider_refs(provider="razorpay", subscription_id=subscription_id, metadata=metadata)
        await self._record_event(provider="razorpay", event_id=event_id, event_type=event_type, company_id=company_id, payload=event)
        if company_id:
            mapped_status = razorpay_subscription_status(event_type)
            if mapped_status == "active":
                await self.session.execute(
                    text("""UPDATE public.company_subscriptions SET provider='razorpay',provider_subscription_id=COALESCE(:subscription,provider_subscription_id),status='active',plan=COALESCE(requested_plan,plan),billing_interval=COALESCE(requested_interval,billing_interval),requested_plan=NULL,requested_interval=NULL,updated_at=now() WHERE company_id=:company_id"""),
                    {"subscription": subscription_id, "company_id": company_id},
                )
            elif mapped_status == "past_due":
                await self.session.execute(text("UPDATE public.company_subscriptions SET status='past_due',payment_failure_at=now(),updated_at=now() WHERE company_id=:company_id"), {"company_id": company_id})
            elif mapped_status == "cancelled":
                await self.session.execute(text("UPDATE public.company_subscriptions SET status='cancelled',updated_at=now() WHERE company_id=:company_id"), {"company_id": company_id})
        await self.session.commit()
        return False
