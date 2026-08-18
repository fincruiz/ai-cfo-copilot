from __future__ import annotations

from dataclasses import dataclass

from app.core.config import settings
from app.services.market_service import resolve_market


PLACEHOLDER_VALUES = {
    "", "dummy", "test", "placeholder", "changeme", "change_me", "todo",
    "your_key_here", "your_secret_here", "null", "none",
}


@dataclass(frozen=True)
class BillingCheck:
    key: str
    label: str
    status: str
    detail: str


def _present(value: str | None) -> bool:
    if value is None:
        return False
    cleaned = value.strip()
    return bool(cleaned) and cleaned.lower() not in PLACEHOLDER_VALUES


def stripe_mode(secret_key: str | None = None) -> str:
    value = (secret_key if secret_key is not None else settings.stripe_secret_key) or ""
    if value.startswith("sk_test_"):
        return "test"
    if value.startswith("sk_live_"):
        return "live"
    return "unknown"


def razorpay_mode(key_id: str | None = None) -> str:
    value = (key_id if key_id is not None else settings.razorpay_key_id) or ""
    if value.startswith("rzp_test_"):
        return "test"
    if value.startswith("rzp_live_"):
        return "live"
    return "unknown"


def provider_for_country(country_code: str | None) -> str:
    return "razorpay" if resolve_market(country_code).market_code == "IN" else "stripe"


def validate_frontend_url() -> BillingCheck:
    url = (settings.billing_frontend_url or "").strip()
    if settings.is_production:
        ok = url.startswith("https://") and "localhost" not in url and "127.0.0.1" not in url
        return BillingCheck(
            "frontend_url",
            "Billing return URL",
            "ready" if ok else "blocked",
            "Production billing return URL is public HTTPS." if ok else "Set BILLING_FRONTEND_URL to the public HTTPS FinCruiz frontend.",
        )
    return BillingCheck("frontend_url","Billing return URL","ready",f"Development return URL: {url or 'not set'}")


def provider_checks(country_code: str | None) -> tuple[str, str, list[BillingCheck]]:
    provider = provider_for_country(country_code)
    checks: list[BillingCheck] = [validate_frontend_url()]

    if provider == "razorpay":
        mode = razorpay_mode()
        required = {
            "razorpay_key_id": settings.razorpay_key_id,
            "razorpay_key_secret": settings.razorpay_key_secret,
            "razorpay_webhook_secret": settings.razorpay_webhook_secret,
            "razorpay_starter_monthly_plan_id": settings.razorpay_starter_monthly_plan_id,
            "razorpay_starter_annual_plan_id": settings.razorpay_starter_annual_plan_id,
            "razorpay_growth_monthly_plan_id": settings.razorpay_growth_monthly_plan_id,
            "razorpay_growth_annual_plan_id": settings.razorpay_growth_annual_plan_id,
        }
    else:
        mode = stripe_mode()
        required = {
            "stripe_secret_key": settings.stripe_secret_key,
            "stripe_webhook_secret": settings.stripe_webhook_secret,
            "stripe_starter_monthly_price_id": settings.stripe_starter_monthly_price_id,
            "stripe_starter_annual_price_id": settings.stripe_starter_annual_price_id,
            "stripe_growth_monthly_price_id": settings.stripe_growth_monthly_price_id,
            "stripe_growth_annual_price_id": settings.stripe_growth_annual_price_id,
        }

    missing = [key for key,value in required.items() if not _present(value)]
    checks.append(BillingCheck(
        "provider_configuration",
        f"{provider.title()} configuration",
        "ready" if not missing else "attention",
        "All required provider values are configured." if not missing else "Missing/placeholder: " + ", ".join(missing),
    ))

    if mode == "live" and not settings.billing_allow_live_payments:
        checks.append(BillingCheck(
            "live_payment_gate",
            "Live payment safety gate",
            "blocked",
            "Live provider credentials detected while BILLING_ALLOW_LIVE_PAYMENTS is false.",
        ))
    elif mode == "live":
        checks.append(BillingCheck("live_payment_gate","Live payment safety gate","ready","Live payments are explicitly enabled."))
    elif mode == "test":
        checks.append(BillingCheck("live_payment_gate","Provider mode","ready","Provider credentials are in test/sandbox mode."))
    else:
        checks.append(BillingCheck("live_payment_gate","Provider mode","attention","Provider mode cannot yet be identified from the configured key."))

    return provider, mode, checks


def assert_checkout_allowed(country_code: str | None) -> tuple[str, str]:
    provider, mode, checks = provider_checks(country_code)
    blocked = [check for check in checks if check.status == "blocked"]
    if blocked:
        from app.core.exceptions import ApplicationError
        raise ApplicationError(
            message=blocked[0].detail,
            error_code="BILLING_CERTIFICATION_BLOCKED",
            status_code=503,
        )
    if mode == "live" and not settings.billing_allow_live_payments:
        from app.core.exceptions import ApplicationError
        raise ApplicationError(
            message="Live payment processing is disabled until billing certification is complete.",
            error_code="LIVE_BILLING_DISABLED",
            status_code=503,
        )
    return provider, mode
