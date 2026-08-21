from datetime import datetime, timedelta, timezone
from pathlib import Path

from app.core.config import Settings
from app.services.billing.certification import (
    razorpay_subscription_status,
    sandbox_lifecycle_coverage,
    stripe_subscription_status,
)
from app.services.paid_launch_certification_service import paid_launch_summary


ROOT = Path(__file__).resolve().parents[1]


def production_settings(now: datetime, **overrides) -> Settings:
    values = dict(
        environment="production",
        billing_frontend_url="https://app.fincruiz.example",
        billing_allow_live_payments=True,
        paid_launch_providers="stripe,razorpay",
        stripe_secret_key="sk_live_example",
        stripe_webhook_secret="whsec_live_example",
        stripe_starter_monthly_price_id="price_sm",
        stripe_starter_annual_price_id="price_sa",
        stripe_growth_monthly_price_id="price_gm",
        stripe_growth_annual_price_id="price_ga",
        razorpay_key_id="rzp_live_example",
        razorpay_key_secret="rzp_secret",
        razorpay_webhook_secret="rzp_webhook",
        razorpay_starter_monthly_plan_id="plan_sm",
        razorpay_starter_annual_plan_id="plan_sa",
        razorpay_growth_monthly_plan_id="plan_gm",
        razorpay_growth_annual_plan_id="plan_ga",
        stripe_sandbox_certified_at=now - timedelta(days=1),
        razorpay_sandbox_certified_at=now - timedelta(days=1),
        deployment_region="ap-south-1",
        database_region="ap-south-1",
        production_performance_certified_at=now - timedelta(days=1),
        import_staging_dir="/persistent/fincruiz/imports",
        import_staging_persistent=True,
        backup_restore_verified_at=now - timedelta(days=7),
        error_monitoring_dsn="https://monitoring.example/project",
        error_monitoring_verified_at=now - timedelta(days=1),
        support_contact_email="support@example.com",
        support_runbook_url="https://internal.example/runbooks/fincruiz",
    )
    values.update(overrides)
    return Settings(_env_file=None, **values)


def test_stripe_sandbox_lifecycle_contract_covers_success_failure_update_cancel():
    events = {
        "checkout.session.completed",
        "invoice.paid",
        "invoice.payment_failed",
        "customer.subscription.updated",
        "customer.subscription.deleted",
    }
    result = sandbox_lifecycle_coverage("stripe", events)
    assert result["complete"] is True
    assert result["missing"] == []
    assert stripe_subscription_status("invoice.paid") == "active"
    assert stripe_subscription_status("invoice.payment_failed") == "past_due"
    assert stripe_subscription_status("customer.subscription.updated", "trialing") == "active"
    assert stripe_subscription_status("customer.subscription.updated", "unpaid") == "past_due"
    assert stripe_subscription_status("customer.subscription.deleted") == "cancelled"


def test_razorpay_sandbox_lifecycle_contract_covers_success_failure_cancel():
    events = {"subscription.activated", "subscription.charged", "subscription.halted", "subscription.cancelled"}
    result = sandbox_lifecycle_coverage("razorpay", events)
    assert result["complete"] is True
    assert razorpay_subscription_status("subscription.activated") == "active"
    assert razorpay_subscription_status("subscription.charged") == "active"
    assert razorpay_subscription_status("subscription.halted") == "past_due"
    assert razorpay_subscription_status("subscription.cancelled") == "cancelled"


def test_missing_sandbox_failure_path_does_not_certify():
    result = sandbox_lifecycle_coverage("stripe", {"checkout.session.completed", "invoice.paid"})
    assert result["complete"] is False
    assert "invoice.payment_failed" in result["missing"]


def test_paid_launch_summary_goes_green_only_when_every_explicit_gate_is_ready():
    now = datetime(2026, 8, 21, 8, 0, tzinfo=timezone.utc)
    result = paid_launch_summary(production_settings(now), now=now)
    assert result["status"] == "ready"
    assert result["live_paid_launch_approved"] is True
    assert result["score"] == 100
    assert all(item["status"] == "ready" for item in result["checks"])


def test_live_payment_switch_is_fail_closed():
    now = datetime(2026, 8, 21, 8, 0, tzinfo=timezone.utc)
    result = paid_launch_summary(production_settings(now, billing_allow_live_payments=False), now=now)
    switch = next(item for item in result["checks"] if item["key"] == "live_payment_switch")
    assert switch["status"] == "blocked"
    assert result["live_paid_launch_approved"] is False


def test_non_production_or_unverified_operator_evidence_blocks_paid_launch():
    now = datetime(2026, 8, 21, 8, 0, tzinfo=timezone.utc)
    result = paid_launch_summary(
        production_settings(
            now,
            environment="development",
            backup_restore_verified_at=now - timedelta(days=120),
            production_performance_certified_at=None,
        ),
        now=now,
    )
    states = {item["key"]: item["status"] for item in result["checks"]}
    assert states["environment"] == "blocked"
    assert states["backup_restore"] == "blocked"
    assert states["production_performance"] == "blocked"
    assert result["status"] == "blocked"


def test_region_alignment_and_persistent_ingestion_are_required():
    now = datetime(2026, 8, 21, 8, 0, tzinfo=timezone.utc)
    result = paid_launch_summary(
        production_settings(now, database_region="ap-southeast-2", import_staging_persistent=False),
        now=now,
    )
    states = {item["key"]: item["status"] for item in result["checks"]}
    assert states["region_alignment"] == "blocked"
    assert states["persistent_ingestion"] == "blocked"


def test_stage10_release_script_and_operator_runbook_are_present():
    script = (ROOT / "scripts/launch_certify.py").read_text(encoding="utf-8")
    runbook = (ROOT / "docs/STAGE10_PAID_LAUNCH_CERTIFICATION.md").read_text(encoding="utf-8")
    assert "paid_launch_configuration_checks" in script
    assert "BILLING_ALLOW_LIVE_PAYMENTS" in runbook
    assert "restore drill" in runbook.lower()
    assert "load_test_finance" in runbook
