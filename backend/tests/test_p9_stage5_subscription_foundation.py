from datetime import datetime, timedelta, timezone

from app.services.subscription_service import days_remaining, entitlements_for_plan


def test_founding_plan_keeps_beta_capabilities_enabled():
    e = entitlements_for_plan("founding")
    assert e["decision_simulator"] is True
    assert e["forecasting"] is True
    assert e["integrations"] == 3
    assert e["ai_queries_monthly"] >= 1000


def test_entitlement_override_is_whitelisted_to_known_capabilities():
    e = entitlements_for_plan("trial", {"users": 7, "unknown_capability": True})
    assert e["users"] == 7
    assert "unknown_capability" not in e


def test_trial_days_never_go_negative():
    now = datetime(2026, 8, 17, tzinfo=timezone.utc)
    assert days_remaining(now + timedelta(days=8, hours=2), now) == 8
    assert days_remaining(now - timedelta(days=1), now) == 0
    assert days_remaining(None, now) is None
