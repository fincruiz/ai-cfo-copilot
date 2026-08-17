from app.services.subscription_service import PLAN_ENTITLEMENTS, entitlements_for_plan
from app.services.market_service import PLAN_LABELS


def test_commercial_plan_labels_are_customer_facing():
    assert PLAN_LABELS["founding"] == "Starter"
    assert PLAN_LABELS["growth"] == "Growth"
    assert PLAN_LABELS["enterprise"] == "Scale / Enterprise"


def test_starter_and_growth_have_materially_different_entitlements():
    starter = entitlements_for_plan("founding")
    growth = entitlements_for_plan("growth")
    assert starter["users"] == 5
    assert starter["integrations"] == 1
    assert starter["branches"] == 5
    assert starter["benchmarking"] is False
    assert starter["board_packs"] is False
    assert growth["users"] == 25
    assert growth["integrations"] == 5
    assert growth["branches"] == 25
    assert growth["benchmarking"] is True
    assert growth["board_packs"] is True


def test_trial_exposes_product_value_without_enterprise_governance():
    trial = entitlements_for_plan("trial")
    assert trial["ai_cfo"] is True
    assert trial["forecasting"] is True
    assert trial["decision_simulator"] is True
    assert trial["advanced_governance"] is False


def test_entitlement_overrides_cannot_create_unknown_capabilities():
    value = entitlements_for_plan("founding", {"users": 12, "made_up_feature": True})
    assert value["users"] == 12
    assert "made_up_feature" not in value
