from datetime import datetime,timezone,timedelta
from app.services.subscription_service import entitlements_for_plan,days_remaining
from app.services.market_service import PRICE_CATALOG

def test_plan_entitlements_are_separate_from_market_price():
    assert entitlements_for_plan('growth')['integrations']==5
    assert PRICE_CATALOG['IN']['growth']['monthly'] != PRICE_CATALOG['AU']['growth']['monthly']

def test_trial_days_never_go_negative():
    assert days_remaining(datetime.now(timezone.utc)-timedelta(days=2))==0

def test_trial_and_growth_expose_decision_simulator():
    assert entitlements_for_plan('trial')['decision_simulator'] is True
    assert entitlements_for_plan('growth')['decision_simulator'] is True
