from app.services.billing.certification import provider_for_country, razorpay_mode, stripe_mode


def test_provider_routing_is_market_specific():
    assert provider_for_country("IN") == "razorpay"
    assert provider_for_country("AU") == "stripe"
    assert provider_for_country("GB") == "stripe"


def test_provider_modes_detect_test_and_live_keys():
    assert stripe_mode("sk_test_example") == "test"
    assert stripe_mode("sk_live_example") == "live"
    assert stripe_mode("dummy") == "unknown"
    assert razorpay_mode("rzp_test_example") == "test"
    assert razorpay_mode("rzp_live_example") == "live"


def test_unknown_credentials_never_look_certified():
    assert stripe_mode("placeholder") == "unknown"
    assert razorpay_mode("") == "unknown"
