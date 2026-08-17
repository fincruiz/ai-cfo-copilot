from app.services.market_service import market_payload, resolve_market


def test_core_markets_have_local_currency_and_terminology():
    india = market_payload("IN")
    australia = market_payload("AU")
    uae = market_payload("AE")
    uk = market_payload("GB")

    assert india["currency_code"] == "INR"
    assert india["tax_label"] == "GST"
    assert india["number_format"] == "indian"
    assert australia["currency_code"] == "AUD"
    assert australia["tax_return_label"] == "BAS"
    assert uae["currency_code"] == "AED"
    assert uae["tax_label"] == "VAT"
    assert uk["currency_code"] == "GBP"
    assert uk["registration_label"] == "Company number"


def test_pricing_is_market_specific_not_fx_generated():
    india = market_payload("IN")
    australia = market_payload("AU")
    assert india["pricing"][0]["monthly_amount_minor"] == 799900
    assert australia["pricing"][0]["monthly_amount_minor"] == 14900
    assert india["pricing"][0]["currency_code"] == "INR"
    assert australia["pricing"][0]["currency_code"] == "AUD"


def test_unknown_country_uses_global_profile():
    profile = resolve_market("ZZ")
    assert profile.market_code == "GLOBAL"
    assert profile.currency_code == "USD"
