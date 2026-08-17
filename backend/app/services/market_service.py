from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class MarketProfile:
    market_code: str
    country_code: str
    country_name: str
    currency_code: str
    locale_code: str
    registration_label: str
    tax_label: str
    tax_return_label: str
    financial_year_label: str
    default_fye_month: int
    number_format: str


MARKETS: dict[str, MarketProfile] = {
    "IN": MarketProfile("IN", "IN", "India", "INR", "en-IN", "GSTIN / CIN", "GST", "GST return", "Financial year", 3, "indian"),
    "AU": MarketProfile("AU", "AU", "Australia", "AUD", "en-AU", "ABN", "GST", "BAS", "Financial year", 6, "international"),
    "AE": MarketProfile("AE", "AE", "United Arab Emirates", "AED", "en-AE", "TRN / trade licence", "VAT", "VAT return", "Financial year", 12, "international"),
    "GB": MarketProfile("GB", "GB", "United Kingdom", "GBP", "en-GB", "Company number", "VAT", "VAT return", "Financial year", 12, "international"),
    "US": MarketProfile("US", "US", "United States", "USD", "en-US", "EIN / registration", "Sales tax", "Tax return", "Fiscal year", 12, "international"),
}

GLOBAL_MARKET = MarketProfile("GLOBAL", "GLOBAL", "International", "USD", "en-US", "Registration number", "Tax", "Tax return", "Financial year", 12, "international")

# Launch pricing is intentionally configured separately by market instead of FX conversion.
# amount_minor is stored in the smallest currency unit and can later be replaced by billing-provider price IDs.
PRICE_CATALOG: dict[str, dict[str, dict[str, int | None]]] = {
    "IN": {"founding": {"monthly": 799900, "annual": 7999000}, "growth": {"monthly": 1799900, "annual": 17999000}, "enterprise": {"monthly": None, "annual": None}},
    "AU": {"founding": {"monthly": 14900, "annual": 149000}, "growth": {"monthly": 29900, "annual": 299000}, "enterprise": {"monthly": None, "annual": None}},
    "AE": {"founding": {"monthly": 39900, "annual": 399000}, "growth": {"monthly": 79900, "annual": 799000}, "enterprise": {"monthly": None, "annual": None}},
    "GB": {"founding": {"monthly": 7900, "annual": 79000}, "growth": {"monthly": 15900, "annual": 159000}, "enterprise": {"monthly": None, "annual": None}},
    "US": {"founding": {"monthly": 9900, "annual": 99000}, "growth": {"monthly": 19900, "annual": 199000}, "enterprise": {"monthly": None, "annual": None}},
    "GLOBAL": {"founding": {"monthly": 9900, "annual": 99000}, "growth": {"monthly": 19900, "annual": 199000}, "enterprise": {"monthly": None, "annual": None}},
}

PLAN_LABELS = {"founding": "Essentials", "growth": "Growth", "enterprise": "Enterprise"}


def resolve_market(country_code: str | None) -> MarketProfile:
    code = (country_code or "").strip().upper()
    return MARKETS.get(code, GLOBAL_MARKET)


def market_payload(country_code: str | None) -> dict:
    market = resolve_market(country_code)
    pricing = PRICE_CATALOG.get(market.market_code, PRICE_CATALOG["GLOBAL"])
    return {
        "market_code": market.market_code,
        "country_code": market.country_code,
        "country_name": market.country_name,
        "currency_code": market.currency_code,
        "locale_code": market.locale_code,
        "registration_label": market.registration_label,
        "tax_label": market.tax_label,
        "tax_return_label": market.tax_return_label,
        "financial_year_label": market.financial_year_label,
        "default_fye_month": market.default_fye_month,
        "number_format": market.number_format,
        "pricing": [
            {
                "plan": plan,
                "display_name": PLAN_LABELS[plan],
                "currency_code": market.currency_code,
                "monthly_amount_minor": values["monthly"],
                "annual_amount_minor": values["annual"],
                "contact_sales": values["monthly"] is None,
            }
            for plan, values in pricing.items()
        ],
    }
