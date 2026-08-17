from pydantic import BaseModel


class RegionalPlanPrice(BaseModel):
    plan: str
    display_name: str
    currency_code: str
    monthly_amount_minor: int | None = None
    annual_amount_minor: int | None = None
    contact_sales: bool = False


class MarketProfileOut(BaseModel):
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
    pricing: list[RegionalPlanPrice]
