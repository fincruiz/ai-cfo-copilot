from datetime import datetime
from decimal import Decimal
from uuid import UUID

from pydantic import BaseModel, ConfigDict, Field, HttpUrl


class CreateCompanyRequest(BaseModel):
    legal_name: str = Field(
        min_length=2,
        max_length=255,
    )

    trading_name: str | None = Field(
        default=None,
        max_length=255,
    )

    abn: str | None = Field(
        default=None,
        max_length=20,
    )

    country_code: str = Field(
        default="AU",
        min_length=2,
        max_length=2,
    )

    currency_code: str = Field(
        default="AUD",
        min_length=3,
        max_length=3,
    )

    financial_year_end_month: int = Field(
        default=6,
        ge=1,
        le=12,
    )

    industry: str | None = Field(
        default=None,
        max_length=255,
    )

    business_model: str | None = Field(
        default=None,
        max_length=255,
    )

    employee_count: int | None = Field(
        default=None,
        ge=0,
    )

    annual_revenue: Decimal | None = Field(
        default=None,
        ge=0,
    )

    logo_path: str | None = Field(
        default=None,
        max_length=500,
    )

    website_url: HttpUrl | None = None


class CompanyResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    legal_name: str
    trading_name: str | None = None
    abn: str | None = None

    country_code: str
    currency_code: str

    financial_year_end_month: int = Field(
        ge=1,
        le=12,
    )

    industry: str | None = None
    business_model: str | None = None
    employee_count: int | None = None
    annual_revenue: Decimal | None = None
    logo_path: str | None = None
    website_url: str | None = None
    is_active: bool
    created_by: UUID | None = None
    created_at: datetime
    updated_at: datetime

class UpdateCompanyRequest(BaseModel):
    legal_name: str | None = Field(default=None, min_length=2, max_length=255)
    trading_name: str | None = Field(default=None, max_length=255)
    abn: str | None = Field(default=None, max_length=20)
    country_code: str | None = Field(default=None, min_length=2, max_length=2)
    currency_code: str | None = Field(default=None, min_length=3, max_length=3)
    financial_year_end_month: int | None = Field(default=None, ge=1, le=12)
    industry: str | None = Field(default=None, max_length=255)
    business_model: str | None = Field(default=None, max_length=255)
    employee_count: int | None = Field(default=None, ge=0)
    annual_revenue: Decimal | None = Field(default=None, ge=0)
    website_url: HttpUrl | None = None


class CompanyPreferencesRequest(BaseModel):
    theme_preference: str = "system"
    number_format: str = "international"
    reporting_frequency: str = "monthly"
    default_report_view: str = "consolidated"
    show_ai_assistant: bool = True
    email_notifications: bool = True
    variance_warning_percent: Decimal = Field(default=10, ge=0, le=100)


class CompanyPreferencesResponse(CompanyPreferencesRequest):
    company_id: UUID
