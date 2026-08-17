from datetime import date
from decimal import Decimal
from typing import Any
from uuid import UUID
from pydantic import BaseModel, Field

class ForecastDriversInput(BaseModel):
    gross_margin: Decimal = Decimal('0.42')
    payroll_pct_revenue: Decimal = Decimal('0.22')
    other_opex_pct_revenue: Decimal = Decimal('0.14')
    annual_interest_rate: Decimal = Decimal('0.075')
    tax_rate: Decimal = Decimal('0.30')
    dso_days: Decimal = Decimal('42')
    dpo_days: Decimal = Decimal('35')
    inventory_days: Decimal = Decimal('45')
    capex_pct_revenue: Decimal = Decimal('0.04')
    useful_life_months: int = 60
    scheduled_debt_repayment: Decimal = Decimal('10000')
    minimum_cash: Decimal = Decimal('100000')
    revolver_limit: Decimal = Decimal('500000')
    dividend_pct_net_income: Decimal = Decimal('0')

class OpeningBalanceSheetInput(BaseModel):
    cash: Decimal = Decimal('250000')
    accounts_receivable: Decimal = Decimal('300000')
    inventory: Decimal = Decimal('180000')
    other_current_assets: Decimal = Decimal('40000')
    gross_ppe: Decimal = Decimal('600000')
    accumulated_depreciation: Decimal = Decimal('-140000')
    other_non_current_assets: Decimal = Decimal('25000')
    accounts_payable: Decimal = Decimal('210000')
    accrued_expenses: Decimal = Decimal('75000')
    other_current_liabilities: Decimal = Decimal('35000')
    debt_current: Decimal = Decimal('60000')
    debt_non_current: Decimal = Decimal('420000')
    other_non_current_liabilities: Decimal = Decimal('20000')
    share_capital: Decimal = Decimal('200000')
    retained_earnings: Decimal | None = None

class AdvancedForecastRequest(BaseModel):
    run_name: str = 'Management Forecast'
    forecast_start: date
    forecast_months: int = Field(default=24, ge=1, le=60)
    branch_id: UUID | None = None
    budget_version_id: UUID | None = None
    trend_weight: Decimal = Decimal('0.45')
    budget_weight: Decimal = Decimal('0.35')
    run_rate_weight: Decimal = Decimal('0.20')
    seasonality_enabled: bool = True
    drivers: ForecastDriversInput = Field(default_factory=ForecastDriversInput)
    opening_balance_sheet: OpeningBalanceSheetInput = Field(default_factory=OpeningBalanceSheetInput)

class PowerOfOneRequest(AdvancedForecastRequest):
    price_change_percent: Decimal = Decimal('0')
    volume_change_percent: Decimal = Decimal('0')
    gross_margin_points: Decimal = Decimal('0')
    payroll_change_percent: Decimal = Decimal('0')
    other_opex_change_percent: Decimal = Decimal('0')
    dso_change_days: Decimal = Decimal('0')
    dpo_change_days: Decimal = Decimal('0')
    inventory_change_days: Decimal = Decimal('0')
    capex_change_percent: Decimal = Decimal('0')
    interest_rate_points: Decimal = Decimal('0')

class ForecastRunResponse(BaseModel):
    run_id: UUID
    run_name: str
    summary: dict[str, Any]
    profit_and_loss: list[dict[str, Any]]
    balance_sheet: list[dict[str, Any]]
    cash_flow: list[dict[str, Any]]
    ratios: list[dict[str, Any]]
    checks: list[dict[str, Any]]
    scenarios: list[dict[str, Any]]
    diagnostics: list[dict[str, Any]]

class PowerOfOneResponse(BaseModel):
    base: dict[str, Any]
    adjusted: dict[str, Any]
    impact: dict[str, Any]

class PlanningVersionCreate(BaseModel):
    plan_type: str
    version_name: str
    financial_year_start: date
    financial_year_end: date
    # P9 Stage 7 assisted planning. Existing callers using seed_from_actuals
    # remain compatible; seed_mode is the new explicit control.
    seed_from_actuals: bool = True
    seed_mode: str = 'actuals'  # actuals | previous_budget | blank
    detail_level: str = 'high_level'  # high_level | detailed
    allocation_method: str = 'actuals_ratio'
    seed_growth_percent: Decimal = Decimal('0')
    seed_version_id: UUID | None = None
    seed_imported_version: str | None = None


class HighLevelBudgetAllocationRequest(BaseModel):
    annual_targets: dict[str, Decimal]
    detail_level: str = 'detailed'
    seasonality: str = 'historical'  # historical | equal

class NativePlanLineInput(BaseModel):
    period: date
    branch_id: UUID | None = None
    reporting_group: str
    reporting_subgroup: str | None = None
    source_account_code: str | None = None
    amount: Decimal
    driver_type: str = 'manual'
    driver_value: Decimal | None = None
    notes: str | None = None

class PlanningVersionResponse(BaseModel):
    id: UUID
    plan_type: str
    version_name: str
    financial_year_start: date
    financial_year_end: date
    status: str
    assumptions: dict[str, Any]
    lines: list[dict[str, Any]] = Field(default_factory=list)

class BoardPackGenerateRequest(BaseModel):
    pack_name: str
    reporting_period: str
    forecast_run_id: UUID | None = None
    sections: list[str] = Field(default_factory=lambda: [
        'executive_summary','financial_highlights','profit_and_loss','balance_sheet',
        'cash_flow','kpis','monthly_trends','branch_comparison','working_capital',
        'forecast','scenarios','risks_actions'
    ])
    formats: list[str] = Field(default_factory=lambda: ['pptx','pdf','xlsx'])
    management_outlook: str = ''
    strategic_priorities: str = ''
    principal_risks: str = ''
    decisions_required: str = ''

class ArtifactResponse(BaseModel):
    id: UUID
    artifact_type: str
    file_name: str
    download_url: str
    file_size_bytes: int

class DecisionSimulatorRequest(AdvancedForecastRequest):
    revenue_change_percent: Decimal = Decimal('0')
    price_change_percent: Decimal = Decimal('0')
    volume_change_percent: Decimal = Decimal('0')
    gross_margin_points: Decimal = Decimal('0')
    headcount_change: int = 0
    monthly_cost_per_hire: Decimal = Decimal('0')
    payroll_change_percent: Decimal = Decimal('0')
    other_opex_change_percent: Decimal = Decimal('0')
    dso_change_days: Decimal = Decimal('0')
    dpo_change_days: Decimal = Decimal('0')
    inventory_change_days: Decimal = Decimal('0')
    capex_change_percent: Decimal = Decimal('0')

class DecisionSimulatorResponse(BaseModel):
    scenario_name: str
    assumptions: dict[str, Any]
    base_summary: dict[str, Any]
    scenario_summary: dict[str, Any]
    impact: dict[str, Any]
    assessment: dict[str, Any]
    comparison_series: list[dict[str, Any]]
    base_checks: list[dict[str, Any]]
    scenario_checks: list[dict[str, Any]]
