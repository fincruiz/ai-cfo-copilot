from datetime import date
from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

import pandas as pd
import pytest

from app.domain.finance.advanced_forecasting import (
    ForecastConfig,
    ForecastDrivers,
    HistoricalData,
    OpeningBalanceSheet,
    TrendBudgetForecastBuilder,
    ThreeWayForecastEngine,
)
from app.domain.finance.mapping.classifier import suggest_mapping
from app.domain.finance.reporting.adapters import rows_to_account_balances
from app.schemas.finance.advanced_forecasting import PowerOfOneRequest
from app.services.finance.advanced_forecasting_service import AdvancedForecastingService
from app.services.finance.reporting_service import ReportingService


def test_mapping_keeps_payroll_inside_opex_and_accumulated_dep_as_contra_asset():
    payroll = suggest_mapping("6100", "Salaries and wages")
    assert payroll.reporting_group == "Operating Expenses"
    assert payroll.reporting_subgroup == "Payroll / People"
    assert payroll.sign_convention == "debit"

    accum = suggest_mapping("1599", "Accumulated Depreciation - Equipment")
    assert accum.reporting_group == "Non Current Assets"
    assert accum.reporting_subgroup == "Accumulated Depreciation"
    assert accum.sign_convention == "credit"


def test_generic_positive_sign_does_not_hide_reversals_or_contra_balances():
    rows = [
        SimpleNamespace(
            source_account_code="4000",
            account_name="Revenue reversal",
            reporting_group="Revenue",
            reporting_subgroup="Sales",
            sign_convention="positive",
            debit=Decimal("150"),
            credit=Decimal("100"),
        ),
        SimpleNamespace(
            source_account_code="1599",
            account_name="Accumulated Depreciation",
            reporting_group="Non Current Assets",
            reporting_subgroup="Accumulated Depreciation",
            sign_convention="positive",
            debit=Decimal("0"),
            credit=Decimal("80"),
        ),
    ]
    balances = rows_to_account_balances(rows)
    assert balances[0].signed_amount == Decimal("-50")
    assert balances[1].signed_amount == Decimal("80")


@pytest.mark.asyncio
async def test_default_reporting_period_uses_company_financial_year_and_latest_actual():
    class Repo:
        session = None

        async def latest_transaction_date(self, *, company_id, branch_id=None):
            return date(2025, 12, 15)

    service = ReportingService(Repo())

    async def fy_month(_company_id):
        return 6

    service._financial_year_end_month = fy_month
    start, end = await service._resolve_income_period(uuid4(), None, None, None)
    assert start == date(2025, 7, 1)
    assert end == date(2025, 12, 15)


@pytest.mark.asyncio
async def test_default_reporting_period_handles_calendar_year_company():
    class Repo:
        session = None

        async def latest_transaction_date(self, *, company_id, branch_id=None):
            return date(2026, 8, 20)

    service = ReportingService(Repo())

    async def fy_month(_company_id):
        return 12

    service._financial_year_end_month = fy_month
    start, end = await service._resolve_income_period(uuid4(), None, None, None)
    assert start == date(2026, 1, 1)
    assert end == date(2026, 8, 20)


def test_existing_opening_ppe_continues_depreciating_without_new_capex():
    forecast = pd.DataFrame({
        "Period": ["2026-07-01", "2026-08-01"],
        "Revenue": [100_000.0, 100_000.0],
        "COGS": [50_000.0, 50_000.0],
        "Payroll": [20_000.0, 20_000.0],
        "Other Opex": [10_000.0, 10_000.0],
    })
    opening = OpeningBalanceSheet(
        cash=100_000,
        accounts_receivable=0,
        inventory=0,
        other_current_assets=0,
        gross_ppe=600_000,
        accumulated_depreciation=-120_000,
        other_non_current_assets=0,
        accounts_payable=0,
        accrued_expenses=0,
        other_current_liabilities=0,
        debt_current=0,
        debt_non_current=0,
        other_non_current_liabilities=0,
        share_capital=0,
        retained_earnings=None,
    )
    drivers = ForecastDrivers(
        capex_pct_revenue=0,
        useful_life_months=60,
        tax_rate=0,
        minimum_cash=0,
        revolver_limit=0,
    )
    result = ThreeWayForecastEngine(
        forecast,
        ForecastConfig(opening_balance_sheet=opening, drivers=drivers),
    ).run()
    assert float(result.profit_and_loss.iloc[0]["Depreciation"]) > 0
    assert float(result.balance_sheet.iloc[0]["Accumulated Depreciation"]) < -120_000
    assert bool(result.checks["Balanced"].all())


@pytest.mark.asyncio
async def test_power_of_one_operating_levers_change_operating_forecast():
    service = AdvancedForecastingService(None)
    history = pd.DataFrame({
        "Period": pd.date_range("2025-01-01", periods=12, freq="MS"),
        "Revenue": [100_000.0] * 12,
        "COGS": [60_000.0] * 12,
        "Payroll": [20_000.0] * 12,
        "Other Opex": [10_000.0] * 12,
    })

    async def history_stub(company_id, branch_id):
        return history.copy()

    async def budget_stub(company_id, version_id):
        return None

    async def calculate_stub(company_id, request, persist=False):
        config = service._config(request)
        build = TrendBudgetForecastBuilder(HistoricalData(history.copy()), None, config).build()
        result = ThreeWayForecastEngine(build.forecast, config).run()
        return {"summary": {
            "forecast_revenue": float(result.profit_and_loss["Revenue"].sum()),
            "forecast_ebitda": float(result.profit_and_loss["EBITDA"].sum()),
            "forecast_net_income": float(result.profit_and_loss["Net Income"].sum()),
            "closing_cash": float(result.balance_sheet["Cash"].iloc[-1]),
            "closing_debt": float(result.balance_sheet["Current Debt"].iloc[-1] + result.balance_sheet["Non-current Debt"].iloc[-1]),
        }}

    service._history = history_stub
    service._budget = budget_stub
    service.calculate = calculate_stub

    common = dict(
        forecast_start=date(2026, 1, 1),
        forecast_months=12,
        opening_balance_sheet={
            "cash": 500_000,
            "accounts_receivable": 100_000,
            "inventory": 50_000,
            "other_current_assets": 0,
            "gross_ppe": 0,
            "accumulated_depreciation": 0,
            "other_non_current_assets": 0,
            "accounts_payable": 80_000,
            "accrued_expenses": 0,
            "other_current_liabilities": 0,
            "debt_current": 0,
            "debt_non_current": 0,
            "other_non_current_liabilities": 0,
            "share_capital": 0,
            "retained_earnings": None,
        },
    )
    margin = await service.power_of_one(uuid4(), PowerOfOneRequest(**common, gross_margin_points=Decimal("1")))
    payroll = await service.power_of_one(uuid4(), PowerOfOneRequest(**common, payroll_change_percent=Decimal("10")))
    opex = await service.power_of_one(uuid4(), PowerOfOneRequest(**common, other_opex_change_percent=Decimal("10")))

    assert margin["adjusted"]["forecast_ebitda"] > margin["base"]["forecast_ebitda"]
    assert payroll["adjusted"]["forecast_ebitda"] < payroll["base"]["forecast_ebitda"]
    assert opex["adjusted"]["forecast_ebitda"] < opex["base"]["forecast_ebitda"]


def test_three_way_ui_requires_actual_baseline_and_exposes_assumptions():
    text = Path("../frontend/app/dashboard/three-way-forecast/page.tsx").read_text(encoding="utf-8")
    assert "planningBaseline" in text
    assert 'forecast_start: ""' in text
    assert "2027-01-01" not in text
    assert "Actual-data forecast baseline" in text
    assert "Opening Balance Sheet" in text
    assert "Payroll % of revenue" in text
    assert "Capex % of revenue" in text
    assert "PercentField" in text


def test_dashboard_hides_legacy_floating_ai_trigger():
    text = Path("../frontend/components/ai-cfo-floating.tsx").read_text(encoding="utf-8")
    assert 'pathname !== "/dashboard"' in text
    contextual = Path("../frontend/components/contextual-ai-bar.tsx").read_text(encoding="utf-8")
    assert 'if (pathname === "/dashboard") return null;' in contextual
