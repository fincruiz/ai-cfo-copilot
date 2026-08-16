from __future__ import annotations

from decimal import Decimal
import pandas as pd
import pytest

from app.security.company_roles import can_company_admin, can_finance_write
from app.domain.finance.forecasting.engine import build_run_rate_forecast, build_trend_forecast
from app.domain.finance.advanced_forecasting.models import ForecastConfig, ForecastDrivers, OpeningBalanceSheet
from app.domain.finance.advanced_forecasting.three_way import ThreeWayForecastEngine


def D(value: str) -> Decimal:
    return Decimal(value)


def test_monthly_run_rate_forecast_uses_recent_average_and_bounds():
    history = [
        ("2026-01", D("100.00")),
        ("2026-02", D("120.00")),
        ("2026-03", D("140.00")),
        ("2026-04", D("160.00")),
    ]
    points = build_run_rate_forecast(history, ["2026-05", "2026-06"], recent_months=3)
    assert [p.base for p in points] == [D("140.00"), D("140.00")]
    assert points[0].downside == D("126.0000")
    assert points[0].upside == D("154.0000")


def test_monthly_trend_forecast_preserves_linear_direction():
    history = [
        ("2026-01", D("100")),
        ("2026-02", D("120")),
        ("2026-03", D("140")),
    ]
    points = build_trend_forecast(history, ["2026-04", "2026-05"])
    assert points[0].base == D("160")
    assert points[1].base == D("180")


def test_three_way_forecast_balances_every_period_and_cash_rolls_forward():
    forecast = pd.DataFrame(
        {
            "Period": ["2026-07-31", "2026-08-31", "2026-09-30"],
            "Revenue": [300_000.0, 315_000.0, 330_000.0],
            "COGS": [180_000.0, 189_000.0, 198_000.0],
            "Payroll": [60_000.0, 60_000.0, 62_000.0],
            "Other Opex": [30_000.0, 31_000.0, 32_000.0],
        }
    )
    opening = OpeningBalanceSheet(
        cash=250_000.0,
        accounts_receivable=140_000.0,
        inventory=100_000.0,
        other_current_assets=20_000.0,
        gross_ppe=300_000.0,
        accumulated_depreciation=-80_000.0,
        other_non_current_assets=10_000.0,
        accounts_payable=110_000.0,
        accrued_expenses=30_000.0,
        other_current_liabilities=15_000.0,
        debt_current=20_000.0,
        debt_non_current=120_000.0,
        other_non_current_liabilities=5_000.0,
        share_capital=150_000.0,
        retained_earnings=None,
    )
    drivers = ForecastDrivers(
        tax_rate=0.25,
        dso_days=40.0,
        dpo_days=35.0,
        inventory_days=45.0,
        capex_pct_revenue=0.02,
        scheduled_debt_repayment=5_000.0,
        minimum_cash=75_000.0,
    )
    result = ThreeWayForecastEngine(
        forecast,
        ForecastConfig(opening_balance_sheet=opening, drivers=drivers),
    ).run()

    assert len(result.balance_sheet) == 3
    assert result.balance_sheet["Balance Check"].abs().max() < 0.01

    opening_cash = 250_000.0
    expected_first_close = opening_cash + float(result.cash_flow.iloc[0]["Net Change in Cash"])
    assert abs(float(result.cash_flow.iloc[0]["Cash"]) - expected_first_close) < 0.01

    for i in range(1, len(result.cash_flow)):
        previous_close = float(result.cash_flow.iloc[i - 1]["Cash"])
        movement = float(result.cash_flow.iloc[i]["Net Change in Cash"])
        close = float(result.cash_flow.iloc[i]["Cash"])
        assert abs(close - (previous_close + movement)) < 0.01


@pytest.mark.parametrize("role", ["owner", "admin", "cfo", "finance_manager", "accountant"])
def test_finance_write_role_matrix_allows_finance_roles(role: str):
    assert can_finance_write(role)


@pytest.mark.parametrize("role", ["board_member", "viewer"])
def test_finance_write_role_matrix_denies_read_only_roles(role: str):
    assert not can_finance_write(role)


@pytest.mark.parametrize("role", ["owner", "admin"])
def test_company_admin_role_matrix_allows_admin_roles(role: str):
    assert can_company_admin(role)


@pytest.mark.parametrize("role", ["cfo", "finance_manager", "accountant", "board_member", "viewer"])
def test_company_admin_role_matrix_denies_non_admin_roles(role: str):
    assert not can_company_admin(role)


def test_multi_branch_consolidation_matches_sum_of_branch_truth():
    from app.domain.finance.reporting import AccountBalance, build_profit_and_loss

    branch_a = [
        AccountBalance("4000", "Revenue", "Revenue", "Sales", credit=D("150000.00"), signed_amount=D("150000.00")),
        AccountBalance("5000", "COGS", "Cost of Sales", "COGS", debit=D("90000.00"), signed_amount=D("90000.00")),
        AccountBalance("6100", "Payroll", "Operating Expenses", "Payroll", debit=D("25000.00"), signed_amount=D("25000.00")),
    ]
    branch_b = [
        AccountBalance("4000", "Revenue", "Revenue", "Sales", credit=D("100000.00"), signed_amount=D("100000.00")),
        AccountBalance("5000", "COGS", "Cost of Sales", "COGS", debit=D("55000.00"), signed_amount=D("55000.00")),
        AccountBalance("6100", "Payroll", "Operating Expenses", "Payroll", debit=D("20000.00"), signed_amount=D("20000.00")),
    ]
    a = build_profit_and_loss(branch_a)
    b = build_profit_and_loss(branch_b)
    consolidated = build_profit_and_loss(branch_a + branch_b)

    assert consolidated.revenue == a.revenue + b.revenue == D("250000.00")
    assert consolidated.cost_of_sales == a.cost_of_sales + b.cost_of_sales == D("145000.00")
    assert consolidated.gross_profit == a.gross_profit + b.gross_profit == D("105000.00")
    assert consolidated.operating_expenses == a.operating_expenses + b.operating_expenses == D("45000.00")
    assert consolidated.net_profit == a.net_profit + b.net_profit == D("60000.00")
