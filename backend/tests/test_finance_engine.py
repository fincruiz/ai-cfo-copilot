from decimal import Decimal

from app.domain.finance.forecasting import build_run_rate_forecast
from app.domain.finance.kpis import calculate_ratios
from app.domain.finance.mapping import suggest_mapping
from app.domain.finance.reporting import AccountBalance, build_balance_sheet, build_profit_and_loss, build_trial_balance


def test_finance_reports_and_ratios():
    pnl = build_profit_and_loss([
        AccountBalance("4000", "Sales", "Revenue", "Sales", credit=Decimal("10000"), signed_amount=Decimal("10000")),
        AccountBalance("5000", "COGS", "Cost of Sales", None, debit=Decimal("4000"), signed_amount=Decimal("4000")),
        AccountBalance("6000", "Rent", "Operating Expenses", None, debit=Decimal("1000"), signed_amount=Decimal("1000")),
    ])
    assert pnl.gross_profit == Decimal("6000")
    assert pnl.net_profit == Decimal("5000")
    bs = build_balance_sheet([
        AccountBalance("1000", "Bank", "Current Assets", "Cash", debit=Decimal("5000"), signed_amount=Decimal("5000")),
        AccountBalance("2000", "Payables", "Current Liabilities", "Trade Payables", credit=Decimal("2000"), signed_amount=Decimal("2000")),
        AccountBalance("3000", "Equity", "Equity", None, credit=Decimal("3000"), signed_amount=Decimal("3000")),
    ])
    assert bs.balance_difference == Decimal("0")
    ratios = calculate_ratios(pnl, bs, cash=Decimal("5000"), payables=Decimal("2000"))
    assert any(r.name == "Gross Margin" for r in ratios)


def test_trial_balance_mapping_and_forecast():
    tb = build_trial_balance([
        AccountBalance("1000", "Bank", "Current Assets", None, debit=Decimal("100"), signed_amount=Decimal("100")),
        AccountBalance("3000", "Capital", "Equity", None, credit=Decimal("100"), signed_amount=Decimal("100")),
    ])
    assert tb.difference == 0
    assert suggest_mapping("1000", "Main Bank").reporting_group == "Current Assets"
    points = build_run_rate_forecast([("2026-01", Decimal("100")), ("2026-02", Decimal("200"))], ["2026-03"])
    assert points[0].base == Decimal("150")


def test_balance_sheet_includes_current_period_earnings():
    from decimal import Decimal
    from app.domain.finance.reporting.balance_sheet import build_balance_sheet
    from app.domain.finance.reporting.models import AccountBalance

    accounts = [
        AccountBalance("1000", "Bank", "Current Assets", "Cash", debit=Decimal("150"), credit=Decimal("0"), signed_amount=Decimal("150")),
        AccountBalance("2000", "Payables", "Current Liabilities", "Trade Payables", debit=Decimal("0"), credit=Decimal("50"), signed_amount=Decimal("50")),
        AccountBalance("3000", "Capital", "Equity", "Capital", debit=Decimal("0"), credit=Decimal("80"), signed_amount=Decimal("80")),
    ]

    report = build_balance_sheet(accounts, current_period_earnings=Decimal("20"))

    assert report.contributed_equity == Decimal("80")
    assert report.current_period_earnings == Decimal("20")
    assert report.equity == Decimal("100")
    assert report.balance_difference == Decimal("0")
