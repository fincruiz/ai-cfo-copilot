"""P9 canonical financial truth regression tests.

These fixtures intentionally use exact Decimal values. They protect the deterministic
finance layer from silent changes while the dashboard and AI experience evolve.
"""
from decimal import Decimal

from app.domain.finance.kpis import calculate_ratios
from app.domain.finance.reporting import AccountBalance, build_balance_sheet, build_profit_and_loss, build_trial_balance


def D(value: str) -> Decimal:
    return Decimal(value)


def test_service_company_financial_truth():
    income = [
        AccountBalance("4000", "Consulting Revenue", "Revenue", "Sales", credit=D("250000.00"), signed_amount=D("250000.00")),
        AccountBalance("6100", "Salaries", "Operating Expenses", "Payroll", debit=D("90000.00"), signed_amount=D("90000.00")),
        AccountBalance("6200", "Rent", "Operating Expenses", "Occupancy", debit=D("20000.00"), signed_amount=D("20000.00")),
    ]
    pnl = build_profit_and_loss(income)
    assert pnl.revenue == D("250000.00")
    assert pnl.gross_profit == D("250000.00")
    assert pnl.operating_expenses == D("110000.00")
    assert pnl.net_profit == D("140000.00")

    balance = [
        AccountBalance("1000", "Bank", "Current Assets", "Cash", debit=D("180000.00"), signed_amount=D("180000.00")),
        AccountBalance("1100", "Receivables", "Current Assets", "Trade Receivables", debit=D("60000.00"), signed_amount=D("60000.00")),
        AccountBalance("2000", "Payables", "Current Liabilities", "Trade Payables", credit=D("40000.00"), signed_amount=D("40000.00")),
        AccountBalance("3000", "Capital", "Equity", "Capital", credit=D("60000.00"), signed_amount=D("60000.00")),
    ]
    bs = build_balance_sheet(balance, current_period_earnings=pnl.net_profit)
    assert bs.current_assets == D("240000.00")
    assert bs.current_liabilities == D("40000.00")
    assert bs.equity == D("200000.00")
    assert bs.balance_difference == D("0.00")

    ratios = calculate_ratios(pnl, bs, cash=D("180000.00"), receivables=D("60000.00"), payables=D("40000.00"))
    gross_margin = next(r for r in ratios if r.name == "Gross Margin")
    assert gross_margin.value == D("100")


def test_inventory_business_financial_truth():
    income = [
        AccountBalance("4000", "Product Sales", "Revenue", "Sales", credit=D("500000.00"), signed_amount=D("500000.00")),
        AccountBalance("5000", "COGS", "Cost of Sales", "COGS", debit=D("300000.00"), signed_amount=D("300000.00")),
        AccountBalance("6100", "Marketing", "Operating Expenses", "Marketing", debit=D("50000.00"), signed_amount=D("50000.00")),
    ]
    pnl = build_profit_and_loss(income)
    assert pnl.gross_profit == D("200000.00")
    assert pnl.net_profit == D("150000.00")

    balance = [
        AccountBalance("1000", "Bank", "Current Assets", "Cash", debit=D("100000.00"), signed_amount=D("100000.00")),
        AccountBalance("1100", "Receivables", "Current Assets", "Trade Receivables", debit=D("120000.00"), signed_amount=D("120000.00")),
        AccountBalance("1200", "Inventory", "Current Assets", "Inventory", debit=D("180000.00"), signed_amount=D("180000.00")),
        AccountBalance("2000", "Payables", "Current Liabilities", "Trade Payables", credit=D("150000.00"), signed_amount=D("150000.00")),
        AccountBalance("3000", "Capital", "Equity", "Capital", credit=D("100000.00"), signed_amount=D("100000.00")),
    ]
    bs = build_balance_sheet(balance, current_period_earnings=pnl.net_profit)
    assert bs.current_assets == D("400000.00")
    assert bs.current_liabilities == D("150000.00")
    assert bs.equity == D("250000.00")
    assert bs.balance_difference == D("0.00")


def test_trial_balance_preserves_cent_level_precision():
    tb = build_trial_balance([
        AccountBalance("1000", "Bank", "Current Assets", "Cash", debit=D("100.01"), signed_amount=D("100.01")),
        AccountBalance("3000", "Equity", "Equity", "Capital", credit=D("100.01"), signed_amount=D("100.01")),
    ])
    assert tb.total_debit == D("100.01")
    assert tb.total_credit == D("100.01")
    assert tb.difference == D("0.00")
