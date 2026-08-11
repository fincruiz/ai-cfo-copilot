from dataclasses import dataclass, field
from decimal import Decimal


ZERO = Decimal("0")


@dataclass(frozen=True)
class AccountBalance:
    account_code: str
    account_name: str | None
    reporting_group: str
    reporting_subgroup: str | None
    debit: Decimal = ZERO
    credit: Decimal = ZERO
    signed_amount: Decimal = ZERO


@dataclass(frozen=True)
class ReportLine:
    code: str
    label: str
    amount: Decimal
    order: int
    is_total: bool = False
    children: tuple["ReportLine", ...] = field(
        default_factory=tuple
    )


@dataclass(frozen=True)
class ProfitAndLossReport:
    revenue: Decimal
    cost_of_sales: Decimal
    gross_profit: Decimal
    operating_expenses: Decimal
    operating_profit: Decimal
    depreciation: Decimal
    ebit: Decimal
    other_income: Decimal
    other_expenses: Decimal
    finance_costs: Decimal
    profit_before_tax: Decimal
    tax: Decimal
    net_profit: Decimal
    lines: tuple[ReportLine, ...]


@dataclass(frozen=True)
class BalanceSheetReport:
    current_assets: Decimal
    non_current_assets: Decimal
    total_assets: Decimal
    current_liabilities: Decimal
    non_current_liabilities: Decimal
    total_liabilities: Decimal
    contributed_equity: Decimal
    current_period_earnings: Decimal
    equity: Decimal
    total_liabilities_and_equity: Decimal
    balance_difference: Decimal
    lines: tuple[ReportLine, ...]
