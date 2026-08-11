from collections import defaultdict
from decimal import Decimal
from typing import Iterable

from app.domain.finance.reporting.models import (
    AccountBalance,
    ProfitAndLossReport,
    ReportLine,
)
from app.domain.finance.reporting.rules import (
    canonical_reporting_group,
    reporting_group_order,
)


ZERO = Decimal("0")


def _sum_group(
    totals: dict[str, Decimal],
    group_name: str,
) -> Decimal:
    return totals.get(group_name, ZERO)


def build_profit_and_loss(
    account_balances: Iterable[AccountBalance],
) -> ProfitAndLossReport:
    grouped_totals: dict[str, Decimal] = defaultdict(
        lambda: ZERO
    )

    detail_lines: list[ReportLine] = []

    for account in account_balances:
        group = canonical_reporting_group(
            account.reporting_group
        )

        if group is None:
            continue

        grouped_totals[group] += account.signed_amount

        detail_lines.append(
            ReportLine(
                code=account.account_code,
                label=(
                    account.account_name
                    or account.account_code
                ),
                amount=account.signed_amount,
                order=reporting_group_order(group),
            )
        )

    revenue = _sum_group(
        grouped_totals,
        "Revenue",
    )

    cost_of_sales = _sum_group(
        grouped_totals,
        "Cost of Sales",
    )

    gross_profit = revenue - cost_of_sales

    operating_expenses = _sum_group(
        grouped_totals,
        "Operating Expenses",
    )

    operating_profit = (
        gross_profit - operating_expenses
    )

    depreciation = _sum_group(
        grouped_totals,
        "Depreciation",
    )

    ebit = operating_profit - depreciation

    other_income = _sum_group(
        grouped_totals,
        "Other Income",
    )

    other_expenses = _sum_group(
        grouped_totals,
        "Other Expenses",
    )

    finance_costs = _sum_group(
        grouped_totals,
        "Finance Costs",
    )

    profit_before_tax = (
        ebit
        + other_income
        - other_expenses
        - finance_costs
    )

    tax = _sum_group(
        grouped_totals,
        "Tax",
    )

    net_profit = profit_before_tax - tax

    total_lines = [
        ReportLine(
            code="REVENUE",
            label="Revenue",
            amount=revenue,
            order=1,
            is_total=True,
        ),
        ReportLine(
            code="COST_OF_SALES",
            label="Cost of Sales",
            amount=cost_of_sales,
            order=2,
            is_total=True,
        ),
        ReportLine(
            code="GROSS_PROFIT",
            label="Gross Profit",
            amount=gross_profit,
            order=3,
            is_total=True,
        ),
        ReportLine(
            code="OPERATING_EXPENSES",
            label="Operating Expenses",
            amount=operating_expenses,
            order=4,
            is_total=True,
        ),
        ReportLine(
            code="OPERATING_PROFIT",
            label="Operating Profit",
            amount=operating_profit,
            order=5,
            is_total=True,
        ),
        ReportLine(
            code="DEPRECIATION",
            label="Depreciation",
            amount=depreciation,
            order=7,
            is_total=True,
        ),
        ReportLine(
            code="EBIT",
            label="EBIT",
            amount=ebit,
            order=8,
            is_total=True,
        ),
        ReportLine(
            code="PROFIT_BEFORE_TAX",
            label="Profit Before Tax",
            amount=profit_before_tax,
            order=11,
            is_total=True,
        ),
        ReportLine(
            code="TAX",
            label="Tax",
            amount=tax,
            order=12,
            is_total=True,
        ),
        ReportLine(
            code="NET_PROFIT",
            label="Net Profit",
            amount=net_profit,
            order=13,
            is_total=True,
        ),
    ]

    lines = tuple(
        sorted(
            [*detail_lines, *total_lines],
            key=lambda line: (
                line.order,
                line.is_total,
                line.label,
            ),
        )
    )

    return ProfitAndLossReport(
        revenue=revenue,
        cost_of_sales=cost_of_sales,
        gross_profit=gross_profit,
        operating_expenses=operating_expenses,
        operating_profit=operating_profit,
        depreciation=depreciation,
        ebit=ebit,
        other_income=other_income,
        other_expenses=other_expenses,
        finance_costs=finance_costs,
        profit_before_tax=profit_before_tax,
        tax=tax,
        net_profit=net_profit,
        lines=lines,
    )