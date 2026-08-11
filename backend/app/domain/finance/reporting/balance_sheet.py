from collections import defaultdict
from decimal import Decimal
from typing import Iterable

from app.domain.finance.reporting.models import (
    AccountBalance,
    BalanceSheetReport,
    ReportLine,
)
from app.domain.finance.reporting.rules import (
    canonical_reporting_group,
    reporting_group_order,
)

ZERO = Decimal("0")


def build_balance_sheet(
    account_balances: Iterable[AccountBalance],
    *,
    current_period_earnings: Decimal = ZERO,
) -> BalanceSheetReport:
    totals: dict[str, Decimal] = defaultdict(lambda: ZERO)
    details: list[ReportLine] = []

    for account in account_balances:
        group = canonical_reporting_group(account.reporting_group)
        if group is None:
            continue

        totals[group] += account.signed_amount
        details.append(
            ReportLine(
                account.account_code,
                account.account_name or account.account_code,
                account.signed_amount,
                reporting_group_order(group),
            )
        )

    current_assets = totals["Current Assets"]
    non_current_assets = totals["Non Current Assets"]
    current_liabilities = totals["Current Liabilities"]
    non_current_liabilities = totals["Non Current Liabilities"]
    contributed_equity = totals["Equity"]

    total_assets = current_assets + non_current_assets
    total_liabilities = current_liabilities + non_current_liabilities

    # Current-period profit belongs in equity until it is closed to retained
    # earnings. Including it here makes the live management balance sheet
    # reconcile without mutating the source ledger.
    equity = contributed_equity + current_period_earnings
    total_liabilities_and_equity = total_liabilities + equity
    balance_difference = total_assets - total_liabilities_and_equity

    lines = details + [
        ReportLine("CURRENT_ASSETS", "Current Assets", current_assets, 21, True),
        ReportLine("NON_CURRENT_ASSETS", "Non Current Assets", non_current_assets, 22, True),
        ReportLine("TOTAL_ASSETS", "Total Assets", total_assets, 23, True),
        ReportLine("CURRENT_LIABILITIES", "Current Liabilities", current_liabilities, 31, True),
        ReportLine("NON_CURRENT_LIABILITIES", "Non Current Liabilities", non_current_liabilities, 32, True),
        ReportLine("TOTAL_LIABILITIES", "Total Liabilities", total_liabilities, 33, True),
        ReportLine("CONTRIBUTED_EQUITY", "Contributed Equity / Retained Earnings", contributed_equity, 39, True),
        ReportLine("CURRENT_PERIOD_EARNINGS", "Current Period Earnings", current_period_earnings, 40, True),
        ReportLine("EQUITY", "Total Equity", equity, 41, True),
        ReportLine(
            "TOTAL_LIABILITIES_EQUITY",
            "Total Liabilities and Equity",
            total_liabilities_and_equity,
            42,
            True,
        ),
    ]

    return BalanceSheetReport(
        current_assets=current_assets,
        non_current_assets=non_current_assets,
        total_assets=total_assets,
        current_liabilities=current_liabilities,
        non_current_liabilities=non_current_liabilities,
        total_liabilities=total_liabilities,
        contributed_equity=contributed_equity,
        current_period_earnings=current_period_earnings,
        equity=equity,
        total_liabilities_and_equity=total_liabilities_and_equity,
        balance_difference=balance_difference,
        lines=tuple(
            sorted(
                lines,
                key=lambda line: (
                    line.order,
                    line.is_total,
                    line.label,
                ),
            )
        ),
    )
