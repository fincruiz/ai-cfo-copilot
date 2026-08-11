from collections import defaultdict
from datetime import date
from decimal import Decimal
from uuid import UUID

from app.domain.finance.kpis.ratio_engine import calculate_ratios
from app.domain.finance.reporting.adapters import rows_to_account_balances
from app.domain.finance.reporting.balance_sheet import build_balance_sheet
from app.domain.finance.reporting.pnl import build_profit_and_loss
from app.domain.finance.reporting.rules import canonical_reporting_group
from app.domain.finance.reporting.trial_balance import build_trial_balance
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository


class ReportingService:
    def __init__(self, repository: GLTransactionRepository):
        self.repository = repository

    async def trial_balance(
        self,
        company_id: UUID,
        start_date: date | None = None,
        end_date: date | None = None,
        branch_id: UUID | None = None,
    ):
        rows = await self.repository.account_balances(
            company_id=company_id,
            start_date=start_date,
            end_date=end_date,
            branch_id=branch_id,
        )
        return build_trial_balance(
            rows_to_account_balances(rows, include_unmapped=True)
        )

    async def pnl(
        self,
        company_id: UUID,
        start_date: date | None = None,
        end_date: date | None = None,
        branch_id: UUID | None = None,
    ):
        rows = await self.repository.account_balances(
            company_id=company_id,
            start_date=start_date,
            end_date=end_date,
            statement="income_statement",
            branch_id=branch_id,
        )
        return build_profit_and_loss(rows_to_account_balances(rows))

    async def balance_sheet(
        self,
        company_id: UUID,
        end_date: date | None = None,
        branch_id: UUID | None = None,
    ):
        rows = await self.repository.account_balances(
            company_id=company_id,
            end_date=end_date,
            statement="balance_sheet",
            branch_id=branch_id,
        )
        pnl = await self.pnl(
            company_id,
            end_date=end_date,
            branch_id=branch_id,
        )
        return build_balance_sheet(
            rows_to_account_balances(rows),
            current_period_earnings=pnl.net_profit,
        )

    async def kpis(
        self,
        company_id: UUID,
        start_date: date | None = None,
        end_date: date | None = None,
        period_days=365,
        employees=None,
        branch_id: UUID | None = None,
    ):
        pnl = await self.pnl(company_id, start_date, end_date, branch_id)
        bs = await self.balance_sheet(company_id, end_date, branch_id)
        rows = await self.repository.account_balances(
            company_id=company_id,
            end_date=end_date,
            statement="balance_sheet",
            branch_id=branch_id,
        )
        inv = cash = recv = pay = debt = Decimal("0")
        for row in rows:
            label = (
                f"{row.reporting_group or ''} "
                f"{row.reporting_subgroup or ''} "
                f"{row.account_name or ''}"
            ).lower()
            amount = abs(Decimal(row.debit or 0) - Decimal(row.credit or 0))
            if "inventory" in label or "stock" in label:
                inv += amount
            if "cash" in label or "bank" in label:
                cash += amount
            if "receivable" in label or "debtor" in label:
                recv += amount
            if "payable" in label or "creditor" in label:
                pay += amount
            if "loan" in label or "borrow" in label or "debt" in label:
                debt += amount
        return calculate_ratios(
            pnl,
            bs,
            inventory=inv,
            cash=cash,
            receivables=recv,
            payables=pay,
            debt=debt,
            period_days=period_days,
            employees=employees,
        )

    async def monthly_actuals(
        self,
        company_id: UUID,
        *,
        branch_id: UUID | None = None,
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[dict]:
        rows = await self.repository.monthly_actuals(
            company_id=company_id,
            branch_id=branch_id,
            start_date=start_date,
            end_date=end_date,
        )
        by_month: dict[date, dict[str, Decimal]] = defaultdict(
            lambda: defaultdict(lambda: Decimal("0"))
        )
        for row in rows:
            month = row.month.date() if hasattr(row.month, "date") else row.month
            group = canonical_reporting_group(row.reporting_group)
            if group:
                by_month[month][group] += Decimal(row.amount or 0)

        result = []
        for month in sorted(by_month):
            totals = by_month[month]
            revenue = totals["Revenue"]
            cost = totals["Cost of Sales"]
            gross = revenue - cost
            opex = totals["Operating Expenses"]
            depreciation = totals["Depreciation"]
            ebit = gross - opex - depreciation
            finance_costs = totals["Finance Costs"]
            tax = totals["Tax"]
            net_profit = (
                ebit
                + totals["Other Income"]
                - totals["Other Expenses"]
                - finance_costs
                - tax
            )
            result.append(
                {
                    "month": month,
                    "revenue": revenue,
                    "cost_of_sales": cost,
                    "gross_profit": gross,
                    "operating_expenses": opex,
                    "depreciation": depreciation,
                    "ebit": ebit,
                    "finance_costs": finance_costs,
                    "tax": tax,
                    "net_profit": net_profit,
                }
            )
        return result

    async def branch_comparison(
        self,
        company_id: UUID,
        start_date: date | None = None,
        end_date: date | None = None,
    ) -> list[dict]:
        result = []
        for branch_id, branch_code, branch_name in (
            await self.repository.branch_ids_with_activity(company_id)
        ):
            report = await self.pnl(
                company_id,
                start_date,
                end_date,
                branch_id,
            )
            gross_margin = (
                report.gross_profit / report.revenue * Decimal("100")
                if report.revenue
                else None
            )
            net_margin = (
                report.net_profit / report.revenue * Decimal("100")
                if report.revenue
                else None
            )
            result.append(
                {
                    "branch_id": branch_id,
                    "branch_code": branch_code,
                    "branch_name": branch_name,
                    "revenue": report.revenue,
                    "gross_profit": report.gross_profit,
                    "operating_expenses": report.operating_expenses,
                    "ebit": report.ebit,
                    "net_profit": report.net_profit,
                    "gross_margin_percent": gross_margin,
                    "net_margin_percent": net_margin,
                }
            )
        return result
