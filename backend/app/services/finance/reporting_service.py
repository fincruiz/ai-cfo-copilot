from collections import defaultdict
from datetime import date
from decimal import Decimal

from sqlalchemy import select, text
from uuid import UUID

from app.domain.finance.kpis.ratio_engine import calculate_ratios
from app.domain.finance.reporting.adapters import rows_to_account_balances
from app.domain.finance.reporting.balance_sheet import build_balance_sheet
from app.domain.finance.reporting.pnl import build_profit_and_loss
from app.domain.finance.reporting.rules import canonical_reporting_group
from app.domain.finance.reporting.trial_balance import build_trial_balance
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.database.models.core.company import Company


class ReportingService:
    def __init__(self, repository: GLTransactionRepository):
        self.repository = repository

    async def _financial_year_end_month(self, company_id: UUID) -> int:
        value = (
            await self.repository.session.execute(
                select(Company.financial_year_end_month).where(Company.id == company_id)
            )
        ).scalar_one_or_none()
        return int(value or 6)

    async def _resolve_income_period(
        self,
        company_id: UUID,
        start_date: date | None,
        end_date: date | None,
        branch_id: UUID | None,
    ) -> tuple[date | None, date | None]:
        resolved_end = end_date or await self.repository.latest_transaction_date(
            company_id=company_id, branch_id=branch_id
        )
        if resolved_end is None:
            return start_date, end_date
        if start_date is not None:
            return start_date, resolved_end

        fy_end_month = await self._financial_year_end_month(company_id)
        start_month = fy_end_month % 12 + 1
        if start_month == 1:
            start_year = resolved_end.year
        else:
            start_year = resolved_end.year if resolved_end.month >= start_month else resolved_end.year - 1
        return date(start_year, start_month, 1), resolved_end

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
        start_date, end_date = await self._resolve_income_period(
            company_id, start_date, end_date, branch_id
        )
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
        resolved_end = end_date or await self.repository.latest_transaction_date(
            company_id=company_id, branch_id=branch_id
        )
        rows = await self.repository.account_balances(
            company_id=company_id,
            end_date=resolved_end,
            statement="balance_sheet",
            branch_id=branch_id,
        )
        pnl_start, pnl_end = await self._resolve_income_period(
            company_id, None, resolved_end, branch_id
        )
        pnl = await self.pnl(
            company_id,
            start_date=pnl_start,
            end_date=pnl_end,
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
        start_date, end_date = await self._resolve_income_period(
            company_id, start_date, end_date, branch_id
        )
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
            debit = Decimal(row.debit or 0)
            credit = Decimal(row.credit or 0)
            asset_amount = debit - credit
            liability_amount = credit - debit
            if "inventory" in label or "stock" in label:
                inv += asset_amount
            if "cash" in label or "bank" in label:
                cash += asset_amount
            if "receivable" in label or "debtor" in label:
                recv += asset_amount
            if "payable" in label or "creditor" in label:
                pay += liability_amount
            if "loan" in label or "borrow" in label or "debt" in label:
                debt += liability_amount
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


    async def data_health(self, company_id: UUID) -> dict:
        session = self.repository.session
        summary = (
            await session.execute(
                text(
                    """
                    SELECT
                      COUNT(gt.id)::int AS transaction_count,
                      COUNT(DISTINCT gt.source_account_code)::int AS account_count,
                      MIN(gt.transaction_date) AS first_transaction_date,
                      MAX(gt.transaction_date) AS last_transaction_date,
                      COALESCE(SUM(gt.debit), 0) AS total_debit,
                      COALESCE(SUM(gt.credit), 0) AS total_credit,
                      COUNT(*) FILTER (WHERE gt.validation_status <> 'valid')::int AS invalid_transaction_count
                    FROM public.gl_transactions gt
                    JOIN public.file_uploads fu ON fu.id = gt.file_upload_id
                    WHERE gt.company_id=:company_id
                      AND fu.company_id=:company_id
                      AND fu.is_active=true
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        upload_counts = (
            await session.execute(
                text(
                    """
                    SELECT COUNT(*)::int AS upload_count,
                           COUNT(*) FILTER (WHERE is_active=true)::int AS active_upload_count
                    FROM public.file_uploads WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        mapping_counts = (
            await session.execute(
                text(
                    """
                    WITH accounts AS (
                      SELECT DISTINCT gt.source_account_code
                      FROM public.gl_transactions gt
                      JOIN public.file_uploads fu ON fu.id=gt.file_upload_id
                      WHERE gt.company_id=:company_id AND fu.is_active=true
                    )
                    SELECT
                      COUNT(*) FILTER (WHERE fam.source_account_code IS NOT NULL)::int AS mapped_account_count,
                      COUNT(*) FILTER (WHERE fam.source_account_code IS NULL)::int AS unmapped_account_count
                    FROM accounts a
                    LEFT JOIN public.finance_account_mappings fam
                      ON fam.company_id=:company_id
                     AND fam.source_account_code=a.source_account_code
                     AND fam.is_confirmed=true
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        duplicate_candidates = int(
            (
                await session.execute(
                    text(
                        """
                        SELECT COALESCE(SUM(group_count - 1), 0)::int
                        FROM (
                          SELECT COUNT(*) AS group_count
                          FROM public.gl_transactions gt
                          JOIN public.file_uploads fu ON fu.id=gt.file_upload_id
                          WHERE gt.company_id=:company_id AND fu.is_active=true
                          GROUP BY gt.transaction_date, gt.source_account_code, gt.debit, gt.credit,
                                   COALESCE(gt.document_number,''), COALESCE(gt.description,'')
                          HAVING COUNT(*) > 1
                        ) duplicates
                        """
                    ),
                    {"company_id": company_id},
                )
            ).scalar_one()
            or 0
        )

        bs = await self.balance_sheet(company_id)
        total_debit = Decimal(summary["total_debit"] or 0)
        total_credit = Decimal(summary["total_credit"] or 0)
        tb_difference = total_debit - total_credit
        bs_difference = Decimal(bs.balance_difference or 0)
        tolerance = Decimal("0.01")
        tb_balanced = abs(tb_difference) <= tolerance
        bs_balanced = abs(bs_difference) <= tolerance
        mapping_complete = int(mapping_counts["unmapped_account_count"] or 0) == 0
        invalid_count = int(summary["invalid_transaction_count"] or 0)

        if not int(summary["transaction_count"] or 0):
            overall = "empty"
        elif tb_balanced and bs_balanced and mapping_complete and invalid_count == 0:
            overall = "healthy"
        else:
            overall = "attention_required"

        return {
            "transaction_count": int(summary["transaction_count"] or 0),
            "upload_count": int(upload_counts["upload_count"] or 0),
            "active_upload_count": int(upload_counts["active_upload_count"] or 0),
            "account_count": int(summary["account_count"] or 0),
            "mapped_account_count": int(mapping_counts["mapped_account_count"] or 0),
            "unmapped_account_count": int(mapping_counts["unmapped_account_count"] or 0),
            "invalid_transaction_count": invalid_count,
            "duplicate_candidate_count": duplicate_candidates,
            "first_transaction_date": summary["first_transaction_date"],
            "last_transaction_date": summary["last_transaction_date"],
            "total_debit": total_debit,
            "total_credit": total_credit,
            "trial_balance_difference": tb_difference,
            "balance_sheet_difference": bs_difference,
            "is_trial_balance_balanced": tb_balanced,
            "is_balance_sheet_balanced": bs_balanced,
            "is_mapping_complete": mapping_complete,
            "overall_status": overall,
        }

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
