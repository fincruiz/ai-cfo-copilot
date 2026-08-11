from __future__ import annotations

from decimal import Decimal
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.schemas.finance.imports import (
    AgeingBucketResponse,
    PartyExposureResponse,
    WorkingCapitalSummaryResponse,
)
from app.services.finance.reporting_service import ReportingService


class AnalyticsService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session
        self.reporting = ReportingService(GLTransactionRepository(session))

    async def working_capital_summary(
        self,
        company_id: UUID,
        ageing_type: str,
    ) -> WorkingCapitalSummaryResponse | None:
        totals = (
            await self.session.execute(
                text(
                    """
                    SELECT
                        COALESCE(SUM(outstanding_amount), 0) AS total_outstanding,
                        COALESCE(SUM(CASE WHEN COALESCE(days_overdue, 0) > 0
                            THEN outstanding_amount ELSE 0 END), 0) AS overdue_amount,
                        COUNT(*) AS document_count,
                        COUNT(DISTINCT party_name) AS party_count,
                        CASE WHEN SUM(ABS(outstanding_amount)) = 0 THEN NULL
                             ELSE SUM(
                                outstanding_amount * GREATEST(COALESCE(days_overdue, 0), 0)
                             ) / SUM(ABS(outstanding_amount))
                        END AS weighted_days
                    FROM public.finance_ageing_documents
                    WHERE company_id = :company_id AND ageing_type = :ageing_type
                    """
                ),
                {"company_id": company_id, "ageing_type": ageing_type},
            )
        ).mappings().one()

        if int(totals["document_count"] or 0) == 0:
            return None

        bucket_rows = (
            await self.session.execute(
                text(
                    """
                    SELECT age_bucket AS bucket,
                           COALESCE(SUM(outstanding_amount), 0) AS amount,
                           COUNT(*) AS document_count
                    FROM public.finance_ageing_documents
                    WHERE company_id = :company_id AND ageing_type = :ageing_type
                    GROUP BY age_bucket
                    ORDER BY CASE age_bucket
                        WHEN 'Current' THEN 1
                        WHEN '1-30' THEN 2
                        WHEN '31-60' THEN 3
                        WHEN '61-90' THEN 4
                        WHEN '90+' THEN 5
                        ELSE 6 END
                    """
                ),
                {"company_id": company_id, "ageing_type": ageing_type},
            )
        ).mappings().all()

        party_rows = (
            await self.session.execute(
                text(
                    """
                    SELECT
                        party_name,
                        COALESCE(SUM(outstanding_amount), 0) AS outstanding_amount,
                        COALESCE(SUM(CASE WHEN COALESCE(days_overdue, 0) > 0
                            THEN outstanding_amount ELSE 0 END), 0) AS overdue_amount,
                        COUNT(*) AS document_count,
                        MIN(due_date) AS oldest_due_date,
                        CASE WHEN SUM(ABS(outstanding_amount)) = 0 THEN NULL
                             ELSE SUM(
                                outstanding_amount * GREATEST(COALESCE(days_overdue, 0), 0)
                             ) / SUM(ABS(outstanding_amount))
                        END AS weighted_days_overdue
                    FROM public.finance_ageing_documents
                    WHERE company_id = :company_id AND ageing_type = :ageing_type
                    GROUP BY party_name
                    ORDER BY outstanding_amount DESC
                    LIMIT 15
                    """
                ),
                {"company_id": company_id, "ageing_type": ageing_type},
            )
        ).mappings().all()

        total = Decimal(totals["total_outstanding"] or 0)
        overdue = Decimal(totals["overdue_amount"] or 0)
        overdue_percent = (
            overdue / total * Decimal("100")
            if total
            else Decimal("0")
        )

        return WorkingCapitalSummaryResponse(
            ageing_type=ageing_type,
            total_outstanding=total,
            overdue_amount=overdue,
            overdue_percent=overdue_percent,
            current_amount=total - overdue,
            document_count=int(totals["document_count"] or 0),
            party_count=int(totals["party_count"] or 0),
            weighted_average_days_overdue=(
                Decimal(totals["weighted_days"])
                if totals["weighted_days"] is not None
                else None
            ),
            buckets=[
                AgeingBucketResponse(
                    bucket=str(row["bucket"]),
                    amount=Decimal(row["amount"] or 0),
                    document_count=int(row["document_count"] or 0),
                )
                for row in bucket_rows
            ],
            top_parties=[
                PartyExposureResponse(
                    party_name=str(row["party_name"]),
                    outstanding_amount=Decimal(row["outstanding_amount"] or 0),
                    overdue_amount=Decimal(row["overdue_amount"] or 0),
                    document_count=int(row["document_count"] or 0),
                    oldest_due_date=row["oldest_due_date"],
                    weighted_days_overdue=(
                        Decimal(row["weighted_days_overdue"])
                        if row["weighted_days_overdue"] is not None
                        else None
                    ),
                )
                for row in party_rows
            ],
        )

    async def overview(self, company_id: UUID) -> dict:
        monthly = await self.reporting.monthly_actuals(company_id)
        branch = await self.reporting.branch_comparison(company_id)
        ar = await self.working_capital_summary(company_id, "AR")
        ap = await self.working_capital_summary(company_id, "AP")

        insights: list[str] = []
        if ar:
            if ar.overdue_percent >= Decimal("40"):
                insights.append(
                    f"Receivables risk is elevated: {ar.overdue_percent:.1f}% of AR is overdue."
                )
            elif ar.overdue_percent >= Decimal("20"):
                insights.append(
                    f"Receivables need attention: {ar.overdue_percent:.1f}% of AR is overdue."
                )
            else:
                insights.append(
                    f"Receivables ageing is controlled: overdue AR is {ar.overdue_percent:.1f}%."
                )

        if ap and ar:
            if ap.total_outstanding > ar.total_outstanding:
                insights.append(
                    "Outstanding payables exceed receivables; review near-term cash commitments."
                )
            else:
                insights.append(
                    "Outstanding receivables exceed payables, but collection timing remains important."
                )

        if len(monthly) >= 2:
            current = Decimal(str(monthly[-1]["revenue"]))
            prior = Decimal(str(monthly[-2]["revenue"]))
            if prior:
                change = (current - prior) / abs(prior) * Decimal("100")
                insights.append(f"Latest monthly revenue changed by {change:.1f}% versus the prior month.")

        if not insights:
            insights.append(
                "Upload GL, COA and AR/AP ageing files to activate deeper analytics."
            )

        return {
            "monthly_actuals": monthly,
            "branch_comparison": branch,
            "ar_summary": ar,
            "ap_summary": ap,
            "insights": insights,
        }
