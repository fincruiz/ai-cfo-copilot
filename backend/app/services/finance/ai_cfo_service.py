from uuid import UUID

from sqlalchemy.ext.asyncio import AsyncSession

from app.services.finance.analytics_service import AnalyticsService
from app.services.finance.reporting_service import ReportingService
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository


class AICFOService:
    def __init__(self, session: AsyncSession) -> None:
        self.analytics = AnalyticsService(session)
        self.reporting = ReportingService(GLTransactionRepository(session))

    async def answer(self, company_id: UUID, question: str) -> dict:
        q = question.lower()
        overview = await self.analytics.overview(company_id)
        pnl = await self.reporting.pnl(company_id)
        bs = await self.reporting.balance_sheet(company_id)
        ar = overview.get("ar_summary")
        ap = overview.get("ap_summary")

        if any(word in q for word in ("receivable", "customer", "collection", "ar ")):
            if not ar:
                answer = (
                    "AR ageing has not been uploaded yet. Upload a customer invoice ageing "
                    "report with Party Name and Outstanding Amount to activate collection analysis."
                )
            else:
                answer = (
                    f"Total receivables are {ar.total_outstanding:,.2f}. "
                    f"Overdue receivables are {ar.overdue_amount:,.2f}, "
                    f"or {ar.overdue_percent:.1f}% of AR. "
                    f"The largest current exposure is {ar.top_parties[0].party_name} "
                    f"at {ar.top_parties[0].outstanding_amount:,.2f}."
                )
        elif any(word in q for word in ("payable", "vendor", "supplier", "payment", "ap ")):
            if not ap:
                answer = (
                    "AP ageing has not been uploaded yet. Upload a supplier invoice ageing "
                    "report to analyse payment timing, overdue obligations and vendor exposure."
                )
            else:
                answer = (
                    f"Total payables are {ap.total_outstanding:,.2f}. "
                    f"Overdue payables are {ap.overdue_amount:,.2f}, "
                    f"or {ap.overdue_percent:.1f}% of AP. "
                    f"The largest current vendor exposure is {ap.top_parties[0].party_name} "
                    f"at {ap.top_parties[0].outstanding_amount:,.2f}."
                )
        elif any(word in q for word in ("profit", "margin", "revenue", "p&l", "pnl")):
            answer = (
                f"Revenue is {pnl.revenue:,.2f}, gross profit is {pnl.gross_profit:,.2f}, "
                f"and net profit is {pnl.net_profit:,.2f}. "
                "Use the Analytics page to review monthly movement and branch drivers."
            )
        elif any(word in q for word in ("balance", "cash", "liability", "asset")):
            answer = (
                f"Total assets are {bs.total_assets:,.2f}, liabilities are "
                f"{bs.total_liabilities:,.2f}, and equity is {bs.equity:,.2f}. "
                f"The balance-sheet difference is {bs.balance_difference:,.2f}."
            )
        elif any(word in q for word in ("mapping", "coa", "chart of accounts")):
            answer = (
                "Upload a COA file from the Import Centre. Confirmed mappings are saved "
                "against the company and reused for future GL uploads."
            )
        else:
            answer = (
                "I can help with revenue, margin, branch performance, balance-sheet health, "
                "AR collections, AP payments, mapping and forecasting. "
                + " ".join(overview.get("insights", [])[:2])
            )

        return {
            "answer": answer,
            "mode": "grounded_finance_assistant",
            "suggested_questions": [
                "What is the biggest AR collection risk?",
                "Which vendors have the largest exposure?",
                "How is monthly revenue trending?",
                "Is the balance sheet balanced?",
            ],
        }
