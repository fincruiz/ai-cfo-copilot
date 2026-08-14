from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal
from typing import Any
from uuid import UUID

from sqlalchemy import select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.ai_cfo_service import AICFOService
from app.services.finance.analytics_service import AnalyticsService
from app.services.finance.assurance_service import FinancialAssuranceService
from app.services.finance.reporting_service import ReportingService


class BrainService:
    """Deterministic management-intelligence layer.

    This service prepares management-facing facts from the finance engine and the
    integration store. AI can explain these facts elsewhere, but the metrics and
    priority rules in this overview do not depend on a language model.
    """

    def __init__(self, session: AsyncSession):
        self.session = session
        self.reporting = ReportingService(GLTransactionRepository(session))
        self.analytics = AnalyticsService(session)
        self.ai = AICFOService(session)
        self.assurance = FinancialAssuranceService(self.reporting)

    @staticmethod
    def _number(value: Any) -> float:
        if value is None:
            return 0.0
        if isinstance(value, Decimal):
            return float(value)
        try:
            return float(value)
        except (TypeError, ValueError):
            return 0.0

    @staticmethod
    def _pct_change(current: Any, previous: Any) -> float | None:
        c = BrainService._number(current)
        p = BrainService._number(previous)
        if abs(p) < 0.000001:
            return None
        return ((c - p) / abs(p)) * 100

    @staticmethod
    def _iso(value: Any) -> Any:
        if value is None:
            return None
        if hasattr(value, "isoformat"):
            return value.isoformat()
        return value

    async def _integration_context(self, company_id: UUID) -> tuple[list[dict], list[dict]]:
        try:
            con = (
                await self.session.execute(
                    text(
                        """
                        SELECT provider,status,external_tenant_name,last_synced_at,
                               last_sync_status,last_sync_message
                        FROM public.integration_connections
                        WHERE company_id=:c
                        ORDER BY provider
                        """
                    ),
                    {"c": company_id},
                )
            ).mappings().all()
            counts = (
                await self.session.execute(
                    text(
                        """
                        SELECT provider,entity_type,count(*)::int AS count
                        FROM public.integration_records
                        WHERE company_id=:c
                        GROUP BY provider,entity_type
                        ORDER BY provider,entity_type
                        """
                    ),
                    {"c": company_id},
                )
            ).mappings().all()
            return [dict(x) for x in con], [dict(x) for x in counts]
        except Exception:
            # Integration tables are optional for older workspaces/migrations. The
            # Intelligence Center should still load from finance data.
            await self.session.rollback()
            return [], []

    async def _memories(self, company_id: UUID) -> list[dict]:
        try:
            rows = (
                await self.session.execute(
                    text(
                        """
                        SELECT id,title,content,memory_type,importance,created_at
                        FROM public.organizational_memory
                        WHERE company_id=:c AND is_active=true
                        ORDER BY created_at DESC
                        LIMIT 20
                        """
                    ),
                    {"c": company_id},
                )
            ).mappings().all()
            return [dict(x) for x in rows]
        except Exception:
            await self.session.rollback()
            return []

    async def overview(self, company_id: UUID):
        company = await self.session.scalar(select(Company).where(Company.id == company_id))
        connections, source_counts = await self._integration_context(company_id)
        memories = await self._memories(company_id)

        monthly = await self.reporting.monthly_actuals(company_id)
        analytics = await self.analytics.overview(company_id)
        signals_result = await self.ai.proactive_signals(company_id)
        assurance = await self.assurance.assess(company_id)

        latest = monthly[-1] if monthly else {}
        previous = monthly[-2] if len(monthly) >= 2 else {}
        latest_revenue = self._number(latest.get("revenue"))
        latest_gp = self._number(latest.get("gross_profit"))
        latest_profit = self._number(latest.get("net_profit"))
        prev_revenue = self._number(previous.get("revenue"))
        prev_profit = self._number(previous.get("net_profit"))
        latest_margin = (latest_gp / latest_revenue * 100) if latest_revenue else None
        previous_margin = None
        prev_gp = self._number(previous.get("gross_profit"))
        if prev_revenue:
            previous_margin = prev_gp / prev_revenue * 100

        revenue_change = self._pct_change(latest_revenue, prev_revenue)
        profit_change = self._pct_change(latest_profit, prev_profit)
        margin_change = (
            latest_margin - previous_margin
            if latest_margin is not None and previous_margin is not None
            else None
        )

        ar = analytics.get("ar_summary")
        ap = analytics.get("ap_summary")
        ar_total = self._number(getattr(ar, "total_outstanding", 0) if ar else 0)
        ar_overdue = self._number(getattr(ar, "overdue_amount", 0) if ar else 0)
        ar_overdue_pct = self._number(getattr(ar, "overdue_percent", 0) if ar else 0)
        ap_total = self._number(getattr(ap, "total_outstanding", 0) if ap else 0)
        ap_overdue = self._number(getattr(ap, "overdue_amount", 0) if ap else 0)

        snapshot = [
            {
                "key": "revenue",
                "label": "Revenue",
                "value": latest_revenue,
                "format": "currency",
                "change": revenue_change,
                "change_unit": "percent",
                "context": "Latest reporting month",
            },
            {
                "key": "gross_margin",
                "label": "Gross margin",
                "value": latest_margin,
                "format": "percent",
                "change": margin_change,
                "change_unit": "points",
                "context": "Profit left after direct costs",
            },
            {
                "key": "net_profit",
                "label": "Net profit",
                "value": latest_profit,
                "format": "currency",
                "change": profit_change,
                "change_unit": "percent",
                "context": "Latest reporting month",
            },
            {
                "key": "overdue_ar",
                "label": "Overdue receivables",
                "value": ar_overdue,
                "format": "currency",
                "change": ar_overdue_pct if ar else None,
                "change_unit": "of_ar",
                "context": f"{ar_overdue_pct:.1f}% of receivables overdue" if ar else "Upload AR ageing to activate",
            },
            {
                "key": "financial_confidence",
                "label": "Financial confidence",
                "value": assurance.get("score", 0),
                "format": "score",
                "change": None,
                "change_unit": None,
                "context": f"Grade {assurance.get('grade', '—')} · structural checks",
            },
        ]

        trends = []
        for row in monthly[-12:]:
            revenue = self._number(row.get("revenue"))
            gp = self._number(row.get("gross_profit"))
            trends.append(
                {
                    "month": self._iso(row.get("month")),
                    "revenue": revenue,
                    "gross_profit": gp,
                    "net_profit": self._number(row.get("net_profit")),
                    "gross_margin": (gp / revenue * 100) if revenue else None,
                }
            )

        priorities: list[dict] = []
        for signal in signals_result.get("signals", []):
            severity = str(signal.get("severity", "low"))
            level = {
                "high": "critical",
                "medium": "attention",
                "positive": "positive",
                "low": "monitor",
            }.get(severity, "attention")
            priorities.append(
                {
                    "level": level,
                    "title": signal.get("title"),
                    "evidence": signal.get("evidence"),
                    "action": signal.get("action"),
                    "source": "Finance engine",
                }
            )

        if ar and ar_overdue_pct >= 20:
            priorities.append(
                {
                    "level": "critical" if ar_overdue_pct >= 40 else "attention",
                    "title": f"{ar_overdue_pct:.1f}% of receivables are overdue",
                    "evidence": f"{ar_overdue:,.0f} overdue out of {ar_total:,.0f} total receivables.",
                    "action": "Prioritise the largest overdue customer balances and confirm collection dates.",
                    "source": "AR ageing",
                }
            )

        if ap and ap_total > ar_total and ar_total > 0:
            priorities.append(
                {
                    "level": "attention",
                    "title": "Payables currently exceed receivables",
                    "evidence": f"Payables are {ap_total:,.0f} versus receivables of {ar_total:,.0f}.",
                    "action": "Review near-term payment commitments against collections and cash availability.",
                    "source": "Working capital",
                }
            )

        if assurance.get("score", 0) < 90:
            failed = [
                c.get("label")
                for c in assurance.get("checks", [])
                if c.get("status") != "pass"
            ]
            priorities.append(
                {
                    "level": "attention",
                    "title": "Review data quality before relying on all insights",
                    "evidence": ", ".join(failed[:3]) if failed else "Financial confidence is below the ready threshold.",
                    "action": "Resolve the Financial Confidence checks before making high-impact decisions from the reports.",
                    "source": "Financial assurance",
                }
            )

        order = {"critical": 0, "attention": 1, "positive": 2, "monitor": 3}
        priorities = sorted(priorities, key=lambda x: order.get(x["level"], 9))[:6]

        critical_count = sum(1 for p in priorities if p["level"] == "critical")
        attention_count = sum(1 for p in priorities if p["level"] == "attention")
        positive_count = sum(1 for p in priorities if p["level"] == "positive")

        if critical_count:
            headline = f"{critical_count} issue{'s' if critical_count != 1 else ''} need management attention now"
        elif attention_count:
            headline = f"{attention_count} area{'s' if attention_count != 1 else ''} deserve management attention"
        elif positive_count:
            headline = "The latest signals are broadly constructive"
        elif monthly:
            headline = "No major exception is currently above FinCruiz thresholds"
        else:
            headline = "Connect or upload business data to activate management intelligence"

        summary_bits = []
        if revenue_change is not None:
            summary_bits.append(f"Revenue moved {revenue_change:+.1f}% month on month")
        if margin_change is not None:
            summary_bits.append(f"gross margin moved {margin_change:+.1f} pts")
        if ar:
            summary_bits.append(f"{ar_overdue_pct:.1f}% of receivables are overdue")
        narrative = ". ".join(summary_bits[:3]) + ("." if summary_bits else "")
        if not narrative:
            narrative = "FinCruiz will explain the most important movements once sufficient finance or connected-system data is available."

        freshness = []
        for con in connections:
            freshness.append(
                {
                    "provider": con.get("provider"),
                    "name": con.get("external_tenant_name") or str(con.get("provider", "")).title(),
                    "status": con.get("status"),
                    "last_synced_at": self._iso(con.get("last_synced_at")),
                    "last_sync_status": con.get("last_sync_status"),
                    "last_sync_message": con.get("last_sync_message"),
                }
            )

        return {
            "company": {
                "name": (company.trading_name or company.legal_name) if company else "Company",
                "currency": company.currency_code if company else "AUD",
                "industry": company.industry if company else None,
                "business_model": company.business_model if company else None,
            },
            "executive_summary": {
                "headline": headline,
                "narrative": narrative,
                "critical_count": critical_count,
                "attention_count": attention_count,
                "positive_count": positive_count,
                "generated_at": datetime.now(timezone.utc).isoformat(),
            },
            "financial_snapshot": snapshot,
            "monthly_trends": trends,
            "priorities": priorities,
            "connections": connections,
            "source_counts": source_counts,
            "source_freshness": freshness,
            "memories": memories,
            "signals": signals_result.get("signals", []),
            "financial_assurance": assurance,
            "working_capital": {
                "receivables": {
                    "total": ar_total,
                    "overdue": ar_overdue,
                    "overdue_percent": ar_overdue_pct,
                } if ar else None,
                "payables": {
                    "total": ap_total,
                    "overdue": ap_overdue,
                } if ap else None,
            },
            "suggested_questions": [
                "What should management focus on this month?",
                "Why did profit move differently from revenue?",
                "Where is cash getting tied up?",
                "What are the biggest risks over the next 90 days?",
            ],
        }

    async def add_memory(self, company_id: UUID, user_id: UUID, payload):
        row = (
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.organizational_memory(
                        company_id,memory_type,title,content,importance,created_by
                    )
                    VALUES(:c,:t,:title,:content,:i,:u)
                    RETURNING id,title,content,memory_type,importance,created_at
                    """
                ),
                {
                    "c": company_id,
                    "t": payload.memory_type,
                    "title": payload.title,
                    "content": payload.content,
                    "i": payload.importance,
                    "u": user_id,
                },
            )
        ).mappings().first()
        await self.session.commit()
        return dict(row)
