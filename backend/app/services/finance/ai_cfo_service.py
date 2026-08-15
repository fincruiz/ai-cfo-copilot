from __future__ import annotations

import json
from decimal import Decimal
from typing import Any
from uuid import UUID

import httpx
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.database.models.core.company import Company
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.analytics_service import AnalyticsService
from app.services.finance.reporting_service import ReportingService
from app.services.finance.assurance_service import FinancialAssuranceService


class AICFOService:
    """Finance-first assistant.

    Deterministic platform guidance is handled locally. Financial analysis is built
    from aggregated company metrics. When external context is useful and an OpenAI
    key is configured, the service can enrich the answer with live web search.
    Raw GL rows are never included in the web prompt.
    """

    def __init__(self, session: AsyncSession) -> None:
        self.session = session
        self.analytics = AnalyticsService(session)
        self.reporting = ReportingService(GLTransactionRepository(session))
        self.assurance = FinancialAssuranceService(self.reporting)

    @staticmethod
    def _guidance(question: str) -> dict | None:
        q = question.lower()
        if any(term in q for term in ("where do i upload", "how do i upload", "load gl", "upload gl", "general ledger")):
            return {
                "answer": (
                    "Start with your General Ledger in Data Uploads. I’ll help you validate the file, "
                    "then guide you to Account Mapping before you open the reports. A CSV export with "
                    "date, account code/name and debit/credit columns is the best starting point."
                ),
                "action": {"label": "Go to Data Uploads", "route": "/dashboard/uploads"},
            }
        if any(term in q for term in ("chart of accounts", "coa", "ar ageing", "ap ageing", "receivable ageing", "payable ageing")) and any(term in q for term in ("upload", "load", "where", "import")):
            return {
                "answer": (
                    "Use the Finance Import Centre for Chart of Accounts, AR ageing and AP ageing files. "
                    "Those supporting files unlock persistent mappings, customer collection analysis and vendor exposure."
                ),
                "action": {"label": "Open Import Centre", "route": "/dashboard/import-center"},
            }
        if any(term in q for term in ("map account", "mapping", "unmapped")):
            return {
                "answer": (
                    "Open Account Mapping after the GL is uploaded. Review any unmapped accounts, confirm their "
                    "reporting group, and FinCruiz will reuse confirmed mappings on future uploads."
                ),
                "action": {"label": "Review Account Mapping", "route": "/dashboard/mapping"},
            }
        if any(term in q for term in (
            "what if", "can we afford", "can i afford", "can we hire", "hire", "new employee",
            "new staff", "increase salary", "raise salaries", "capex", "buy equipment", "new warehouse",
            "open a branch", "price increase", "change price", "reduce price", "cash impact", "cash effect",
            "scenario", "model this decision", "model the decision"
        )):
            return {
                "answer": (
                    "This is a decision-modelling question, so the best FinCruiz tool is the Integrated Three-Way Forecast. "
                    "It links Profit & Loss, Balance Sheet and Cash Flow so you can test the decision against profit, working capital, "
                    "closing cash and debt together rather than looking at one statement in isolation. Open the model, adjust the relevant "
                    "drivers, then compare the resulting cash and profitability impact."
                ),
                "action": {"label": "Model it in Three-Way Forecast", "route": "/dashboard/three-way-forecast"},
            }
        if any(term in q for term in ("reset data", "delete data", "privacy", "delete profile", "delete account", "wrong file")):
            route = "/dashboard/settings"
            label = "Open Data & Privacy"
            if any(term in q for term in ("ar", "receivable")):
                route, label = "/dashboard/import-center", "Open AR / AP Import Centre"
            elif any(term in q for term in ("branch", "business unit")):
                route, label = "/dashboard/branches", "Open Branches"
            elif any(term in q for term in ("general ledger", " gl", "ledger")):
                route, label = "/dashboard/uploads", "Open General Ledger"
            elif any(term in q for term in ("mapping", "coa", "chart of accounts")):
                route, label = "/dashboard/mapping", "Open Account Mapping"
            return {
                "answer": (
                    "You do not need to erase the whole workspace for most mistakes. FinCruiz supports module-level resets "
                    "for the General Ledger, mappings/COA, AR, AP, branches, planning, forecasts and board packs. "
                    "The confirmation dialog explains exactly what will be removed before the action runs. Reset All and permanent profile deletion remain separate controls."
                ),
                "action": {"label": label, "route": route},
            }
        return None

    @staticmethod
    def _external_context_useful(question: str) -> bool:
        q = question.lower()
        triggers = (
            "industry", "economic", "economy", "inflation", "interest rate", "rates",
            "market", "benchmark", "outlook", "macro", "external", "online", "competitor",
            "sector", "demand", "consumer", "exchange rate", "fx", "regulation", "trend",
            "risk", "opportunity", "strategy", "forecast assumption",
        )
        return any(term in q for term in triggers)

    @staticmethod
    def _jsonable(value: Any) -> Any:
        if isinstance(value, Decimal):
            return float(value)
        if hasattr(value, "isoformat"):
            return value.isoformat()
        if isinstance(value, dict):
            return {k: AICFOService._jsonable(v) for k, v in value.items()}
        if isinstance(value, (list, tuple)):
            return [AICFOService._jsonable(v) for v in value]
        if hasattr(value, "model_dump"):
            return AICFOService._jsonable(value.model_dump())
        return value

    async def _company_context(self, company_id: UUID) -> tuple[Company | None, dict]:
        company = await self.session.scalar(select(Company).where(Company.id == company_id))
        overview = await self.analytics.overview(company_id)
        pnl = await self.reporting.pnl(company_id)
        bs = await self.reporting.balance_sheet(company_id)
        kpis = await self.reporting.kpis(company_id)
        monthly = overview.get("monthly_actuals", [])[-12:]
        assurance = await self.assurance.assess(company_id)
        integration_rows = (await self.session.execute(
            __import__("sqlalchemy").text(
                "SELECT provider,status,external_tenant_name,last_synced_at FROM public.integration_connections WHERE company_id=:company_id"
            ),
            {"company_id": company_id},
        )).mappings().all()
        memory_rows = (await self.session.execute(
            __import__("sqlalchemy").text(
                "SELECT title,content,memory_type,importance FROM public.organizational_memory WHERE company_id=:company_id AND is_active=true ORDER BY created_at DESC LIMIT 20"
            ),
            {"company_id": company_id},
        )).mappings().all()
        source_counts = (await self.session.execute(
            __import__("sqlalchemy").text(
                "SELECT provider,entity_type,count(*) AS count FROM public.integration_records WHERE company_id=:company_id GROUP BY provider,entity_type"
            ),
            {"company_id": company_id},
        )).mappings().all()

        context = {
            "company": {
                "name": (company.trading_name or company.legal_name) if company else None,
                "country": company.country_code if company else None,
                "currency": company.currency_code if company else None,
                "industry": company.industry if company else None,
                "business_model": company.business_model if company else None,
                "employee_count": company.employee_count if company else None,
            },
            "pnl": self._jsonable(pnl),
            "balance_sheet": self._jsonable(bs),
            "kpis": self._jsonable(kpis),
            "monthly_actuals": self._jsonable(monthly),
            "ar_summary": self._jsonable(overview.get("ar_summary")),
            "ap_summary": self._jsonable(overview.get("ap_summary")),
            "existing_insights": overview.get("insights", []),
            "financial_assurance": self._jsonable(assurance),
            "connected_systems": self._jsonable([dict(row) for row in integration_rows]),
            "organizational_memory": self._jsonable([dict(row) for row in memory_rows]),
            "integration_source_counts": self._jsonable([dict(row) for row in source_counts]),
        }
        return company, context

    async def _openai_web_answer(self, question: str, context: dict) -> tuple[str, list[dict]] | None:
        if not settings.openai_api_key:
            return None

        prompt = f"""You are the AI CFO inside FinCruiz. Answer the management question using BOTH the company's aggregated finance context and current external industry/economic information when relevant.

Rules:
- Never invent company figures. Use only numbers in COMPANY_CONTEXT for company-specific claims.
- Treat financial_assurance as the reliability gate. If it is below 90, explicitly flag the weak checks before making strong recommendations.
- Raw transactions are not provided; do not imply that you inspected them.
- Search the live web for relevant macroeconomic, industry, regulatory, demand, rate, FX, or sector facts that materially affect this company.
- Prefer authoritative sources (government/statistical agencies, central banks, regulators, major industry bodies) over generic blogs.
- Separate internal evidence from external context in the reasoning.
- Be concise, commercial and action-oriented: state what changed/what matters, why it matters, and 2-4 management actions.
- If the industry is missing or ambiguous, say that and keep external analysis broad rather than guessing.
- Treat external information as context, not certainty about the company's future.

COMPANY_CONTEXT:
{json.dumps(context, default=str, separators=(",", ":"))}

MANAGEMENT_QUESTION:
{question}
"""

        payload = {
            "model": settings.openai_model,
            "reasoning": {"effort": "low"},
            "tools": [{"type": "web_search", "search_context_size": "low"}],
            "input": prompt,
        }
        headers = {
            "Authorization": f"Bearer {settings.openai_api_key}",
            "Content-Type": "application/json",
        }
        try:
            async with httpx.AsyncClient(timeout=45.0) as client:
                response = await client.post("https://api.openai.com/v1/responses", json=payload, headers=headers)
            response.raise_for_status()
            data = response.json()
        except (httpx.HTTPError, ValueError):
            return None

        text_parts: list[str] = []
        sources: list[dict] = []
        seen_urls: set[str] = set()
        for item in data.get("output", []):
            if item.get("type") != "message":
                continue
            for content in item.get("content", []):
                if content.get("type") == "output_text" and content.get("text"):
                    text_parts.append(content["text"])
                for annotation in content.get("annotations", []) or []:
                    citation = annotation.get("url_citation", annotation)
                    url = citation.get("url") if isinstance(citation, dict) else None
                    if url and url not in seen_urls:
                        seen_urls.add(url)
                        sources.append({"title": citation.get("title") or url, "url": url})
        answer = "\n".join(text_parts).strip() or str(data.get("output_text") or "").strip()
        return (answer, sources[:8]) if answer else None

    @staticmethod
    def _visualization_for_question(question: str, context: dict) -> dict | None:
        """Build chart specs only from deterministic company context.

        The LLM never chooses or fabricates chart values. The question only selects
        which already-grounded dataset is most useful to visualize.
        """
        q = question.lower()
        currency = (context.get("company") or {}).get("currency") or "AUD"
        monthly = context.get("monthly_actuals") or []

        def number(row: dict, key: str) -> float:
            try:
                return float(row.get(key) or 0)
            except (TypeError, ValueError):
                return 0.0

        def label(row: dict, index: int) -> str:
            return str(row.get("month") or row.get("period") or row.get("label") or f"P{index + 1}")

        # Working-capital composition is easiest to read as a donut.
        if any(term in q for term in ("receivable", "collection", "ar ")):
            ar = context.get("ar_summary")
            if ar:
                total = float(ar.get("total_outstanding") or 0) if isinstance(ar, dict) else float(getattr(ar, "total_outstanding", 0) or 0)
                overdue = float(ar.get("overdue_amount") or 0) if isinstance(ar, dict) else float(getattr(ar, "overdue_amount", 0) or 0)
                current = max(total - overdue, 0)
                return {"type":"donut","title":"Receivables exposure","subtitle":"Current versus overdue AR","labels":["Current","Overdue"],"series":[{"name":"Receivables","data":[current,overdue]}],"value_format":"currency","currency":currency}
        if any(term in q for term in ("payable", "supplier", "vendor", "ap ")):
            ap = context.get("ap_summary")
            if ap:
                total = float(ap.get("total_outstanding") or 0) if isinstance(ap, dict) else float(getattr(ap, "total_outstanding", 0) or 0)
                overdue = float(ap.get("overdue_amount") or 0) if isinstance(ap, dict) else float(getattr(ap, "overdue_amount", 0) or 0)
                current = max(total - overdue, 0)
                return {"type":"donut","title":"Payables exposure","subtitle":"Current versus overdue AP","labels":["Current","Overdue"],"series":[{"name":"Payables","data":[current,overdue]}],"value_format":"currency","currency":currency}

        # Balance-sheet questions are better as a direct comparison.
        if any(term in q for term in ("balance", "asset", "liability", "equity")):
            bs = context.get("balance_sheet") or {}
            values = [float(bs.get("total_assets") or 0), float(bs.get("total_liabilities") or 0), float(bs.get("equity") or 0)] if isinstance(bs, dict) else []
            if any(values):
                return {"type":"bar","title":"Balance sheet structure","subtitle":"Assets, liabilities and equity","labels":["Assets","Liabilities","Equity"],"series":[{"name":"Value","data":values}],"value_format":"currency","currency":currency}

        if monthly:
            labels = [label(row, i) for i, row in enumerate(monthly)]
            revenue = [number(row, "revenue") for row in monthly]
            profit = [number(row, "net_profit") for row in monthly]
            gross_profit = [number(row, "gross_profit") for row in monthly]

            if "margin" in q:
                margins = [(gross_profit[i] / revenue[i] * 100) if revenue[i] else 0 for i in range(len(monthly))]
                return {"type":"area","title":"Gross margin trend","subtitle":"Margin by reporting period","labels":labels,"series":[{"name":"Gross margin","data":margins}],"value_format":"percent","currency":currency}
            if any(term in q for term in ("revenue", "sales", "growth", "trend")):
                return {"type":"line","title":"Revenue trend","subtitle":"Revenue across recent reporting periods","labels":labels,"series":[{"name":"Revenue","data":revenue}],"value_format":"currency","currency":currency}
            if any(term in q for term in ("profit", "performance", "management", "focus", "brief", "risk")):
                return {"type":"line","title":"Revenue & net profit trend","subtitle":"Management performance view","labels":labels,"series":[{"name":"Revenue","data":revenue},{"name":"Net profit","data":profit}],"value_format":"currency","currency":currency}

        return None

    async def proactive_signals(self, company_id: UUID) -> dict:
        monthly = await self.reporting.monthly_actuals(company_id)
        if len(monthly) < 2:
            return {"signals": [], "generated_from_months": len(monthly)}

        latest = monthly[-1]
        previous = monthly[-2]
        signals: list[dict] = []

        def pct_change(current, prior):
            current = float(current or 0)
            prior = float(prior or 0)
            return None if abs(prior) < 0.000001 else ((current - prior) / abs(prior)) * 100

        revenue_change = pct_change(latest.get("revenue"), previous.get("revenue"))
        profit_change = pct_change(latest.get("net_profit"), previous.get("net_profit"))
        latest_revenue = float(latest.get("revenue") or 0)
        previous_revenue = float(previous.get("revenue") or 0)
        latest_gp = float(latest.get("gross_profit") or 0)
        previous_gp = float(previous.get("gross_profit") or 0)
        latest_margin = (latest_gp / latest_revenue * 100) if latest_revenue else None
        previous_margin = (previous_gp / previous_revenue * 100) if previous_revenue else None

        if revenue_change is not None and abs(revenue_change) >= 10:
            direction = "increased" if revenue_change > 0 else "declined"
            signals.append({
                "severity": "positive" if revenue_change > 0 else "high",
                "title": f"Revenue {direction} {abs(revenue_change):.1f}% month on month",
                "evidence": f"Latest revenue {latest_revenue:,.0f} versus {previous_revenue:,.0f} in the prior month.",
                "action": "Review the customer, product and branch drivers behind the movement before updating the forecast.",
            })

        if latest_margin is not None and previous_margin is not None:
            margin_move = latest_margin - previous_margin
            if abs(margin_move) >= 2:
                direction = "improved" if margin_move > 0 else "compressed"
                signals.append({
                    "severity": "positive" if margin_move > 0 else "high",
                    "title": f"Gross margin {direction} by {abs(margin_move):.1f} pts",
                    "evidence": f"Gross margin moved from {previous_margin:.1f}% to {latest_margin:.1f}%.",
                    "action": "Check pricing, mix, supplier cost and freight movements to confirm whether the change is structural or temporary.",
                })

        if profit_change is not None and revenue_change is not None and profit_change < revenue_change - 10:
            signals.append({
                "severity": "medium",
                "title": "Profit is lagging revenue growth",
                "evidence": f"Revenue moved {revenue_change:+.1f}% while net profit moved {profit_change:+.1f}% month on month.",
                "action": "Review operating expense growth and margin leakage before committing to additional spend.",
            })

        if not signals:
            signals.append({
                "severity": "low",
                "title": "No material month-on-month anomaly detected",
                "evidence": "Latest revenue, margin and profit movements remained within the current materiality thresholds.",
                "action": "Continue monitoring working capital, forecast variance and branch-level performance.",
            })
        return {"signals": signals[:5], "generated_from_months": len(monthly)}

    async def answer(self, company_id: UUID, question: str, include_external_context: bool = True) -> dict:
        guidance = self._guidance(question)
        if guidance:
            return {
                **guidance,
                "mode": "platform_guide",
                "suggested_questions": [
                    "Where should I upload my General Ledger?",
                    "What files improve working-capital analysis?",
                    "How do I review unmapped accounts?",
                    "How can I reset my data?",
                ],
                "sources": [],
                "external_context_used": False,
                "visualization": None,
            }

        company, context = await self._company_context(company_id)
        q = question.lower()
        visualization = self._visualization_for_question(question, context)
        overview = await self.analytics.overview(company_id)
        pnl = await self.reporting.pnl(company_id)
        bs = await self.reporting.balance_sheet(company_id)
        ar = overview.get("ar_summary")
        ap = overview.get("ap_summary")

        use_web = include_external_context and self._external_context_useful(question)
        if use_web:
            enriched = await self._openai_web_answer(question, context)
            if enriched:
                answer, sources = enriched
                return {
                    "answer": answer,
                    "mode": "company_plus_live_market_context",
                    "suggested_questions": [
                        "Which external risks could affect our next 90 days?",
                        "Are our margins moving in line with industry conditions?",
                        "What should I change in the forecast assumptions?",
                        "What is the biggest management priority in our data?",
                    ],
                    "sources": sources,
                    "action": None,
                    "external_context_used": True,
                    "visualization": visualization,
                }

        if any(word in q for word in ("receivable", "customer", "collection", "ar ")):
            if not ar:
                answer = "AR ageing has not been uploaded yet. Upload it in the Import Centre to activate collection-risk analysis."
                action = {"label": "Upload AR Ageing", "route": "/dashboard/import-center"}
            else:
                top_ar = (
                    f" The largest current exposure is {ar.top_parties[0].party_name} at {ar.top_parties[0].outstanding_amount:,.2f}."
                    if ar.top_parties else ""
                )
                answer = (
                    f"Total receivables are {ar.total_outstanding:,.2f}. Overdue receivables are "
                    f"{ar.overdue_amount:,.2f}, or {ar.overdue_percent:.1f}% of AR." + top_ar
                )
                action = {"label": "Open Working Capital", "route": "/dashboard/working-capital"}
        elif any(word in q for word in ("payable", "vendor", "supplier", "payment", "ap ")):
            if not ap:
                answer = "AP ageing has not been uploaded yet. Upload it in the Import Centre to activate vendor and payment-timing analysis."
                action = {"label": "Upload AP Ageing", "route": "/dashboard/import-center"}
            else:
                top_ap = (
                    f" The largest vendor exposure is {ap.top_parties[0].party_name} at {ap.top_parties[0].outstanding_amount:,.2f}."
                    if ap.top_parties else ""
                )
                answer = (
                    f"Total payables are {ap.total_outstanding:,.2f}. Overdue payables are {ap.overdue_amount:,.2f}, "
                    f"or {ap.overdue_percent:.1f}% of AP." + top_ap
                )
                action = {"label": "Open Working Capital", "route": "/dashboard/working-capital"}
        elif any(word in q for word in ("profit", "margin", "revenue", "p&l", "pnl")):
            answer = (
                f"Revenue is {pnl.revenue:,.2f}, gross profit is {pnl.gross_profit:,.2f}, and net profit is {pnl.net_profit:,.2f}. "
                "Use Analytics to review monthly movement and branch drivers."
            )
            action = {"label": "Open Analytics", "route": "/dashboard/analytics"}
        elif any(word in q for word in ("balance", "cash", "liability", "asset")):
            answer = (
                f"Total assets are {bs.total_assets:,.2f}, liabilities are {bs.total_liabilities:,.2f}, equity is {bs.equity:,.2f}, "
                f"and the balance-sheet difference is {bs.balance_difference:,.2f}."
            )
            action = {"label": "Open Reports", "route": "/dashboard/reports"}
        else:
            company_name = (company.trading_name or company.legal_name) if company else "your company"
            answer = (
                f"I’m analysing {company_name} from the finance data currently loaded. "
                + " ".join(overview.get("insights", [])[:3])
                + " Ask me about profitability, cash, working capital, forecasts, or an industry/economic risk and I can connect the internal numbers with live external context."
            )
            action = {"label": "Open Analytics", "route": "/dashboard/analytics"}

        return {
            "answer": answer,
            "mode": "grounded_finance_assistant",
            "suggested_questions": [
                "What is the biggest management priority in our data?",
                "What external risks could affect us this quarter?",
                "Where should I upload AR and AP ageing?",
                "How is monthly revenue trending?",
            ],
            "sources": [],
            "action": action,
            "external_context_used": False,
            "visualization": visualization,
        }
