from __future__ import annotations

import json
import re

import httpx

from app.core.config import settings


DEMO_CONTEXT = {
    "company": "Nova Retail",
    "currency": "INR",
    "period": "FY2026 synthetic demo",
    "revenue": 24_800_000,
    "prior_revenue": 22_670_000,
    "net_profit": 4_120_000,
    "prior_net_profit": 3_960_000,
    "cash": 6_210_000,
    "gross_margin": 42.4,
    "overdue_ar": 1_180_000,
    "overdue_ar_percent": 28.0,
    "debtor_days": 54,
    "branches": [
        {"name": "Central", "revenue": 11_800_000, "gp_percent": 44.8, "net_profit": 2_050_000},
        {"name": "North", "revenue": 7_400_000, "gp_percent": 41.6, "net_profit": 1_240_000},
        {"name": "West", "revenue": 5_600_000, "gp_percent": 36.7, "net_profit": 830_000},
    ],
    "monthly_revenue": [1.62, 1.68, 1.75, 1.82, 1.89, 1.94, 2.02, 2.08, 2.15, 2.22, 2.31, 2.42],
    "monthly_cash": [5.90, 5.70, 5.50, 5.40, 5.10, 4.90, 4.80, 4.70, 4.90, 5.20, 5.70, 6.21],
    "budget": {
        "revenue": 24_000_000,
        "net_profit": 4_300_000,
        "gross_margin": 43.0,
    },
    "forecast": {
        "base": {"revenue": 26_040_000, "net_profit": 4_410_000, "closing_cash": 6_780_000},
        "revenue_growth_10": {"revenue": 27_280_000, "net_profit": 4_780_000, "closing_cash": 7_420_000},
    },
    "scenario": {
        "hire_3": {"closing_cash": 4_080_000, "profit": 3_460_000, "buffer": 3_500_000},
        "downside": {"closing_cash": 2_920_000},
        "collections_improve_10_days": {"cash_release": 820_000, "closing_cash": 7_030_000},
    },
}

SUGGESTIONS = [
    "What should management focus on?",
    "Why is profit up but cash tight?",
    "Which branch needs attention?",
    "Can we afford to hire 3 people?",
    "What happens if revenue grows 10%?",
    "Build a 12-month forecast.",
    "Where are we losing margin?",
    "How can we improve working capital?",
]


def _money(value: float | int) -> str:
    return f"₹{value / 1_000_000:.2f}M"


def _chart(title: str, labels: list[str], series: list[dict], fmt: str = "currency") -> dict:
    return {
        "type": "line",
        "title": title,
        "subtitle": "Synthetic demo data",
        "labels": labels,
        "series": series,
        "value_format": fmt,
        "currency": "INR",
    }


def _suggest(*items: str) -> list[str]:
    result: list[str] = []
    for item in items:
        if item not in result:
            result.append(item)
    for item in SUGGESTIONS:
        if item not in result:
            result.append(item)
    return result[:4]


def deterministic_demo_answer(question: str) -> dict:
    q = question.lower().strip()
    evidence: list[dict] = []
    action = None
    visualization = None

    def add_evidence(label: str, value: str, source: str = "Synthetic management data") -> None:
        evidence.append({"label": label, "value": value, "source": source})

    # Scenario intent must be checked before generic revenue / growth language.
    growth_match = re.search(r"(?:revenue|sales).*?(\d{1,2})\s*%", q)
    growth_intent = bool(growth_match) and any(token in q for token in ("what happens", "grow", "growth", "increase", "scenario", "forecast"))

    if growth_intent:
        requested = int(growth_match.group(1)) if growth_match else 10
        if requested == 10:
            scenario = DEMO_CONTEXT["forecast"]["revenue_growth_10"]
            answer = (
                "In the synthetic +10% revenue scenario, annual revenue rises to ₹27.28M, net profit to ₹4.78M and projected closing cash to ₹7.42M. "
                "The decision still depends on maintaining collection discipline and avoiding further margin leakage in West, so the growth target should be reviewed with both cash conversion and branch margin assumptions visible."
            )
            add_evidence("Scenario revenue", "₹27.28M", "Synthetic 12-month forecast")
            add_evidence("Scenario net profit", "₹4.78M", "Synthetic 12-month forecast")
            add_evidence("Scenario closing cash", "₹7.42M", "Synthetic three-way forecast")
            visualization = {
                "type": "bar",
                "title": "Current vs +10% revenue scenario",
                "subtitle": "Synthetic forecast comparison",
                "labels": ["Current", "+10% scenario"],
                "series": [{"name": "Revenue (₹M)", "data": [24.8, 27.28]}],
                "value_format": "currency",
                "currency": "INR",
            }
            action = {"label": "Explore Forecasting", "demo_anchor": "forecasting"}
            suggestions = _suggest("Can we afford to hire 3 people?", "How can we improve working capital?", "Where are we losing margin?")
        else:
            answer = (
                f"The fixed public demo has a prepared +10% revenue scenario rather than a {requested}% scenario. "
                "To avoid inventing forecast outputs, I will not fabricate the requested case. Try the +10% revenue scenario or create a workspace to model your own assumptions."
            )
            suggestions = _suggest("What happens if revenue grows 10%?", "Build a 12-month forecast.")

    elif any(token in q for token in ("forecast", "next 12 months", "12-month", "12 month", "forward outlook")):
        forecast = DEMO_CONTEXT["forecast"]["base"]
        answer = (
            "The synthetic base forecast projects ₹26.04M of revenue, ₹4.41M of net profit and ₹6.78M of closing cash over the next 12 months. "
            "The main sensitivity is cash conversion: collections and West branch margin remain the two assumptions management should challenge before relying on the upside case."
        )
        add_evidence("Forecast revenue", "₹26.04M", "Synthetic 12-month forecast")
        add_evidence("Forecast net profit", "₹4.41M", "Synthetic 12-month forecast")
        add_evidence("Forecast closing cash", "₹6.78M", "Synthetic three-way forecast")
        visualization = _chart(
            "Historical revenue trend",
            [f"M{i}" for i in range(1, 13)],
            [{"name": "Revenue (₹M)", "data": DEMO_CONTEXT["monthly_revenue"]}],
        )
        action = {"label": "Explore Forecasting", "demo_anchor": "forecasting"}
        suggestions = _suggest("What happens if revenue grows 10%?", "Can we afford to hire 3 people?", "What are the biggest risks management should discuss?")

    elif any(token in q for token in ("cash", "working capital", "receivable", "debtor", "collection")):
        if any(token in q for token in ("improve", "reduce debtor", "10 days", "ten days")):
            improvement = DEMO_CONTEXT["scenario"]["collections_improve_10_days"]
            answer = (
                "A 10-day improvement in debtor days releases about ₹0.82M of cash in the synthetic scenario, lifting projected cash to roughly ₹7.03M. "
                "That is more valuable than treating overdue AR as a reporting metric alone because it gives management a measurable collection target."
            )
            add_evidence("Cash released", "₹0.82M", "Synthetic working-capital scenario")
            add_evidence("Projected cash", "₹7.03M", "Synthetic three-way forecast")
            add_evidence("Current debtor days", "54 days", "Synthetic working-capital KPI")
            action = {"label": "Explore Working Capital", "demo_anchor": "working-capital"}
            suggestions = _suggest("Why is profit up but cash tight?", "What should management focus on?", "Can we afford to hire 3 people?")
        else:
            answer = (
                "Cash is tighter than profit suggests because receivables are absorbing working capital. The synthetic business has ₹1.18M overdue in AR and debtor days of 54, so reported profit is converting to cash more slowly than management would want."
            )
            add_evidence("Overdue AR", "₹1.18M", "Synthetic AR ageing")
            add_evidence("Debtor days", "54 days", "Synthetic working-capital KPI")
            add_evidence("Cash", "₹6.21M", "Synthetic balance sheet")
            visualization = _chart(
                "Revenue vs cash conversion",
                [f"M{i}" for i in range(1, 13)],
                [
                    {"name": "Revenue", "data": DEMO_CONTEXT["monthly_revenue"]},
                    {"name": "Cash", "data": DEMO_CONTEXT["monthly_cash"]},
                ],
            )
            action = {"label": "Explore Working Capital", "demo_anchor": "working-capital"}
            suggestions = _suggest("How can we improve working capital?", "Can we afford to hire 3 people?", "What should management focus on?")

    elif any(token in q for token in ("branch", "west", "location", "site", "losing margin", "margin leakage")):
        branches = DEMO_CONTEXT["branches"]
        weak = min(branches, key=lambda item: item["gp_percent"])
        answer = (
            f"West needs the closest management attention. Its {weak['gp_percent']:.1f}% gross margin is materially below Central at 44.8%, so the issue is margin quality rather than simply revenue volume. "
            "Management should investigate pricing, discounting and cost mix in West before pushing more growth through that branch."
        )
        add_evidence("West GP%", "36.7%", "Synthetic branch P&L")
        add_evidence("Central GP%", "44.8%", "Synthetic branch P&L")
        add_evidence("West revenue", "₹5.60M", "Synthetic branch P&L")
        visualization = {
            "type": "bar",
            "title": "Branch gross margin",
            "subtitle": "Synthetic branch comparison",
            "labels": [item["name"] for item in branches],
            "series": [{"name": "GP%", "data": [item["gp_percent"] for item in branches]}],
            "value_format": "percent",
            "currency": "INR",
        }
        action = {"label": "Compare Branch Drivers", "demo_anchor": "branches"}
        suggestions = _suggest("Where are we losing margin?", "What should management focus on?", "What happens if revenue grows 10%?")

    elif any(token in q for token in ("hire", "can we afford", "afford to hire", "add staff", "new staff", "additional people")):
        scenario = DEMO_CONTEXT["scenario"]["hire_3"]
        downside = DEMO_CONTEXT["scenario"]["downside"]
        answer = (
            "In the base hiring scenario, three additional people are affordable: projected closing cash is ₹4.08M, still above the ₹3.50M management buffer. "
            "The downside case is the risk — slower collections would push closing cash to about ₹2.92M, below the buffer."
        )
        add_evidence("Hiring closing cash", "₹4.08M", "Synthetic three-way forecast")
        add_evidence("Management buffer", "₹3.50M", "Synthetic policy assumption")
        add_evidence("Downside closing cash", "₹2.92M", "Synthetic downside scenario")
        add_evidence("Scenario profit", "₹3.46M", "Synthetic hiring scenario")
        action = {"label": "Explore Decision Simulation", "demo_anchor": "decision"}
        suggestions = _suggest("How can we improve working capital?", "What happens if revenue grows 10%?", "What are the biggest financial risks?")

    elif any(token in q for token in ("budget", "variance", "against budget", "vs budget")):
        budget = DEMO_CONTEXT["budget"]
        revenue_variance = DEMO_CONTEXT["revenue"] - budget["revenue"]
        profit_variance = DEMO_CONTEXT["net_profit"] - budget["net_profit"]
        answer = (
            "Revenue is ₹0.80M ahead of the synthetic budget, but net profit is ₹0.18M behind budget. That means the business is growing faster than planned without converting all of that upside into profit. "
            "West branch margin and working-capital pressure are the two areas management should investigate before treating the revenue beat as fully positive."
        )
        add_evidence("Revenue vs budget", f"+{_money(revenue_variance)}", "Synthetic budget comparison")
        add_evidence("Net profit vs budget", _money(profit_variance), "Synthetic budget comparison")
        add_evidence("Actual gross margin", "42.4%", "Synthetic P&L")
        action = {"label": "Explore Reporting", "demo_anchor": "reporting"}
        suggestions = _suggest("Where are we losing margin?", "Which branch needs attention?", "Build a 12-month forecast.")

    elif any(token in q for token in ("board", "board pack", "board meeting")):
        answer = (
            "For the next board discussion, the synthetic evidence points to four items: revenue growth remains positive, cash conversion needs attention, West branch margin is diluting group quality, and downside collections can push cash below the management buffer. "
            "The board story should connect those risks to specific actions rather than present them as separate reports."
        )
        add_evidence("Revenue", "₹24.80M", "Synthetic P&L")
        add_evidence("Overdue AR", "₹1.18M", "Synthetic AR ageing")
        add_evidence("West GP%", "36.7%", "Synthetic branch P&L")
        add_evidence("Downside cash", "₹2.92M", "Synthetic scenario")
        action = {"label": "Explore Board Story", "demo_anchor": "board"}
        suggestions = _suggest("What are the biggest financial risks?", "What should management focus on?", "Can we afford to hire 3 people?")

    elif any(token in q for token in ("risk", "priority", "focus", "management")):
        answer = (
            "The three highest-priority issues are working-capital conversion, West branch margin leakage, and downside cash resilience if collections slow. "
            "Growth is positive, but management should protect cash quality before accelerating fixed-cost commitments."
        )
        add_evidence("Overdue AR", "₹1.18M", "Synthetic AR ageing")
        add_evidence("West GP%", "36.7%", "Synthetic branch P&L")
        add_evidence("Downside cash", "₹2.92M", "Synthetic scenario")
        action = {"label": "See Management Story", "demo_anchor": "reporting"}
        suggestions = _suggest("Why is profit up but cash tight?", "Which branch needs attention?", "Can we afford to hire 3 people?")

    elif any(token in q for token in ("revenue", "sales", "growth", "trend")):
        revenue_growth = (DEMO_CONTEXT["revenue"] / DEMO_CONTEXT["prior_revenue"] - 1) * 100
        answer = (
            f"Revenue is trending upward, reaching ₹24.80M for the synthetic year and growing about {revenue_growth:.1f}% versus the prior period. "
            "The latest synthetic month is ₹2.42M. Management should still monitor margin and cash conversion so growth does not hide deteriorating quality."
        )
        add_evidence("Annual revenue", "₹24.80M", "Synthetic P&L")
        add_evidence("Revenue growth", f"{revenue_growth:.1f}%", "Synthetic period comparison")
        add_evidence("Latest monthly revenue", "₹2.42M", "Synthetic monthly actuals")
        visualization = _chart(
            "Revenue trend",
            [f"M{i}" for i in range(1, 13)],
            [{"name": "Revenue (₹M)", "data": DEMO_CONTEXT["monthly_revenue"]}],
        )
        action = {"label": "Explore Forecasting", "demo_anchor": "forecasting"}
        suggestions = _suggest("What happens if revenue grows 10%?", "Where are we losing margin?", "Why is profit up but cash tight?")

    elif any(token in q for token in ("profit", "margin", "gross")):
        answer = (
            "The synthetic company is profitable at ₹4.12M net profit with a 42.4% gross margin. Profit is slightly above the prior period, but the management issue is quality: West's 36.7% gross margin is diluting an otherwise healthier group result, while receivables are slowing cash conversion."
        )
        add_evidence("Net profit", "₹4.12M", "Synthetic P&L")
        add_evidence("Gross margin", "42.4%", "Synthetic P&L")
        add_evidence("West GP%", "36.7%", "Synthetic branch P&L")
        action = {"label": "Compare Branch Drivers", "demo_anchor": "branches"}
        suggestions = _suggest("Where are we losing margin?", "Why is profit up but cash tight?", "Which branch needs attention?")

    else:
        answer = (
            "The demo can analyse profitability, cash and working capital, branch performance, budget variance, trends, forecasts, risks and prepared decision scenarios from the synthetic Nova Retail dataset. "
            "This question needs evidence outside that demo dataset, so I won't invent an answer. Try a management question about revenue, margin, cash, branches, hiring, forecasting, working capital or risk."
        )
        suggestions = _suggest("What should management focus on?", "Build a 12-month forecast.", "What happens if revenue grows 10%?")

    return {
        "answer": answer,
        "mode": "synthetic_demo",
        "evidence": evidence,
        "confidence": "high" if evidence else "medium",
        "confidence_reason": "This answer is grounded only in the fixed synthetic Nova Retail demo dataset; no customer data is used.",
        "suggested_questions": suggestions,
        "visualization": visualization,
        "action": action,
    }


async def answer_demo_question(question: str) -> dict:
    fallback = deterministic_demo_answer(question)
    if not settings.openai_api_key:
        return fallback

    # Public demo AI receives only synthetic aggregates. It never receives
    # authenticated customer/company context or raw customer transactions.
    prompt = {
        "question": question,
        "synthetic_context": DEMO_CONTEXT,
        "rules": [
            "Use only synthetic_context.",
            "Do not claim access to live customer data.",
            "If evidence is missing, say so.",
            "Keep the answer under 150 words.",
            "Use management language, not accounting jargon.",
            "Do not override the deterministic evidence, confidence or action supplied by the application.",
        ],
    }

    try:
        payload = {
            "model": settings.openai_model,
            "input": [
                {
                    "role": "system",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "You are the public FinCruiz product demo. Answer only from the supplied synthetic company context. Never invent missing evidence.",
                        }
                    ],
                },
                {"role": "user", "content": [{"type": "input_text", "text": json.dumps(prompt)}]},
            ],
            "max_output_tokens": 320,
        }
        async with httpx.AsyncClient(timeout=18) as client:
            response = await client.post(
                "https://api.openai.com/v1/responses",
                json=payload,
                headers={
                    "Authorization": f"Bearer {settings.openai_api_key}",
                    "Content-Type": "application/json",
                },
            )
        if response.status_code >= 400:
            return fallback

        data = response.json()
        parts: list[str] = []
        for item in data.get("output", []):
            for content in item.get("content", []):
                if content.get("type") in {"output_text", "text"} and content.get("text"):
                    parts.append(str(content["text"]))
        answer = "\n".join(parts).strip()
        if answer:
            return {**fallback, "answer": answer, "mode": "synthetic_demo_ai"}
    except Exception:
        pass

    return fallback
