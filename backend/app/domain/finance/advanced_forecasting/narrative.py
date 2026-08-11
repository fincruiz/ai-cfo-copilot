from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, List, Optional
import json
import os
import pandas as pd

from .models import BenchmarkData, CompanyProfile
from .three_way import ThreeWayForecastResult


@dataclass
class NarrativePackage:
    executive_summary: str
    financial_performance: str
    forecast_outlook: str
    working_capital: str
    liquidity_and_funding: str
    benchmark_analysis: str
    market_environment: str
    risks: List[str]
    opportunities: List[str]
    board_actions: List[str]
    ai_used: bool = False


class NarrativeEngine:
    """Produces deterministic finance commentary and optionally enriches it with AI.

    AI is used only after structured metrics are calculated. This keeps token use
    controlled and prevents the model from performing accounting calculations.
    """

    def __init__(self, model: str = "gpt-4.1-mini"):
        self.model = model

    @staticmethod
    def _pct(value: float) -> str:
        return f"{value:.1%}"

    @staticmethod
    def _money(value: float, currency: str) -> str:
        return f"{currency} {value:,.0f}"

    def deterministic(
        self,
        profile: CompanyProfile,
        result: ThreeWayForecastResult,
        benchmarks: BenchmarkData,
        scenario_comparison: Optional[pd.DataFrame] = None,
    ) -> NarrativePackage:
        pl = result.profit_and_loss
        bs = result.balance_sheet
        ratios = result.ratios

        revenue_total = pl["Revenue"].sum()
        ebitda_total = pl["EBITDA"].sum()
        net_income_total = pl["Net Income"].sum()
        gm = pl["Gross Profit"].sum() / max(revenue_total, 1)
        ebitda_margin = ebitda_total / max(revenue_total, 1)
        closing_cash = bs["Cash"].iloc[-1]
        closing_debt = (
            bs["Current Debt"].iloc[-1]
            + bs["Non-current Debt"].iloc[-1]
        )
        min_cash = bs["Cash"].min()
        dso = ratios["DSO"].iloc[-1]
        ccc = ratios["Cash Conversion Cycle"].iloc[-1]

        executive_summary = (
            f"{profile.company_name} is forecast to generate "
            f"{self._money(revenue_total, profile.currency)} of revenue over the "
            f"forecast horizon, with gross margin of {self._pct(gm)} and EBITDA "
            f"margin of {self._pct(ebitda_margin)}. Closing cash is forecast at "
            f"{self._money(closing_cash, profile.currency)} and closing debt at "
            f"{self._money(closing_debt, profile.currency)}."
        )

        financial_performance = (
            f"Forecast EBITDA is {self._money(ebitda_total, profile.currency)} "
            f"and forecast net income is "
            f"{self._money(net_income_total, profile.currency)}. The model uses "
            f"a blended forecast built from management budget, historical trend, "
            f"recent actual run rate and monthly seasonality."
        )

        forecast_outlook = (
            "The forecast should be treated as a rolling management view rather "
            "than a fixed annual plan. Each month, the latest actual results should "
            "replace the forecast month and the remaining periods should be recalculated."
        )
        if scenario_comparison is not None and not scenario_comparison.empty:
            best = scenario_comparison["EBITDA"].idxmax()
            weakest = scenario_comparison["EBITDA"].idxmin()
            forecast_outlook += (
                f" Scenario analysis identifies {best} as the strongest EBITDA "
                f"outcome and {weakest} as the weakest."
            )

        working_capital = (
            f"Receivables are modelled using DSO of {dso:.1f} days and the "
            f"cash conversion cycle is {ccc:.1f} days. Management should focus on "
            f"collection discipline, inventory velocity and supplier terms because "
            f"small movements in working-capital days can materially change cash."
        )

        liquidity = (
            f"Minimum forecast cash is {self._money(min_cash, profile.currency)}. "
            f"The integrated funding logic draws debt when cash falls below the "
            f"minimum-cash threshold and sweeps surplus cash against debt."
        )

        benchmark_lines = []
        metric_map = {
            "Gross Margin": gm,
            "EBITDA Margin": ebitda_margin,
            "DSO": dso,
            "Cash Conversion Cycle": ccc,
        }
        for metric, actual in metric_map.items():
            if metric in benchmarks.metrics:
                benchmark = benchmarks.metrics[metric]
                difference = actual - benchmark
                if "Margin" in metric:
                    benchmark_lines.append(
                        f"{metric} is {self._pct(actual)} versus benchmark "
                        f"{self._pct(benchmark)}, a variance of "
                        f"{difference * 100:.1f} percentage points."
                    )
                else:
                    benchmark_lines.append(
                        f"{metric} is {actual:.1f} versus benchmark "
                        f"{benchmark:.1f}, a variance of {difference:.1f}."
                    )
        benchmark_analysis = (
            " ".join(benchmark_lines)
            if benchmark_lines
            else "No quantified industry benchmarks were supplied."
        )

        market_environment = " ".join(
            benchmarks.macro_environment
            + benchmarks.industry_trends
            + benchmarks.competitor_observations
        ) or "No external market research was supplied."

        risks = []
        if min_cash < 0:
            risks.append("Forecast liquidity becomes negative.")
        if closing_debt > closing_cash * 2:
            risks.append("Debt remains high relative to closing cash.")
        if ebitda_margin < 0.10:
            risks.append("EBITDA margin is below 10%, limiting downside protection.")
        if dso > 60:
            risks.append("Receivable days exceed 60 days and may pressure liquidity.")
        if not risks:
            risks.append("Execution risk remains around achieving forecast revenue and margin assumptions.")

        opportunities = [
            "Refresh the forecast monthly using actual results and revised business drivers.",
            "Use customer, product, branch and channel data to identify the highest-return growth areas.",
            "Model targeted working-capital improvements and convert them into accountable actions.",
        ]

        board_actions = [
            "Approve the base forecast and downside liquidity trigger points.",
            "Confirm management owners for revenue, gross margin and working-capital actions.",
            "Review actual-versus-forecast performance monthly and reforecast the remaining horizon.",
            "Approve any debt, capex or hiring actions required by the selected scenario.",
        ]

        return NarrativePackage(
            executive_summary=executive_summary,
            financial_performance=financial_performance,
            forecast_outlook=forecast_outlook,
            working_capital=working_capital,
            liquidity_and_funding=liquidity,
            benchmark_analysis=benchmark_analysis,
            market_environment=market_environment,
            risks=risks,
            opportunities=opportunities,
            board_actions=board_actions,
            ai_used=False,
        )

    def enrich_with_openai(
        self,
        base: NarrativePackage,
        profile: CompanyProfile,
        metrics_payload: Dict,
        research_payload: Dict,
    ) -> NarrativePackage:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            return base

        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key)

            payload = {
                "company": profile.__dict__,
                "calculated_metrics": metrics_payload,
                "external_research": research_payload,
                "deterministic_narrative": base.__dict__,
            }
            prompt = (
                "You are preparing board-level finance commentary. "
                "Use only the supplied calculated metrics and research. "
                "Do not recalculate accounting figures. Do not invent facts, "
                "sources, competitors or benchmarks. Return strict JSON with keys: "
                "executive_summary, financial_performance, forecast_outlook, "
                "working_capital, liquidity_and_funding, benchmark_analysis, "
                "market_environment, risks, opportunities, board_actions. "
                "Each narrative section should be detailed, commercially practical "
                "and suitable for directors. Risks, opportunities and board_actions "
                "must be arrays of concise statements.\n\n"
                + json.dumps(payload, default=str)
            )
            response = client.responses.create(
                model=self.model,
                input=prompt,
                temperature=0.2,
            )
            data = json.loads(response.output_text)
            return NarrativePackage(
                executive_summary=data["executive_summary"],
                financial_performance=data["financial_performance"],
                forecast_outlook=data["forecast_outlook"],
                working_capital=data["working_capital"],
                liquidity_and_funding=data["liquidity_and_funding"],
                benchmark_analysis=data["benchmark_analysis"],
                market_environment=data["market_environment"],
                risks=data["risks"],
                opportunities=data["opportunities"],
                board_actions=data["board_actions"],
                ai_used=True,
            )
        except Exception:
            # The report must still generate if AI is unavailable.
            return base
