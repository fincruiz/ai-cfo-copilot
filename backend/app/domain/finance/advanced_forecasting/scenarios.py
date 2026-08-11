from __future__ import annotations

from copy import deepcopy
from typing import Dict
import pandas as pd

from .models import ForecastConfig, ScenarioDefinition
from .three_way import ThreeWayForecastEngine, ThreeWayForecastResult


class ScenarioManager:
    def __init__(self, base_operating_forecast: pd.DataFrame, base_config: ForecastConfig):
        self.base_forecast = base_operating_forecast.copy()
        self.base_config = base_config
        self.scenarios: Dict[str, ScenarioDefinition] = {}

    def add(self, scenario: ScenarioDefinition) -> None:
        self.scenarios[scenario.name] = scenario

    def add_standard_scenarios(self) -> None:
        self.add(ScenarioDefinition(
            name="Base Case",
            description="Blended historical trend, budget and recent run rate.",
        ))
        self.add(ScenarioDefinition(
            name="Upside",
            description="Stronger demand, better margin and faster collections.",
            revenue_multiplier=1.10,
            gross_margin_delta=0.03,
            dso_delta_days=-7,
            capex_multiplier=1.10,
        ))
        self.add(ScenarioDefinition(
            name="Downside",
            description="Demand pressure, weaker margin and working-capital stress.",
            revenue_multiplier=0.90,
            gross_margin_delta=-0.05,
            dso_delta_days=15,
            inventory_delta_days=15,
            payroll_pct_delta=0.015,
            other_opex_pct_delta=0.01,
        ))
        self.add(ScenarioDefinition(
            name="Expansion",
            description="Growth investment with higher revenue, payroll and capex.",
            revenue_multiplier=1.18,
            payroll_pct_delta=0.025,
            capex_multiplier=2.0,
            dso_delta_days=5,
        ))

    def _build(self, scenario: ScenarioDefinition) -> ThreeWayForecastResult:
        forecast = self.base_forecast.copy()
        config = deepcopy(self.base_config)
        drivers = config.drivers

        forecast["Revenue"] *= scenario.revenue_multiplier

        historic_gm = (
            (forecast["Revenue"] - forecast["COGS"])
            / forecast["Revenue"].replace(0, pd.NA)
        ).fillna(drivers.gross_margin)
        scenario_gm = (historic_gm + scenario.gross_margin_delta).clip(0.0, 0.95)
        forecast["COGS"] = forecast["Revenue"] * (1 - scenario_gm)

        payroll_ratio = (
            forecast["Payroll"]
            / forecast["Revenue"].replace(0, pd.NA)
        ).fillna(drivers.payroll_pct_revenue)
        other_opex_ratio = (
            forecast["Other Opex"]
            / forecast["Revenue"].replace(0, pd.NA)
        ).fillna(drivers.other_opex_pct_revenue)

        forecast["Payroll"] = forecast["Revenue"] * (
            payroll_ratio + scenario.payroll_pct_delta
        ).clip(lower=0.0)
        forecast["Other Opex"] = forecast["Revenue"] * (
            other_opex_ratio + scenario.other_opex_pct_delta
        ).clip(lower=0.0)

        drivers.dso_days = max(0.0, drivers.dso_days + scenario.dso_delta_days)
        drivers.dpo_days = max(0.0, drivers.dpo_days + scenario.dpo_delta_days)
        drivers.inventory_days = max(
            0.0,
            drivers.inventory_days + scenario.inventory_delta_days,
        )
        drivers.capex_pct_revenue *= scenario.capex_multiplier
        drivers.annual_interest_rate = max(
            0.0,
            drivers.annual_interest_rate + scenario.interest_rate_delta,
        )

        return ThreeWayForecastEngine(forecast, config).run()

    def run_all(self) -> Dict[str, ThreeWayForecastResult]:
        if not self.scenarios:
            self.add_standard_scenarios()
        return {
            name: self._build(scenario)
            for name, scenario in self.scenarios.items()
        }

    @staticmethod
    def comparison(results: Dict[str, ThreeWayForecastResult]) -> pd.DataFrame:
        rows = []
        for name, result in results.items():
            pl = result.profit_and_loss
            bs = result.balance_sheet
            debt = bs["Current Debt"] + bs["Non-current Debt"]
            rows.append({
                "Scenario": name,
                "Revenue": pl["Revenue"].sum(),
                "Gross Profit": pl["Gross Profit"].sum(),
                "EBITDA": pl["EBITDA"].sum(),
                "Net Income": pl["Net Income"].sum(),
                "Closing Cash": bs["Cash"].iloc[-1],
                "Minimum Cash": bs["Cash"].min(),
                "Closing Debt": debt.iloc[-1],
                "Peak Debt": debt.max(),
                "Ending Equity": bs["Total Equity"].iloc[-1],
                "Balanced": bool(result.checks["Balanced"].all()),
            })
        return pd.DataFrame(rows).set_index("Scenario")
