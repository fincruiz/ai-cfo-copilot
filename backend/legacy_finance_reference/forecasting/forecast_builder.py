from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Optional, Tuple
import numpy as np
import pandas as pd

from .models import ForecastConfig, HistoricalData


@dataclass
class ForecastBuildResult:
    forecast: pd.DataFrame
    diagnostics: pd.DataFrame
    seasonality_indices: pd.DataFrame


class TrendBudgetForecastBuilder:
    """Builds a monthly operating forecast by blending:

    1. Historical trend projection.
    2. Management budget.
    3. Recent actual run rate.
    4. Optional monthly seasonality.

    The output is designed to feed the integrated three-way engine.
    """

    FORECAST_LINES = ["Revenue", "COGS", "Payroll", "Other Opex"]

    def __init__(
        self,
        historical: HistoricalData,
        budget: Optional[pd.DataFrame],
        config: ForecastConfig,
    ):
        self.history = historical.validated()
        self.budget = self._prepare_budget(budget)
        self.config = config
        self.config.validate()

    @staticmethod
    def _prepare_budget(budget: Optional[pd.DataFrame]) -> Optional[pd.DataFrame]:
        if budget is None or budget.empty:
            return None
        df = budget.copy()
        if "Period" not in df.columns:
            raise ValueError("Budget must contain a Period column")
        df["Period"] = pd.to_datetime(df["Period"])
        for col in df.columns:
            if col != "Period":
                df[col] = pd.to_numeric(df[col], errors="coerce")
        return df.sort_values("Period").drop_duplicates("Period", keep="last")

    def _linear_trend(self, series: pd.Series, horizon: int) -> Tuple[np.ndarray, float]:
        values = series.astype(float).to_numpy()
        if len(values) < 2:
            return np.repeat(values[-1] if len(values) else 0.0, horizon), 0.0

        limit = max(1, min(self.config.history_months_for_trend, len(values)))
        y = values[-limit:].copy()

        # Winsorise month-on-month changes to limit one-off distortions.
        if len(y) > 2:
            prior = np.where(np.abs(y[:-1]) < 1e-9, 1.0, y[:-1])
            changes = (y[1:] - y[:-1]) / prior
            changes = np.clip(
                changes,
                -self.config.outlier_winsor_limit,
                self.config.outlier_winsor_limit,
            )
            adjusted = [y[0]]
            for change in changes:
                adjusted.append(adjusted[-1] * (1 + change))
            y = np.array(adjusted)

        x = np.arange(len(y), dtype=float)
        slope, intercept = np.polyfit(x, y, 1)
        future_x = np.arange(len(y), len(y) + horizon, dtype=float)
        projected = intercept + slope * future_x

        baseline = max(abs(y[-1]), 1.0)
        monthly_growth = np.clip(
            slope / baseline,
            self.config.min_monthly_growth,
            self.config.max_monthly_growth,
        )
        geometric = y[-1] * (1 + monthly_growth) ** np.arange(1, horizon + 1)

        # Blend linear and geometric projection to reduce unrealistic divergence.
        output = 0.5 * projected + 0.5 * geometric
        return np.maximum(output, 0.0), float(monthly_growth)

    def _seasonality(self, line: str) -> pd.Series:
        df = self.history[["Period", line]].copy()
        df["Month"] = df["Period"].dt.month
        monthly_avg = df.groupby("Month")[line].mean()
        overall = max(df[line].mean(), 1e-9)
        index = monthly_avg / overall
        index = index.reindex(range(1, 13)).fillna(1.0)
        return index / max(index.mean(), 1e-9)

    def _budget_values(self, periods: pd.DatetimeIndex, line: str) -> np.ndarray:
        if self.budget is None or line not in self.budget.columns:
            return np.full(len(periods), np.nan)
        lookup = self.budget.set_index("Period")[line]
        return np.array([lookup.get(period, np.nan) for period in periods], dtype=float)

    def build(self) -> ForecastBuildResult:
        periods = pd.date_range(
            start=pd.Timestamp(self.config.forecast_start),
            periods=self.config.forecast_months,
            freq="MS",
        )
        output = pd.DataFrame({"Period": periods})
        diagnostic_rows = []
        seasonality_rows = []

        for line in self.FORECAST_LINES:
            trend_projection, trend_growth = self._linear_trend(
                self.history[line],
                len(periods),
            )

            recent_count = min(
                self.config.recent_run_rate_months,
                len(self.history),
            )
            recent_run_rate = float(self.history[line].tail(recent_count).mean())
            recent_projection = np.repeat(recent_run_rate, len(periods))

            budget_values = self._budget_values(periods, line)
            budget_available = ~np.isnan(budget_values)

            if self.config.seasonality_enabled:
                seasonality = self._seasonality(line)
                seasonal_factors = np.array(
                    [seasonality.get(period.month, 1.0) for period in periods]
                )
            else:
                seasonality = pd.Series(1.0, index=range(1, 13))
                seasonal_factors = np.ones(len(periods))

            trend_seasonal = trend_projection * seasonal_factors
            run_rate_seasonal = recent_projection * seasonal_factors

            blended = np.zeros(len(periods))
            for idx in range(len(periods)):
                components = [
                    (trend_seasonal[idx], self.config.trend_weight),
                    (run_rate_seasonal[idx], self.config.recent_run_rate_weight),
                ]
                if budget_available[idx]:
                    components.append((budget_values[idx], self.config.budget_weight))
                else:
                    # Reallocate missing budget weight proportionally.
                    non_budget_weight = (
                        self.config.trend_weight
                        + self.config.recent_run_rate_weight
                    )
                    components = [
                        (trend_seasonal[idx], self.config.trend_weight / non_budget_weight),
                        (run_rate_seasonal[idx], self.config.recent_run_rate_weight / non_budget_weight),
                    ]
                blended[idx] = sum(value * weight for value, weight in components)

            output[line] = np.maximum(blended, 0.0)

            budget_coverage = float(budget_available.mean())
            confidence = min(
                1.0,
                0.35
                + min(len(self.history), 24) / 48
                + 0.25 * budget_coverage,
            )
            diagnostic_rows.append({
                "Line": line,
                "History Months": len(self.history),
                "Estimated Monthly Trend": trend_growth,
                "Recent Run Rate": recent_run_rate,
                "Budget Coverage": budget_coverage,
                "Forecast Confidence": confidence,
                "Method": "Trend + Budget + Recent Run Rate + Seasonality",
            })

            for month, factor in seasonality.items():
                seasonality_rows.append({
                    "Line": line,
                    "Month": int(month),
                    "Seasonality Index": float(factor),
                })

        # Preserve realistic P&L structure when history/budget inputs are inconsistent.
        output["COGS"] = np.minimum(output["COGS"], output["Revenue"] * 1.5)
        output["Gross Profit"] = output["Revenue"] - output["COGS"]
        output["EBITDA"] = (
            output["Gross Profit"]
            - output["Payroll"]
            - output["Other Opex"]
        )

        return ForecastBuildResult(
            forecast=output,
            diagnostics=pd.DataFrame(diagnostic_rows),
            seasonality_indices=pd.DataFrame(seasonality_rows),
        )
