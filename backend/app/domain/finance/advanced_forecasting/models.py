from __future__ import annotations

from dataclasses import dataclass, field, asdict
from typing import Dict, List, Optional
import pandas as pd


@dataclass
class CompanyProfile:
    company_name: str = "Demo Company"
    industry: str = "Professional Services"
    country: str = "Australia"
    currency: str = "AUD"
    reporting_month: str = "2026-06-30"
    business_model: str = "B2B services"
    strategy_summary: str = ""
    board_audience: str = "Board of Directors"


@dataclass
class OpeningBalanceSheet:
    cash: float = 250_000.0
    accounts_receivable: float = 300_000.0
    inventory: float = 180_000.0
    other_current_assets: float = 40_000.0
    gross_ppe: float = 600_000.0
    accumulated_depreciation: float = -140_000.0
    other_non_current_assets: float = 25_000.0

    accounts_payable: float = 210_000.0
    accrued_expenses: float = 75_000.0
    other_current_liabilities: float = 35_000.0
    debt_current: float = 60_000.0
    debt_non_current: float = 420_000.0
    other_non_current_liabilities: float = 20_000.0

    share_capital: float = 200_000.0
    retained_earnings: Optional[float] = None

    def normalized(self) -> "OpeningBalanceSheet":
        if self.retained_earnings is None:
            assets = (
                self.cash + self.accounts_receivable + self.inventory
                + self.other_current_assets + self.gross_ppe
                + self.accumulated_depreciation + self.other_non_current_assets
            )
            liabilities = (
                self.accounts_payable + self.accrued_expenses
                + self.other_current_liabilities + self.debt_current
                + self.debt_non_current + self.other_non_current_liabilities
            )
            self.retained_earnings = assets - liabilities - self.share_capital
        return self


@dataclass
class ForecastDrivers:
    gross_margin: float = 0.42
    payroll_pct_revenue: float = 0.22
    other_opex_pct_revenue: float = 0.14
    annual_interest_rate: float = 0.075
    tax_rate: float = 0.30

    dso_days: float = 42.0
    dpo_days: float = 35.0
    inventory_days: float = 45.0

    capex_pct_revenue: float = 0.04
    useful_life_months: int = 60
    scheduled_debt_repayment: float = 10_000.0
    minimum_cash: float = 100_000.0
    revolver_limit: float = 500_000.0
    dividend_pct_net_income: float = 0.0

    other_current_assets_pct_revenue: float = 0.02
    accrued_expenses_pct_opex: float = 0.15
    other_current_liabilities_pct_revenue: float = 0.01


@dataclass
class ForecastConfig:
    forecast_start: str = "2026-07-01"
    forecast_months: int = 24
    history_months_for_trend: int = 12

    trend_weight: float = 0.45
    budget_weight: float = 0.35
    recent_run_rate_weight: float = 0.20

    recent_run_rate_months: int = 3
    seasonality_enabled: bool = True
    outlier_winsor_limit: float = 0.15
    max_monthly_growth: float = 0.15
    min_monthly_growth: float = -0.15

    opening_balance_sheet: OpeningBalanceSheet = field(default_factory=OpeningBalanceSheet)
    drivers: ForecastDrivers = field(default_factory=ForecastDrivers)

    def validate(self) -> None:
        if self.forecast_months < 1 or self.forecast_months > 120:
            raise ValueError("forecast_months must be between 1 and 120")
        weights = self.trend_weight + self.budget_weight + self.recent_run_rate_weight
        if abs(weights - 1.0) > 0.0001:
            raise ValueError("trend, budget and run-rate weights must total 1.0")
        if self.recent_run_rate_months < 1:
            raise ValueError("recent_run_rate_months must be at least 1")
        if not 0 <= self.drivers.gross_margin <= 1:
            raise ValueError("gross_margin must be between 0 and 1")


@dataclass
class ScenarioDefinition:
    name: str
    description: str = ""
    revenue_multiplier: float = 1.0
    gross_margin_delta: float = 0.0
    dso_delta_days: float = 0.0
    dpo_delta_days: float = 0.0
    inventory_delta_days: float = 0.0
    payroll_pct_delta: float = 0.0
    other_opex_pct_delta: float = 0.0
    capex_multiplier: float = 1.0
    interest_rate_delta: float = 0.0


@dataclass
class HistoricalData:
    monthly: pd.DataFrame

    REQUIRED_COLUMNS = [
        "Period", "Revenue", "COGS", "Payroll", "Other Opex",
    ]

    def validated(self) -> pd.DataFrame:
        missing = [c for c in self.REQUIRED_COLUMNS if c not in self.monthly.columns]
        if missing:
            raise ValueError(f"Historical data is missing columns: {missing}")
        df = self.monthly.copy()
        df["Period"] = pd.to_datetime(df["Period"])
        df = df.sort_values("Period").drop_duplicates("Period", keep="last")
        numeric_cols = [c for c in df.columns if c != "Period"]
        for col in numeric_cols:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
        return df


@dataclass
class BenchmarkData:
    metrics: Dict[str, float] = field(default_factory=dict)
    source_notes: Dict[str, str] = field(default_factory=dict)
    research_summary: str = ""
    industry_trends: List[str] = field(default_factory=list)
    macro_environment: List[str] = field(default_factory=list)
    competitor_observations: List[str] = field(default_factory=list)

    def to_frame(self) -> pd.DataFrame:
        rows = []
        for metric, value in self.metrics.items():
            rows.append({
                "Metric": metric,
                "Benchmark": value,
                "Source": self.source_notes.get(metric, ""),
            })
        return pd.DataFrame(rows)
