from __future__ import annotations

import numpy as np
import pandas as pd


def make_demo_history(months: int = 24) -> pd.DataFrame:
    rng = np.random.default_rng(42)
    periods = pd.date_range("2024-07-01", periods=months, freq="MS")
    seasonality = np.array([0.90, 0.92, 0.97, 1.00, 1.04, 1.10, 1.08, 1.03, 0.98, 0.96, 1.02, 1.15])
    revenue = []
    base = 420_000.0
    for idx, period in enumerate(periods):
        trend = base * (1.012 ** idx)
        noise = rng.normal(1.0, 0.025)
        revenue.append(trend * seasonality[period.month - 1] * noise)
    revenue = np.array(revenue)
    gross_margin = 0.40 + rng.normal(0, 0.008, len(periods))
    cogs = revenue * (1 - gross_margin)
    payroll = revenue * (0.215 + rng.normal(0, 0.006, len(periods)))
    other_opex = revenue * (0.14 + rng.normal(0, 0.005, len(periods)))

    return pd.DataFrame({
        "Period": periods,
        "Revenue": revenue,
        "COGS": cogs,
        "Payroll": payroll,
        "Other Opex": other_opex,
    })


def make_demo_budget(start: str = "2026-07-01", months: int = 24) -> pd.DataFrame:
    periods = pd.date_range(start, periods=months, freq="MS")
    revenue = np.array([565_000 * (1.014 ** i) for i in range(months)])
    return pd.DataFrame({
        "Period": periods,
        "Revenue": revenue,
        "COGS": revenue * 0.57,
        "Payroll": revenue * 0.21,
        "Other Opex": revenue * 0.135,
    })
