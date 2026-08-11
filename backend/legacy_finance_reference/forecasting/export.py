from __future__ import annotations

from io import BytesIO
from typing import Dict
import pandas as pd

from .three_way import ThreeWayForecastResult


def forecast_to_excel(
    result: ThreeWayForecastResult,
    diagnostics: pd.DataFrame,
    seasonality: pd.DataFrame,
    scenarios: pd.DataFrame,
) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        result.forecast_basis.to_excel(writer, sheet_name="Forecast Basis")
        result.profit_and_loss.to_excel(writer, sheet_name="P&L")
        result.balance_sheet.to_excel(writer, sheet_name="Balance Sheet")
        result.cash_flow.to_excel(writer, sheet_name="Cash Flow")
        result.ratios.to_excel(writer, sheet_name="Ratios")
        result.checks.to_excel(writer, sheet_name="Checks")
        diagnostics.to_excel(writer, sheet_name="Forecast Diagnostics", index=False)
        seasonality.to_excel(writer, sheet_name="Seasonality", index=False)
        scenarios.to_excel(writer, sheet_name="Scenarios")
        for name, frame in result.schedules.items():
            frame.to_excel(writer, sheet_name=name[:31])
    return buffer.getvalue()
