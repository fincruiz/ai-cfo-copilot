from datetime import date
from decimal import Decimal
from uuid import UUID

from dateutil.relativedelta import relativedelta

from app.domain.finance.forecasting import (
    build_run_rate_forecast,
    build_trend_forecast,
)
from app.repositories.finance.gl_transaction_repository import (
    GLTransactionRepository,
)


class ForecastingService:
    def __init__(self, repository: GLTransactionRepository) -> None:
        self.repository = repository

    async def forecast(
        self,
        *,
        company_id: UUID,
        reporting_group: str,
        future_months: int,
        method: str,
        branch_id: UUID | None,
        downside_factor: Decimal,
        upside_factor: Decimal,
        recent_months: int,
    ):
        rows = await self.repository.monthly_group_totals(
            company_id=company_id,
            reporting_group=reporting_group,
            branch_id=branch_id,
        )
        history = [
            (
                row.month.date().isoformat()
                if hasattr(row.month, "date")
                else str(row.month),
                Decimal(row.net or 0),
            )
            for row in rows
        ]
        if not history:
            raise ValueError(
                "No mapped monthly history is available for this forecast."
            )

        last_month = date.fromisoformat(history[-1][0]).replace(day=1)
        future_periods = [
            (last_month + relativedelta(months=index)).isoformat()
            for index in range(1, future_months + 1)
        ]

        normalized_method = method.strip().lower()
        if normalized_method == "trend":
            points = build_trend_forecast(
                history,
                future_periods,
                downside_factor=downside_factor,
                upside_factor=upside_factor,
            )
        elif normalized_method == "run_rate":
            points = build_run_rate_forecast(
                history,
                future_periods,
                downside_factor=downside_factor,
                upside_factor=upside_factor,
                recent_months=recent_months,
            )
        else:
            raise ValueError(
                "Forecast method must be 'run_rate' or 'trend'."
            )

        count = len(history)
        confidence = "high" if count >= 12 else "medium" if count >= 6 else "low"
        warning = (
            None
            if count >= 12
            else (
                "Limited history: forecasts are less reliable with fewer "
                "than 12 complete months."
            )
        )
        return history, points, normalized_method, confidence, warning
