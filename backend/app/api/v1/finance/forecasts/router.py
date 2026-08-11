from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.schemas.finance.forecasts import (
    ForecastPointResponse,
    ForecastRequest,
    ForecastResponse,
)
from app.schemas.responses import APIResponse
from app.services.finance.forecasting_service import ForecastingService

router = APIRouter(prefix="/forecasts", tags=["Finance Forecasting"])


def get_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> ForecastingService:
    return ForecastingService(GLTransactionRepository(session))


@router.post("", response_model=APIResponse[ForecastResponse])
async def create_forecast(
    request: ForecastRequest,
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[ForecastingService, Depends(get_service)],
):
    try:
        history, points, method, confidence, warning = await service.forecast(
            company_id=current_company.id,
            reporting_group=request.reporting_group,
            future_months=request.future_months,
            method=request.method,
            branch_id=request.branch_id,
            downside_factor=request.downside_factor,
            upside_factor=request.upside_factor,
            recent_months=request.recent_months,
        )
    except ValueError as exc:
        raise ApplicationError(
            message=str(exc),
            error_code="INVALID_FORECAST_REQUEST",
            status_code=422,
        ) from exc

    return APIResponse(
        message="Forecast generated successfully.",
        data=ForecastResponse(
            reporting_group=request.reporting_group,
            method=method,
            branch_id=request.branch_id,
            history_periods=len(history),
            confidence=confidence,
            warning=warning,
            points=[
                ForecastPointResponse(**point.__dict__)
                for point in points
            ],
        ),
    )
