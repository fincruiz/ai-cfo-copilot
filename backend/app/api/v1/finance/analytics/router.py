from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.finance.imports import (
    AnalyticsOverviewResponse,
    WorkingCapitalSummaryResponse,
)
from app.schemas.responses import APIResponse
from app.services.finance.analytics_service import AnalyticsService

router = APIRouter(prefix="/analytics", tags=["Finance Analytics"])


def get_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> AnalyticsService:
    return AnalyticsService(session)


@router.get("/overview", response_model=APIResponse[AnalyticsOverviewResponse])
async def overview(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[AnalyticsService, Depends(get_service)],
):
    data = await service.overview(current_company.id)
    return APIResponse(
        message="Analytics overview retrieved.",
        data=AnalyticsOverviewResponse(**data),
    )


@router.get(
    "/working-capital/{ageing_type}",
    response_model=APIResponse[WorkingCapitalSummaryResponse | None],
)
async def working_capital(
    ageing_type: str,
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[AnalyticsService, Depends(get_service)],
):
    kind = ageing_type.upper()
    if kind not in {"AR", "AP"}:
        raise ValueError("Ageing type must be AR or AP.")
    return APIResponse(
        message=f"{kind} ageing analysis retrieved.",
        data=await service.working_capital_summary(current_company.id, kind),
    )
