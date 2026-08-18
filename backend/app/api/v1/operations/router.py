from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.operations import OperationalReadiness
from app.schemas.responses import APIResponse
from app.services.operations_service import OperationsService

router = APIRouter(prefix="/operations", tags=["Operations"])


@router.get("/readiness", response_model=APIResponse[OperationalReadiness])
async def operational_readiness(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    data = await OperationsService(session).readiness(company.id)
    return APIResponse(message="Operational readiness retrieved.", data=OperationalReadiness(**data))
