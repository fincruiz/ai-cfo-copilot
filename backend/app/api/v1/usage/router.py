from typing import Annotated

from fastapi import APIRouter, Depends, Query
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company, require_company_admin
from app.schemas.auth import CurrentUser
from app.schemas.responses import APIResponse
from app.schemas.usage import UsageEventCreate
from app.services.usage_service import UsageService

router = APIRouter(prefix="/usage", tags=["Product Usage"])


@router.post("/events")
async def record_usage_event(
    payload: UsageEventCreate,
    current_user: Annotated[CurrentUser, Depends(get_current_user)],
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    await UsageService(session).record(
        company_id=current_company.id,
        user_id=current_user.id,
        event_name=payload.event_name,
        path=payload.path,
        session_id=payload.session_id,
        properties=payload.properties,
    )
    return APIResponse(message="Usage event accepted.", data={"recorded": True})


@router.get("/summary")
async def usage_summary(
    current_company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
    days: int = Query(30, ge=1, le=365),
):
    return APIResponse(
        message="Usage summary retrieved.",
        data=await UsageService(session).summary(company_id=current_company.id, days=days),
    )
