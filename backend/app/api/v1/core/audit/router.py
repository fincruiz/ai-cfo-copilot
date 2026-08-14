from typing import Annotated
from fastapi import APIRouter, Depends, Query
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.responses import APIResponse
from app.services.audit_service import AuditService
router=APIRouter(prefix="/audit", tags=["Audit Trail"])
@router.get("/events")
async def audit_events(current_company: Annotated[Company, Depends(get_current_company)], session: Annotated[AsyncSession, Depends(get_db_session)], limit: int = Query(100, ge=1, le=250)):
    return APIResponse(message="Audit events retrieved.", data=await AuditService(session).list(company_id=current_company.id, limit=limit))
