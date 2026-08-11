from typing import Annotated
from uuid import UUID
from fastapi import APIRouter,Depends
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.responses import APIResponse
from app.schemas.finance.advanced_forecasting import PlanningVersionCreate,NativePlanLineInput
from app.services.finance.native_planning_service import NativePlanningService
router=APIRouter(prefix='/native-planning',tags=['Native Planning'])
def svc(session:Annotated[AsyncSession,Depends(get_db_session)]):return NativePlanningService(session)
@router.get('/versions')
async def versions(current_company:Annotated[Company,Depends(get_current_company)],service:Annotated[NativePlanningService,Depends(svc)]):return APIResponse(message='Planning versions retrieved.',data=await service.list_versions(current_company.id))
@router.post('/versions')
async def create(request:PlanningVersionCreate,current_company:Annotated[Company,Depends(get_current_company)],service:Annotated[NativePlanningService,Depends(svc)]):return APIResponse(message='Planning version created.',data=await service.create_version(current_company.id,request))
@router.get('/versions/{version_id}')
async def get(version_id:UUID,current_company:Annotated[Company,Depends(get_current_company)],service:Annotated[NativePlanningService,Depends(svc)]):return APIResponse(message='Planning version retrieved.',data=await service.get_version(current_company.id,version_id))
@router.put('/versions/{version_id}/lines')
async def save(version_id:UUID,lines:list[NativePlanLineInput],current_company:Annotated[Company,Depends(get_current_company)],service:Annotated[NativePlanningService,Depends(svc)]):return APIResponse(message='Planning lines saved.',data=await service.save_lines(current_company.id,version_id,lines))
