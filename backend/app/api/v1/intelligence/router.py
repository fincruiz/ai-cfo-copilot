from typing import Annotated
from fastapi import APIRouter,Depends
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company,require_finance_write
from app.schemas.auth import CurrentUser
from app.schemas.integrations import MemoryCreate
from app.schemas.responses import APIResponse
from app.services.intelligence.brain_service import BrainService
router=APIRouter(prefix='/intelligence',tags=['Organizational Intelligence'])
@router.get('/overview')
async def overview(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)]): return APIResponse(message='Organizational intelligence retrieved.',data=await BrainService(session).overview(company.id))
@router.post('/memory')
async def memory(payload:MemoryCreate,company:Annotated[Company,Depends(get_current_company)],user:Annotated[CurrentUser,Depends(get_current_user)],_:Annotated[object,Depends(require_finance_write)],session:Annotated[AsyncSession,Depends(get_db_session)]): return APIResponse(message='Management context remembered.',data=await BrainService(session).add_memory(company.id,user.id,payload))
