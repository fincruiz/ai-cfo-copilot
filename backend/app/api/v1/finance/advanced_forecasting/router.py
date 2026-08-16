from typing import Annotated
from uuid import UUID
from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company, require_finance_write
from app.schemas.responses import APIResponse
from app.schemas.finance.advanced_forecasting import AdvancedForecastRequest,PowerOfOneRequest,ForecastRunResponse,PowerOfOneResponse
from app.services.finance.advanced_forecasting_service import AdvancedForecastingService
router=APIRouter(prefix='/advanced-forecast',tags=['Advanced Forecasting'])
def svc(session:Annotated[AsyncSession,Depends(get_db_session)]):return AdvancedForecastingService(session)
@router.post('/run',response_model=APIResponse[ForecastRunResponse])
async def run(request:AdvancedForecastRequest,current_company:Annotated[Company,Depends(get_current_company)],_membership:Annotated[object,Depends(require_finance_write)],service:Annotated[AdvancedForecastingService,Depends(svc)]):
    return APIResponse(message='Integrated three-way forecast completed.',data=ForecastRunResponse(**await service.calculate(current_company.id,request)))
@router.post('/power-of-one',response_model=APIResponse[PowerOfOneResponse])
async def power(request:PowerOfOneRequest,current_company:Annotated[Company,Depends(get_current_company)],_membership:Annotated[object,Depends(require_finance_write)],service:Annotated[AdvancedForecastingService,Depends(svc)]):
    return APIResponse(message='Power of One impact calculated.',data=PowerOfOneResponse(**await service.power_of_one(current_company.id,request)))
