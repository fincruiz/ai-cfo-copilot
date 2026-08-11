from typing import Annotated
from fastapi import APIRouter,Depends,status
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.repositories.finance.account_mapping_repository import AccountMappingRepository
from app.schemas.finance.mappings import *
from app.schemas.responses import APIResponse
from app.services.finance.mapping_service import MappingService
router=APIRouter(prefix="/account-mappings",tags=["Finance Mapping"])
def service(session:Annotated[AsyncSession,Depends(get_db_session)]):return MappingService(AccountMappingRepository(session))
@router.get("",response_model=APIResponse[list[AccountMappingResponse]])
async def list_mappings(current_company:Annotated[Company,Depends(get_current_company)],svc:Annotated[MappingService,Depends(service)]):return APIResponse(message="Mappings retrieved.",data=[AccountMappingResponse.model_validate(x) for x in await svc.list_mappings(current_company.id)])
@router.get("/suggestions",response_model=APIResponse[list[MappingSuggestionResponse]])
async def suggestions(current_company:Annotated[Company,Depends(get_current_company)],svc:Annotated[MappingService,Depends(service)]):return APIResponse(message="Mapping suggestions generated.",data=[MappingSuggestionResponse(**x) for x in await svc.suggest_unmapped(current_company.id)])
@router.put("",response_model=APIResponse[dict],status_code=status.HTTP_200_OK)
async def upsert(request:MappingBulkRequest,current_company:Annotated[Company,Depends(get_current_company)],svc:Annotated[MappingService,Depends(service)]):count=await svc.upsert(current_company.id,request.items);return APIResponse(message="Mappings saved.",data={"saved":count})
