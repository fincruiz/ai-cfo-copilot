from typing import Annotated
from fastapi import APIRouter,Depends
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company,require_company_admin
from app.schemas.responses import APIResponse
from app.schemas.subscription import BetaReadinessCheck,BetaReadinessOut,SubscriptionStatusOut,SubscriptionChangeRequest,SubscriptionChangeOut,BillingMarketRequest
from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService
from app.services.integrations.base import IntegrationStore
from app.services.subscription_service import SubscriptionService
router=APIRouter(prefix='/subscription',tags=['Subscription & Beta'])
@router.get('/status',response_model=APIResponse[SubscriptionStatusOut])
async def subscription_status(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)]):return APIResponse(message='Subscription status retrieved.',data=SubscriptionStatusOut(**await SubscriptionService(session).status(company_id=company.id)))
@router.post('/change-request',response_model=APIResponse[SubscriptionChangeOut])
async def change_request(request:SubscriptionChangeRequest,company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    data=await SubscriptionService(session).request_change(company_id=company.id,plan=request.plan,interval=request.billing_interval);return APIResponse(message=data['message'],data=SubscriptionChangeOut(**data))
@router.post('/cancel-request',response_model=APIResponse[SubscriptionStatusOut])
async def cancel_request(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    return APIResponse(message='Cancellation request recorded.',data=SubscriptionStatusOut(**await SubscriptionService(session).request_cancellation(company_id=company.id)))
@router.put('/billing-market',response_model=APIResponse[SubscriptionStatusOut])
async def billing_market(request:BillingMarketRequest,company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    return APIResponse(message='Billing market updated.',data=SubscriptionStatusOut(**await SubscriptionService(session).update_billing_market(company_id=company.id,country_code=request.country_code)))
@router.get('/beta-readiness',response_model=APIResponse[BetaReadinessOut])
async def beta_readiness(company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    workspace=await WorkspaceLifecycleService(session).status(company_id=company.id);subscription=await SubscriptionService(session).status(company_id=company.id);connections=await IntegrationStore(session).list_connections(company.id);connected=[x for x in connections if x.get('status')=='connected'];successful=[x for x in connected if x.get('last_sync_status') in (None,'success')]
    checks=[BetaReadinessCheck(key='subscription',label='Plan access',status='ready' if subscription['is_access_active'] else 'blocked',detail=f"{subscription['display_name']} · {subscription['status'].replace('_',' ')}"),BetaReadinessCheck(key='financial_data',label='Financial data',status='ready' if workspace.get('has_financial_data') else 'attention',detail=f"{workspace.get('transaction_count',0):,} ledger transactions available." if workspace.get('has_financial_data') else 'Load company data or demo data.'),BetaReadinessCheck(key='mapping',label='Account mapping',status='ready' if workspace.get('mapping_count',0)>0 else 'attention',detail=f"{workspace.get('mapping_count',0)} confirmed mapping(s)."),BetaReadinessCheck(key='integration',label='Connected source',status='ready' if successful else 'attention',detail=f"{len(successful)} healthy connected source(s)." if successful else 'Uploads can be used while live integrations are being configured.')]
    weights={'ready':25,'attention':12,'blocked':0};score=min(100,sum(weights[x.status] for x in checks));overall='ready' if score>=85 else 'attention' if score>=50 else 'blocked';return APIResponse(message='Beta readiness retrieved.',data=BetaReadinessOut(score=score,status=overall,checks=checks))
