from __future__ import annotations
from typing import Annotated
from fastapi import APIRouter, Depends, Header, HTTPException
from fastapi.responses import RedirectResponse
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.config import settings
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company, require_company_admin, require_finance_write
from app.schemas.auth import CurrentUser
from app.schemas.integrations import TenantSelection, TallyPushRequest
from app.schemas.responses import APIResponse
from app.services.audit_service import AuditService
from app.services.integrations.base import IntegrationStore
from app.services.integrations.xero import XeroConnector
from app.services.integrations.zoho import ZohoConnector

router=APIRouter(prefix='/integrations',tags=['Integrations'])

def _front(status: str, provider: str): return f"{settings.integration_frontend_url.rstrip('/')}/dashboard/integrations?provider={provider}&status={status}"

@router.get('')
async def list_integrations(company: Annotated[Company,Depends(get_current_company)], session: Annotated[AsyncSession,Depends(get_db_session)]):
    existing={x['provider']:x for x in await IntegrationStore(session).list_connections(company.id)}
    data=[]
    for provider, configured in [('xero',bool(settings.xero_client_id and settings.xero_client_secret)),('zoho',bool(settings.zoho_client_id and settings.zoho_client_secret)),('tally',True)]:
        data.append(existing.get(provider) or {'provider':provider,'status':'disconnected','configured':configured,'metadata':{}})
        data[-1]['configured']=configured
    return APIResponse(message='Integrations retrieved.',data=data)

@router.post('/xero/start')
async def xero_start(company: Annotated[Company,Depends(get_current_company)], user: Annotated[CurrentUser,Depends(get_current_user)], _: Annotated[object,Depends(require_company_admin)], session: Annotated[AsyncSession,Depends(get_db_session)]):
    c=XeroConnector(IntegrationStore(session));
    if not c.configured(): raise HTTPException(503,'Xero credentials are not configured on the server.')
    return APIResponse(message='Xero authorization ready.',data={'authorization_url':await c.authorization_url(company.id,user.id)})

@router.get('/xero/callback')
async def xero_callback(code: str,state: str,session: Annotated[AsyncSession,Depends(get_db_session)]):
    try: await XeroConnector(IntegrationStore(session)).callback(code,state); return RedirectResponse(_front('connected','xero'))
    except Exception: return RedirectResponse(_front('error','xero'))

@router.post('/xero/select-tenant')
async def xero_select(payload: TenantSelection,company: Annotated[Company,Depends(get_current_company)],_: Annotated[object,Depends(require_company_admin)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    await XeroConnector(IntegrationStore(session)).select_tenant(company.id,payload.tenant_id); return APIResponse(message='Xero organisation selected.',data=True)

@router.post('/xero/sync')
async def xero_sync(company: Annotated[Company,Depends(get_current_company)],user: Annotated[CurrentUser,Depends(get_current_user)],_: Annotated[object,Depends(require_finance_write)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    counts=await XeroConnector(IntegrationStore(session)).sync(company.id); await AuditService(session).record(company_id=company.id,user_id=user.id,action='sync',module='integrations',summary='Synced Xero data.',metadata=counts,commit=True); return APIResponse(message='Xero sync complete.',data=counts)

@router.post('/zoho/start')
async def zoho_start(company: Annotated[Company,Depends(get_current_company)],user: Annotated[CurrentUser,Depends(get_current_user)],_: Annotated[object,Depends(require_company_admin)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    c=ZohoConnector(IntegrationStore(session));
    if not c.configured(): raise HTTPException(503,'Zoho credentials are not configured on the server.')
    return APIResponse(message='Zoho authorization ready.',data={'authorization_url':await c.authorization_url(company.id,user.id)})

@router.get('/zoho/callback')
async def zoho_callback(code: str,state: str,session: Annotated[AsyncSession,Depends(get_db_session)]):
    try: await ZohoConnector(IntegrationStore(session)).callback(code,state); return RedirectResponse(_front('connected','zoho'))
    except Exception: return RedirectResponse(_front('error','zoho'))

@router.post('/zoho/select-tenant')
async def zoho_select(payload: TenantSelection,company: Annotated[Company,Depends(get_current_company)],_: Annotated[object,Depends(require_company_admin)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    await ZohoConnector(IntegrationStore(session)).select_tenant(company.id,payload.tenant_id); return APIResponse(message='Zoho Books organisation selected.',data=True)

@router.post('/zoho/sync')
async def zoho_sync(company: Annotated[Company,Depends(get_current_company)],user: Annotated[CurrentUser,Depends(get_current_user)],_: Annotated[object,Depends(require_finance_write)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    counts=await ZohoConnector(IntegrationStore(session)).sync(company.id); await AuditService(session).record(company_id=company.id,user_id=user.id,action='sync',module='integrations',summary='Synced Zoho Books data.',metadata=counts,commit=True); return APIResponse(message='Zoho Books sync complete.',data=counts)

@router.post('/tally/bridge-token')
async def tally_token(company: Annotated[Company,Depends(get_current_company)],user: Annotated[CurrentUser,Depends(get_current_user)],_: Annotated[object,Depends(require_company_admin)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    token=await IntegrationStore(session).tally_bridge_token(company.id,user.id); return APIResponse(message='Tally bridge token created. Copy it now; it is shown once.',data={'bridge_token':token,'ingest_url':'/api/v1/integrations/tally/push'})

@router.post('/tally/push')
async def tally_push(payload: TallyPushRequest,authorization: Annotated[str|None,Header()]=None,session: Annotated[AsyncSession,Depends(get_db_session)]=None):
    if not authorization or not authorization.lower().startswith('bearer '): raise HTTPException(401,'Tally bridge token required.')
    store=IntegrationStore(session); company_id=await store.company_for_bridge(authorization.split(' ',1)[1])
    if not company_id: raise HTTPException(401,'Invalid Tally bridge token.')
    grouped={}
    for r in payload.records: grouped.setdefault(r.entity_type,[]).append(r.model_dump())
    for entity, rows in grouped.items(): await store.replace_records(company_id,'tally',entity,rows)
    await store.upsert_connection(company_id=company_id,provider='tally',status='connected'); await store.mark_sync(company_id,'tally','success',f'Received {len(payload.records)} Tally records.')
    return APIResponse(message='Tally data accepted.',data={'records':len(payload.records)})

@router.delete('/{provider}')
async def disconnect(provider: str,company: Annotated[Company,Depends(get_current_company)],user: Annotated[CurrentUser,Depends(get_current_user)],_: Annotated[object,Depends(require_company_admin)],session: Annotated[AsyncSession,Depends(get_db_session)]):
    if provider not in {'xero','zoho','tally'}: raise HTTPException(404,'Unknown provider.')
    await IntegrationStore(session).disconnect(company.id,provider,True); await AuditService(session).record(company_id=company.id,user_id=user.id,action='disconnect',module='integrations',summary=f'Disconnected {provider} and deleted synchronized records.',commit=True); return APIResponse(message=f'{provider.title()} disconnected and synchronized data deleted.',data=True)
