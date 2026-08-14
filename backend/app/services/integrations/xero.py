from __future__ import annotations
from urllib.parse import urlencode
from uuid import UUID
import httpx
from app.core.config import settings
from app.services.integrations.base import IntegrationStore

AUTH='https://login.xero.com/identity/connect/authorize'
TOKEN='https://identity.xero.com/connect/token'
API='https://api.xero.com/api.xro/2.0'
CONNECTIONS='https://api.xero.com/connections'
SCOPES='offline_access accounting.transactions accounting.settings accounting.contacts accounting.reports.read'

class XeroConnector:
    def __init__(self, store: IntegrationStore): self.store=store
    def configured(self): return bool(settings.xero_client_id and settings.xero_client_secret and settings.xero_redirect_uri)
    async def authorization_url(self, company_id: UUID, user_id: UUID):
        state=await self.store.save_oauth_state(company_id,user_id,'xero')
        return AUTH+'?'+urlencode({'response_type':'code','client_id':settings.xero_client_id,'redirect_uri':settings.xero_redirect_uri,'scope':SCOPES,'state':state})
    async def callback(self, code: str, state: str):
        ctx=await self.store.consume_oauth_state(state,'xero')
        if not ctx: raise ValueError('Invalid or expired Xero OAuth state.')
        async with httpx.AsyncClient(timeout=30) as client:
            tok=await client.post(TOKEN,data={'grant_type':'authorization_code','code':code,'redirect_uri':settings.xero_redirect_uri},auth=(settings.xero_client_id or '',settings.xero_client_secret or '')); tok.raise_for_status(); t=tok.json()
            con=await client.get(CONNECTIONS,headers={'Authorization':f"Bearer {t['access_token']}"}); con.raise_for_status(); tenants=con.json()
        chosen=tenants[0] if len(tenants)==1 else None
        await self.store.upsert_connection(company_id=ctx['company_id'],provider='xero',user_id=ctx['user_id'],status='connected' if chosen else 'selection_required',external_tenant_id=chosen.get('tenantId') if chosen else None,external_tenant_name=chosen.get('tenantName') if chosen else None,access_token=t['access_token'],refresh_token=t.get('refresh_token'),expires_in=t.get('expires_in'),metadata={'tenants':tenants})
        return ctx['company_id']
    async def select_tenant(self, company_id: UUID, tenant_id: str):
        c=await self.store.get(company_id,'xero'); tenants=(c or {}).get('metadata',{}).get('tenants',[])
        chosen=next((x for x in tenants if x.get('tenantId')==tenant_id),None)
        if not chosen: raise ValueError('That Xero organisation is not available for this connection.')
        await self.store.upsert_connection(company_id=company_id,provider='xero',status='connected',external_tenant_id=tenant_id,external_tenant_name=chosen.get('tenantName'))
    async def _refresh(self, company_id: UUID, c: dict):
        if not c.get('refresh_token'): return c
        async with httpx.AsyncClient(timeout=30) as client:
            r=await client.post(TOKEN,data={'grant_type':'refresh_token','refresh_token':c['refresh_token']},auth=(settings.xero_client_id or '',settings.xero_client_secret or '')); r.raise_for_status(); t=r.json()
        await self.store.upsert_connection(company_id=company_id,provider='xero',access_token=t['access_token'],refresh_token=t.get('refresh_token'),expires_in=t.get('expires_in'))
        return await self.store.credentials(company_id,'xero')
    async def sync(self, company_id: UUID):
        c=await self.store.credentials(company_id,'xero')
        if not c or not c.get('external_tenant_id'): raise ValueError('Connect and select a Xero organisation first.')
        c=await self._refresh(company_id,c)
        h={'Authorization':f"Bearer {c['access_token']}",'Xero-tenant-id':c['external_tenant_id'],'Accept':'application/json'}
        endpoints={'account':'Accounts','contact':'Contacts','invoice':'Invoices','bank_transaction':'BankTransactions'}; counts={}
        try:
            async with httpx.AsyncClient(timeout=45) as client:
                for entity, endpoint in endpoints.items():
                    r=await client.get(f'{API}/{endpoint}',headers=h); r.raise_for_status(); body=r.json(); key=endpoint
                    items=body.get(key,[]); normalized=[]
                    for x in items:
                        xid=x.get('AccountID') or x.get('ContactID') or x.get('InvoiceID') or x.get('BankTransactionID')
                        if not xid: continue
                        normalized.append({'external_id':xid,'name':x.get('Name') or x.get('Contact',{}).get('Name') or x.get('InvoiceNumber') or x.get('Reference'),'amount':x.get('Total') or x.get('Balance'),'currency_code':x.get('CurrencyCode'),'payload':x})
                    await self.store.replace_records(company_id,'xero',entity,normalized); counts[entity]=len(normalized)
            await self.store.mark_sync(company_id,'xero','success',f"Synced {sum(counts.values())} Xero records.")
            return counts
        except Exception as e:
            await self.store.mark_sync(company_id,'xero','failed',str(e)); raise
