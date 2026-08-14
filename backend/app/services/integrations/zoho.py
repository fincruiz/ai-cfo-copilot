from __future__ import annotations
from urllib.parse import urlencode
from uuid import UUID
import httpx
from app.core.config import settings
from app.services.integrations.base import IntegrationStore

SCOPES='ZohoBooks.fullaccess.all'
class ZohoConnector:
    def __init__(self, store: IntegrationStore): self.store=store
    def configured(self): return bool(settings.zoho_client_id and settings.zoho_client_secret and settings.zoho_redirect_uri)
    async def authorization_url(self, company_id: UUID, user_id: UUID):
        state=await self.store.save_oauth_state(company_id,user_id,'zoho')
        return settings.zoho_accounts_base_url+'/oauth/v2/auth?'+urlencode({'scope':SCOPES,'client_id':settings.zoho_client_id,'response_type':'code','access_type':'offline','prompt':'consent','redirect_uri':settings.zoho_redirect_uri,'state':state})
    async def callback(self, code: str, state: str):
        ctx=await self.store.consume_oauth_state(state,'zoho')
        if not ctx: raise ValueError('Invalid or expired Zoho OAuth state.')
        async with httpx.AsyncClient(timeout=30) as client:
            r=await client.post(settings.zoho_accounts_base_url+'/oauth/v2/token',params={'grant_type':'authorization_code','client_id':settings.zoho_client_id,'client_secret':settings.zoho_client_secret,'redirect_uri':settings.zoho_redirect_uri,'code':code}); r.raise_for_status(); t=r.json()
            org=await client.get(settings.zoho_api_base_url+'/organizations',headers={'Authorization':f"Zoho-oauthtoken {t['access_token']}"}); org.raise_for_status(); orgs=org.json().get('organizations',[])
        chosen=orgs[0] if len(orgs)==1 else None
        await self.store.upsert_connection(company_id=ctx['company_id'],provider='zoho',user_id=ctx['user_id'],status='connected' if chosen else 'selection_required',external_tenant_id=str(chosen.get('organization_id')) if chosen else None,external_tenant_name=chosen.get('name') if chosen else None,access_token=t['access_token'],refresh_token=t.get('refresh_token'),expires_in=t.get('expires_in_sec') or t.get('expires_in'),metadata={'organizations':orgs})
        return ctx['company_id']
    async def select_tenant(self, company_id: UUID, tenant_id: str):
        c=await self.store.get(company_id,'zoho'); orgs=(c or {}).get('metadata',{}).get('organizations',[]); chosen=next((x for x in orgs if str(x.get('organization_id'))==tenant_id),None)
        if not chosen: raise ValueError('That Zoho Books organisation is not available.')
        await self.store.upsert_connection(company_id=company_id,provider='zoho',status='connected',external_tenant_id=tenant_id,external_tenant_name=chosen.get('name'))
    async def _refresh(self, company_id: UUID, c: dict):
        if not c.get('refresh_token'): return c
        async with httpx.AsyncClient(timeout=30) as client:
            r=await client.post(settings.zoho_accounts_base_url+'/oauth/v2/token',params={'grant_type':'refresh_token','client_id':settings.zoho_client_id,'client_secret':settings.zoho_client_secret,'refresh_token':c['refresh_token']}); r.raise_for_status(); t=r.json()
        await self.store.upsert_connection(company_id=company_id,provider='zoho',access_token=t['access_token'],expires_in=t.get('expires_in_sec') or t.get('expires_in'))
        return await self.store.credentials(company_id,'zoho')
    async def sync(self, company_id: UUID):
        c=await self.store.credentials(company_id,'zoho')
        if not c or not c.get('external_tenant_id'): raise ValueError('Connect and select a Zoho Books organisation first.')
        c=await self._refresh(company_id,c); h={'Authorization':f"Zoho-oauthtoken {c['access_token']}"}; params={'organization_id':c['external_tenant_id'],'per_page':200}; endpoints={'account':'chartofaccounts','contact':'contacts','invoice':'invoices','bill':'bills'}; keys={'account':'chartofaccounts','contact':'contacts','invoice':'invoices','bill':'bills'}; counts={}
        try:
            async with httpx.AsyncClient(timeout=45) as client:
                for entity, endpoint in endpoints.items():
                    r=await client.get(f"{settings.zoho_api_base_url}/{endpoint}",headers=h,params=params); r.raise_for_status(); items=r.json().get(keys[entity],[]); normalized=[]
                    for x in items:
                        xid=x.get('account_id') or x.get('contact_id') or x.get('invoice_id') or x.get('bill_id')
                        if not xid: continue
                        normalized.append({'external_id':xid,'name':x.get('account_name') or x.get('contact_name') or x.get('invoice_number') or x.get('bill_number') or x.get('vendor_name'),'amount':x.get('total') or x.get('balance'),'currency_code':x.get('currency_code'),'payload':x})
                    await self.store.replace_records(company_id,'zoho',entity,normalized); counts[entity]=len(normalized)
            await self.store.mark_sync(company_id,'zoho','success',f"Synced {sum(counts.values())} Zoho Books records."); return counts
        except Exception as e:
            await self.store.mark_sync(company_id,'zoho','failed',str(e)); raise
