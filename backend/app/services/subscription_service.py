from __future__ import annotations
from datetime import datetime, timezone
from uuid import UUID
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from app.core.exceptions import ApplicationError
from app.services.market_service import PLAN_LABELS, PRICE_CATALOG, resolve_market

PLAN_ENTITLEMENTS={
 'trial':{'ai_queries_monthly':200,'users':3,'integrations':1,'forecasting':True,'decision_simulator':True,'board_packs':True,'benchmarking':True,'audit_history_days':30},
 'founding':{'ai_queries_monthly':2000,'users':10,'integrations':3,'forecasting':True,'decision_simulator':True,'board_packs':True,'benchmarking':True,'audit_history_days':365},
 'growth':{'ai_queries_monthly':5000,'users':25,'integrations':5,'forecasting':True,'decision_simulator':True,'board_packs':True,'benchmarking':True,'audit_history_days':730},
 'enterprise':{'ai_queries_monthly':-1,'users':-1,'integrations':-1,'forecasting':True,'decision_simulator':True,'board_packs':True,'benchmarking':True,'audit_history_days':-1},
}
ACTIVE_STATUSES={'trialing','active'}

def entitlements_for_plan(plan:str,overrides:dict|None=None)->dict:
    base=dict(PLAN_ENTITLEMENTS.get(plan,PLAN_ENTITLEMENTS['trial']))
    for key,value in (overrides or {}).items():
        if key in base: base[key]=value
    return base

def days_remaining(trial_ends_at:datetime|None,now:datetime|None=None)->int|None:
    if trial_ends_at is None:return None
    now=now or datetime.now(timezone.utc)
    if trial_ends_at.tzinfo is None:trial_ends_at=trial_ends_at.replace(tzinfo=timezone.utc)
    return max(0,(trial_ends_at-now).days)

class SubscriptionService:
    def __init__(self,session:AsyncSession):self.session=session
    async def _row(self,company_id:UUID):
        return (await self.session.execute(text('''SELECT s.*,c.country_code FROM public.company_subscriptions s JOIN public.companies c ON c.id=s.company_id WHERE s.company_id=:company_id'''),{'company_id':company_id})).mappings().first()
    async def status(self,*,company_id:UUID)->dict:
        row=await self._row(company_id)
        if not row:
            await self.session.execute(text("INSERT INTO public.company_subscriptions(company_id,plan,status,trial_started_at,trial_ends_at) VALUES (:company_id,'trial','trialing',now(),now()+interval '30 days') ON CONFLICT (company_id) DO NOTHING"),{'company_id':company_id});await self.session.commit();row=await self._row(company_id)
        data=dict(row or {});plan=str(data.get('plan') or 'trial');status=str(data.get('status') or 'trialing');trial_end=data.get('trial_ends_at')
        if status=='trialing' and trial_end and days_remaining(trial_end)==0:status='expired'
        billing_country=str(data.get('billing_country_code') or data.get('country_code') or 'GLOBAL').upper()
        return {'plan':plan,'display_name':'Trial' if plan=='trial' else PLAN_LABELS.get(plan,plan.title()),'status':status,'trial_started_at':data.get('trial_started_at'),'trial_ends_at':trial_end,'current_period_ends_at':data.get('current_period_ends_at'),'days_remaining':days_remaining(trial_end),'entitlements':entitlements_for_plan(plan,data.get('entitlements') or {}),'is_access_active':status in ACTIVE_STATUSES,'billing_managed_externally':True,'billing_country_code':billing_country,'billing_interval':data.get('billing_interval') or 'monthly','requested_plan':data.get('requested_plan'),'requested_interval':data.get('requested_interval'),'change_requested_at':data.get('change_requested_at'),'cancellation_requested_at':data.get('cancellation_requested_at')}
    async def request_change(self,*,company_id:UUID,plan:str,interval:str)->dict:
        status=await self.status(company_id=company_id)
        if status['status'] in {'cancelled','expired'} and status['plan']!='trial':
            raise ApplicationError(message='This subscription is not active. Contact FinCruiz support to reactivate it.',error_code='SUBSCRIPTION_INACTIVE',status_code=409)
        market=resolve_market(status['billing_country_code']);catalog=PRICE_CATALOG.get(market.market_code,PRICE_CATALOG['GLOBAL'])
        if plan not in catalog:raise ApplicationError(message='That plan is not available in this market.',error_code='PLAN_NOT_AVAILABLE',status_code=422)
        now=datetime.now(timezone.utc)
        await self.session.execute(text('''UPDATE public.company_subscriptions SET requested_plan=:plan,requested_interval=:interval,change_requested_at=:now,updated_at=now() WHERE company_id=:company_id'''),{'plan':plan,'interval':interval,'now':now,'company_id':company_id});await self.session.commit()
        return {'requested_plan':plan,'requested_interval':interval,'change_requested_at':now,'message':'Plan request recorded. Billing activation remains manual until a payment provider is connected.'}
    async def request_cancellation(self,*,company_id:UUID)->dict:
        now=datetime.now(timezone.utc);await self.session.execute(text('UPDATE public.company_subscriptions SET cancellation_requested_at=:now,updated_at=now() WHERE company_id=:company_id'),{'now':now,'company_id':company_id});await self.session.commit();return await self.status(company_id=company_id)
    async def update_billing_market(self,*,company_id:UUID,country_code:str)->dict:
        current=await self._row(company_id);code=country_code.upper();resolve_market(code)
        if current and current.get('provider_subscription_id'):
            raise ApplicationError(message='Billing country cannot be changed while an external paid subscription is active. Contact support.',error_code='BILLING_COUNTRY_LOCKED',status_code=409)
        await self.session.execute(text('UPDATE public.company_subscriptions SET billing_country_code=:code,updated_at=now() WHERE company_id=:company_id'),{'code':code,'company_id':company_id});await self.session.commit();return await self.status(company_id=company_id)
