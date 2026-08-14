from __future__ import annotations
from uuid import UUID
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from app.services.finance.ai_cfo_service import AICFOService
from app.services.finance.assurance_service import FinancialAssuranceService
from app.services.finance.reporting_service import ReportingService
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository

class BrainService:
    def __init__(self, session: AsyncSession):
        self.session=session
        self.ai=AICFOService(session)
        self.assurance=FinancialAssuranceService(ReportingService(GLTransactionRepository(session)))
    async def overview(self, company_id: UUID):
        con=(await self.session.execute(text("SELECT provider,status,external_tenant_name,last_synced_at,last_sync_status FROM public.integration_connections WHERE company_id=:c ORDER BY provider"),{'c':company_id})).mappings().all()
        counts=(await self.session.execute(text("SELECT provider,entity_type,count(*) AS count FROM public.integration_records WHERE company_id=:c GROUP BY provider,entity_type ORDER BY provider,entity_type"),{'c':company_id})).mappings().all()
        memories=(await self.session.execute(text("SELECT id,title,content,memory_type,importance,created_at FROM public.organizational_memory WHERE company_id=:c AND is_active=true ORDER BY created_at DESC LIMIT 20"),{'c':company_id})).mappings().all()
        signals=await self.ai.proactive_signals(company_id)
        assurance=await self.assurance.assess(company_id)
        return {'connections':[dict(x) for x in con],'source_counts':[dict(x) for x in counts],'memories':[dict(x) for x in memories],'signals':signals.get('signals',[]),'financial_assurance':assurance}
    async def add_memory(self, company_id: UUID, user_id: UUID, payload):
        row=(await self.session.execute(text("""INSERT INTO public.organizational_memory(company_id,memory_type,title,content,importance,created_by) VALUES(:c,:t,:title,:content,:i,:u) RETURNING id,title,content,memory_type,importance,created_at"""),{'c':company_id,'t':payload.memory_type,'title':payload.title,'content':payload.content,'i':payload.importance,'u':user_id})).mappings().first(); await self.session.commit(); return dict(row)
