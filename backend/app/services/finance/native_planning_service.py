from __future__ import annotations
from decimal import Decimal
from uuid import UUID
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService

class NativePlanningService:
    def __init__(self, session: AsyncSession):
        self.session=session; self.reporting=ReportingService(GLTransactionRepository(session))

    async def create_version(self, company_id, request):
        version_id=(await self.session.execute(text('''
          INSERT INTO public.planning_versions(company_id,plan_type,version_name,financial_year_start,financial_year_end,assumptions)
          VALUES (:company_id,:plan_type,:version_name,:start,:end,CAST(:assumptions AS jsonb)) RETURNING id
        '''), {'company_id':company_id,'plan_type':request.plan_type,'version_name':request.version_name,
               'start':request.financial_year_start,'end':request.financial_year_end,
               'assumptions':'{"seed_growth_percent": '+str(float(request.seed_growth_percent))+'}'})).scalar_one()
        if request.seed_from_actuals:
            monthly=await self.reporting.monthly_actuals(company_id)
            growth=Decimal('1')+request.seed_growth_percent/100
            rows=[]
            for row in monthly:
                if request.financial_year_start <= row['month'] <= request.financial_year_end:
                    for key,group in [('revenue','Revenue'),('cost_of_sales','Cost of Sales'),('operating_expenses','Operating Expenses'),('depreciation','Depreciation'),('finance_costs','Finance Costs')]:
                        rows.append({'version_id':version_id,'company_id':company_id,'period':row['month'],'reporting_group':group,'amount':Decimal(row[key])*growth})
            if rows:
                await self.session.execute(text('''INSERT INTO public.native_plan_lines(version_id,company_id,period,reporting_group,amount,driver_type)
                  VALUES (:version_id,:company_id,:period,:reporting_group,:amount,'actual_growth')'''),rows)
        await self.session.commit(); return await self.get_version(company_id,version_id)

    async def list_versions(self, company_id):
        return [dict(r) for r in (await self.session.execute(text('''SELECT id,plan_type,version_name,financial_year_start,financial_year_end,status,assumptions
          FROM public.planning_versions WHERE company_id=:company_id ORDER BY updated_at DESC'''),{'company_id':company_id})).mappings().all()]

    async def get_version(self, company_id,version_id):
        version=(await self.session.execute(text('''SELECT id,plan_type,version_name,financial_year_start,financial_year_end,status,assumptions
          FROM public.planning_versions WHERE company_id=:company_id AND id=:id'''),{'company_id':company_id,'id':version_id})).mappings().one()
        lines=[dict(r) for r in (await self.session.execute(text('''SELECT id,period,branch_id,reporting_group,reporting_subgroup,source_account_code,amount,driver_type,driver_value,notes
          FROM public.native_plan_lines WHERE company_id=:company_id AND version_id=:id ORDER BY period,reporting_group'''),{'company_id':company_id,'id':version_id})).mappings().all()]
        return {**dict(version),'lines':lines}

    async def save_lines(self, company_id,version_id,lines):
        await self.session.execute(text('DELETE FROM public.native_plan_lines WHERE company_id=:company_id AND version_id=:id'),{'company_id':company_id,'id':version_id})
        rows=[{'version_id':version_id,'company_id':company_id,**line.model_dump()} for line in lines]
        if rows:
            await self.session.execute(text('''INSERT INTO public.native_plan_lines(version_id,company_id,period,branch_id,reporting_group,reporting_subgroup,source_account_code,amount,driver_type,driver_value,notes)
              VALUES (:version_id,:company_id,:period,:branch_id,:reporting_group,:reporting_subgroup,:source_account_code,:amount,:driver_type,:driver_value,:notes)'''),rows)
        await self.session.commit(); return await self.get_version(company_id,version_id)
