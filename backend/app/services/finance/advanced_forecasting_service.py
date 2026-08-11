from __future__ import annotations
from copy import deepcopy
from dataclasses import asdict
from decimal import Decimal
from uuid import UUID
import json
import pandas as pd
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession
from app.domain.finance.advanced_forecasting import (
    ForecastConfig, ForecastDrivers, OpeningBalanceSheet, HistoricalData,
    TrendBudgetForecastBuilder, ThreeWayForecastEngine, ScenarioManager,
    ScenarioDefinition,
)
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService
from app.schemas.finance.advanced_forecasting import AdvancedForecastRequest, PowerOfOneRequest


def records(frame: pd.DataFrame) -> list[dict]:
    data = frame.reset_index().copy()
    for col in data.columns:
        if pd.api.types.is_datetime64_any_dtype(data[col]):
            data[col] = data[col].dt.strftime('%Y-%m-%d')
    return json.loads(data.to_json(orient='records', date_format='iso'))

class AdvancedForecastingService:
    def __init__(self, session: AsyncSession):
        self.session=session
        self.reporting=ReportingService(GLTransactionRepository(session))

    async def _history(self, company_id: UUID, branch_id: UUID | None):
        monthly=await self.reporting.monthly_actuals(company_id, branch_id=branch_id)
        if not monthly:
            raise ValueError('At least one month of mapped actuals is required.')
        return pd.DataFrame([{
            'Period': row['month'], 'Revenue': float(row['revenue']),
            'COGS': float(row['cost_of_sales']), 'Payroll': 0.0,
            'Other Opex': float(row['operating_expenses']) + float(row['depreciation']),
        } for row in monthly])

    async def _budget(self, company_id: UUID, version_id: UUID | None):
        if not version_id:
            return None
        rows=(await self.session.execute(text('''
            SELECT period, reporting_group, SUM(amount) amount
            FROM public.native_plan_lines
            WHERE company_id=:company_id AND version_id=:version_id
            GROUP BY period, reporting_group ORDER BY period
        '''), {'company_id':company_id,'version_id':version_id})).mappings().all()
        if not rows: return None
        by={}
        mapping={'Revenue':'Revenue','Cost of Sales':'COGS','Payroll':'Payroll','Operating Expenses':'Other Opex'}
        for r in rows:
            p=r['period']; by.setdefault(p, {'Period':p,'Revenue':0,'COGS':0,'Payroll':0,'Other Opex':0})
            target=mapping.get(r['reporting_group'])
            if target: by[p][target]+=float(r['amount'] or 0)
        return pd.DataFrame(by.values())

    def _config(self, request: AdvancedForecastRequest):
        d=request.drivers; ob=request.opening_balance_sheet
        drivers=ForecastDrivers(**{k:float(v) if isinstance(v,Decimal) else v for k,v in d.model_dump().items()})
        opening=OpeningBalanceSheet(**{k:(float(v) if isinstance(v,Decimal) else v) for k,v in ob.model_dump().items()})
        return ForecastConfig(
            forecast_start=request.forecast_start.isoformat(), forecast_months=request.forecast_months,
            trend_weight=float(request.trend_weight), budget_weight=float(request.budget_weight),
            recent_run_rate_weight=float(request.run_rate_weight), seasonality_enabled=request.seasonality_enabled,
            drivers=drivers, opening_balance_sheet=opening,
        )

    async def calculate(self, company_id: UUID, request: AdvancedForecastRequest, persist=True):
        history=await self._history(company_id, request.branch_id)
        budget=await self._budget(company_id, request.budget_version_id)
        config=self._config(request)
        build=TrendBudgetForecastBuilder(HistoricalData(history), budget, config).build()
        result=ThreeWayForecastEngine(build.forecast, config).run()
        manager=ScenarioManager(build.forecast, config); manager.add_standard_scenarios()
        scenario_results=manager.run_all(); comparison=manager.comparison(scenario_results)
        summary={
            'forecast_revenue':float(result.profit_and_loss['Revenue'].sum()),
            'forecast_ebitda':float(result.profit_and_loss['EBITDA'].sum()),
            'forecast_net_income':float(result.profit_and_loss['Net Income'].sum()),
            'closing_cash':float(result.balance_sheet['Cash'].iloc[-1]),
            'closing_debt':float(result.balance_sheet['Current Debt'].iloc[-1]+result.balance_sheet['Non-current Debt'].iloc[-1]),
            'minimum_cash':float(result.balance_sheet['Cash'].min()),
            'balanced':bool(result.checks['Balanced'].all()),
        }
        payload={
            'profit_and_loss':records(result.profit_and_loss), 'balance_sheet':records(result.balance_sheet),
            'cash_flow':records(result.cash_flow), 'ratios':records(result.ratios), 'checks':records(result.checks),
            'scenarios':records(comparison), 'diagnostics':records(build.diagnostics),
            'forecast_basis':records(build.forecast),
        }
        run_id=None
        if persist:
            row=(await self.session.execute(text('''
                INSERT INTO public.forecast_model_runs
                (company_id,branch_id,run_name,forecast_start,forecast_months,budget_version_id,configuration,summary,result_payload)
                VALUES (:company_id,:branch_id,:run_name,:forecast_start,:forecast_months,:budget_version_id,
                        CAST(:configuration AS jsonb),CAST(:summary AS jsonb),CAST(:payload AS jsonb)) RETURNING id
            '''), {'company_id':company_id,'branch_id':request.branch_id,'run_name':request.run_name,
                   'forecast_start':request.forecast_start,'forecast_months':request.forecast_months,
                   'budget_version_id':request.budget_version_id,
                   'configuration':json.dumps(request.model_dump(mode='json')),'summary':json.dumps(summary),'payload':json.dumps(payload)})).scalar_one()
            await self.session.commit(); run_id=row
        return {'run_id':run_id,'run_name':request.run_name,'summary':summary,**payload}

    async def power_of_one(self, company_id: UUID, request: PowerOfOneRequest):
        base_request=AdvancedForecastRequest(**request.model_dump(exclude={
            'price_change_percent','volume_change_percent','gross_margin_points','payroll_change_percent',
            'other_opex_change_percent','dso_change_days','dpo_change_days','inventory_change_days',
            'capex_change_percent','interest_rate_points'}))
        base=await self.calculate(company_id,base_request,persist=False)
        adjusted=deepcopy(base_request)
        combined=(Decimal('1')+request.price_change_percent/100)*(Decimal('1')+request.volume_change_percent/100)
        # Use scenario multipliers by applying forecast result through config/driver changes and rerun.
        adjusted.drivers.gross_margin += request.gross_margin_points/100
        adjusted.drivers.payroll_pct_revenue *= Decimal('1')+request.payroll_change_percent/100
        adjusted.drivers.other_opex_pct_revenue *= Decimal('1')+request.other_opex_change_percent/100
        adjusted.drivers.dso_days += request.dso_change_days
        adjusted.drivers.dpo_days += request.dpo_change_days
        adjusted.drivers.inventory_days += request.inventory_change_days
        adjusted.drivers.capex_pct_revenue *= Decimal('1')+request.capex_change_percent/100
        adjusted.drivers.annual_interest_rate += request.interest_rate_points/100
        history=await self._history(company_id, request.branch_id); history['Revenue']*=float(combined)
        budget=await self._budget(company_id, request.budget_version_id)
        config=self._config(adjusted)
        build=TrendBudgetForecastBuilder(HistoricalData(history),budget,config).build()
        result=ThreeWayForecastEngine(build.forecast,config).run()
        adj={'forecast_revenue':float(result.profit_and_loss['Revenue'].sum()),'forecast_ebitda':float(result.profit_and_loss['EBITDA'].sum()),
             'forecast_net_income':float(result.profit_and_loss['Net Income'].sum()),'closing_cash':float(result.balance_sheet['Cash'].iloc[-1]),
             'closing_debt':float(result.balance_sheet['Current Debt'].iloc[-1]+result.balance_sheet['Non-current Debt'].iloc[-1])}
        impact={k:adj[k]-base['summary'][k] for k in adj}
        return {'base':base['summary'],'adjusted':adj,'impact':impact}
