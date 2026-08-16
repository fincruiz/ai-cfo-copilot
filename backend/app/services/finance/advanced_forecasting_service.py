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

    async def simulate_decision(self, company_id: UUID, request):
        """Run a management what-if through the deterministic three-way model.

        AI may infer which assumptions the user intends, but this method alone
        calculates the financial impact. It never calls an LLM.
        """
        history = await self._history(company_id, request.branch_id)
        budget = await self._budget(company_id, request.budget_version_id)
        base_request = AdvancedForecastRequest(**request.model_dump(exclude={
            'revenue_change_percent','price_change_percent','volume_change_percent',
            'gross_margin_points','headcount_change','monthly_cost_per_hire',
            'payroll_change_percent','other_opex_change_percent','dso_change_days',
            'dpo_change_days','inventory_change_days','capex_change_percent'
        }))
        base_config = self._config(base_request)
        base_build = TrendBudgetForecastBuilder(HistoricalData(history), budget, base_config).build()
        base_result = ThreeWayForecastEngine(base_build.forecast, base_config).run()

        scenario_forecast = base_build.forecast.copy()
        revenue_factor = (
            (1 + float(request.revenue_change_percent) / 100)
            * (1 + float(request.price_change_percent) / 100)
            * (1 + float(request.volume_change_percent) / 100)
        )
        scenario_forecast['Revenue'] = scenario_forecast['Revenue'] * revenue_factor

        if float(request.gross_margin_points):
            base_revenue = base_build.forecast['Revenue'].replace(0, pd.NA)
            base_margin = ((base_build.forecast['Revenue'] - base_build.forecast['COGS']) / base_revenue).fillna(0.0)
            target_margin = (base_margin + float(request.gross_margin_points) / 100).clip(lower=0.0, upper=1.0)
            scenario_forecast['COGS'] = scenario_forecast['Revenue'] * (1 - target_margin)
        else:
            # Preserve the baseline COGS ratio when revenue changes.
            scenario_forecast['COGS'] = base_build.forecast['COGS'] * revenue_factor

        scenario_forecast['Payroll'] = (
            base_build.forecast['Payroll'] * (1 + float(request.payroll_change_percent) / 100)
            + float(request.headcount_change) * float(request.monthly_cost_per_hire)
        ).clip(lower=0.0)
        scenario_forecast['Other Opex'] = (
            base_build.forecast['Other Opex'] * (1 + float(request.other_opex_change_percent) / 100)
        ).clip(lower=0.0)

        scenario_request = deepcopy(base_request)
        scenario_request.drivers.dso_days += request.dso_change_days
        scenario_request.drivers.dpo_days += request.dpo_change_days
        scenario_request.drivers.inventory_days += request.inventory_change_days
        scenario_request.drivers.capex_pct_revenue *= Decimal('1') + request.capex_change_percent / 100
        scenario_config = self._config(scenario_request)
        scenario_result = ThreeWayForecastEngine(scenario_forecast, scenario_config).run()

        def summarize(result):
            pl, bs = result.profit_and_loss, result.balance_sheet
            minimum_cash = float(bs['Cash'].min())
            breach = bs.index[bs['Cash'] < scenario_config.drivers.minimum_cash]
            return {
                'forecast_revenue': float(pl['Revenue'].sum()),
                'forecast_ebitda': float(pl['EBITDA'].sum()),
                'forecast_net_income': float(pl['Net Income'].sum()),
                'closing_cash': float(bs['Cash'].iloc[-1]),
                'minimum_cash': minimum_cash,
                'closing_debt': float(bs['Current Debt'].iloc[-1] + bs['Non-current Debt'].iloc[-1]),
                'first_cash_pressure_month': breach[0].strftime('%Y-%m-%d') if len(breach) else None,
                'balanced': bool(result.checks['Balanced'].all()),
            }

        base_summary = summarize(base_result)
        scenario_summary = summarize(scenario_result)
        numeric_keys = ['forecast_revenue','forecast_ebitda','forecast_net_income','closing_cash','minimum_cash','closing_debt']
        impact = {key: scenario_summary[key] - base_summary[key] for key in numeric_keys}
        min_cash_target = float(scenario_config.drivers.minimum_cash)
        if not scenario_summary['balanced']:
            level, title = 'red', 'Model integrity requires attention'
        elif scenario_summary['minimum_cash'] < min_cash_target:
            level, title = 'red', 'Cash buffer falls below the management minimum'
        elif impact['closing_cash'] < 0 or impact['forecast_net_income'] < 0:
            level, title = 'amber', 'Decision is financially possible but weakens the outlook'
        else:
            level, title = 'green', 'Decision remains within the current financial guardrails'
        assessment = {
            'level': level,
            'title': title,
            'message': (
                f"Scenario closing cash changes by {impact['closing_cash']:.2f} and forecast net income by "
                f"{impact['forecast_net_income']:.2f}."
            ),
            'minimum_cash_target': min_cash_target,
        }

        comparison = []
        for period in base_result.profit_and_loss.index:
            comparison.append({
                'period': period.strftime('%Y-%m-%d'),
                'base_revenue': float(base_result.profit_and_loss.loc[period, 'Revenue']),
                'scenario_revenue': float(scenario_result.profit_and_loss.loc[period, 'Revenue']),
                'base_net_income': float(base_result.profit_and_loss.loc[period, 'Net Income']),
                'scenario_net_income': float(scenario_result.profit_and_loss.loc[period, 'Net Income']),
                'base_cash': float(base_result.balance_sheet.loc[period, 'Cash']),
                'scenario_cash': float(scenario_result.balance_sheet.loc[period, 'Cash']),
            })
        return {
            'scenario_name': request.run_name,
            'assumptions': {
                'revenue_change_percent': float(request.revenue_change_percent),
                'price_change_percent': float(request.price_change_percent),
                'volume_change_percent': float(request.volume_change_percent),
                'gross_margin_points': float(request.gross_margin_points),
                'headcount_change': request.headcount_change,
                'monthly_cost_per_hire': float(request.monthly_cost_per_hire),
                'payroll_change_percent': float(request.payroll_change_percent),
                'other_opex_change_percent': float(request.other_opex_change_percent),
                'dso_change_days': float(request.dso_change_days),
                'dpo_change_days': float(request.dpo_change_days),
                'inventory_change_days': float(request.inventory_change_days),
                'capex_change_percent': float(request.capex_change_percent),
            },
            'base_summary': base_summary,
            'scenario_summary': scenario_summary,
            'impact': impact,
            'assessment': assessment,
            'comparison_series': comparison,
            'base_checks': records(base_result.checks),
            'scenario_checks': records(scenario_result.checks),
        }
