import pandas as pd
import pytest
from datetime import date
from decimal import Decimal
from uuid import uuid4

from app.schemas.finance.advanced_forecasting import DecisionSimulatorRequest
from app.services.finance.advanced_forecasting_service import AdvancedForecastingService


def request(**overrides):
    payload = dict(
        run_name='Hiring decision', forecast_start=date(2027,1,1), forecast_months=6,
        trend_weight=Decimal('0.45'), budget_weight=Decimal('0.35'), run_rate_weight=Decimal('0.20'),
    )
    payload.update(overrides)
    return DecisionSimulatorRequest(**payload)


@pytest.mark.asyncio
async def test_decision_simulator_is_deterministic_and_balanced():
    service = AdvancedForecastingService(None)
    history = pd.DataFrame([
        {'Period': f'2026-{m:02d}-01', 'Revenue': 100000 + m*2000, 'COGS': 55000 + m*1000, 'Payroll': 18000, 'Other Opex': 12000}
        for m in range(1,13)
    ])
    async def fake_history(company_id, branch_id): return history.copy()
    async def fake_budget(company_id, version_id): return None
    service._history = fake_history
    service._budget = fake_budget
    result = await service.simulate_decision(uuid4(), request(headcount_change=3, monthly_cost_per_hire=Decimal('5000'), dso_change_days=Decimal('10')))
    assert result['scenario_summary']['balanced'] is True
    assert len(result['comparison_series']) == 6
    assert result['scenario_summary']['forecast_net_income'] < result['base_summary']['forecast_net_income']
    assert result['impact']['forecast_net_income'] < 0
    assert result['assumptions']['headcount_change'] == 3


@pytest.mark.asyncio
async def test_revenue_upside_flows_into_scenario_revenue():
    service = AdvancedForecastingService(None)
    history = pd.DataFrame([
        {'Period': f'2026-{m:02d}-01', 'Revenue': 100000, 'COGS': 60000, 'Payroll': 15000, 'Other Opex': 10000}
        for m in range(1,13)
    ])
    async def fake_history(company_id, branch_id): return history.copy()
    async def fake_budget(company_id, version_id): return None
    service._history = fake_history
    service._budget = fake_budget
    result = await service.simulate_decision(uuid4(), request(revenue_change_percent=Decimal('10')))
    assert result['scenario_summary']['forecast_revenue'] == pytest.approx(result['base_summary']['forecast_revenue'] * 1.10)
    assert result['impact']['forecast_revenue'] > 0
