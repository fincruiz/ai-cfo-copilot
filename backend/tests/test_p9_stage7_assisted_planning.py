from datetime import date
from decimal import Decimal

from app.domain.finance.planning_seed_engine import allocate_total, annualize_history, month_range, monthly_weights_from_history


def test_high_level_allocation_preserves_management_target_to_cents():
    weights = [Decimal('100'), Decimal('200'), Decimal('300')]
    allocated = allocate_total(Decimal('2000000.00'), weights)
    assert sum(allocated) == Decimal('2000000.00')
    assert allocated[-1] > allocated[0]


def test_sparse_actuals_are_annualised_for_budget_seed():
    assert annualize_history([Decimal('100'), Decimal('200'), Decimal('300')]) == Decimal('2400.00')


def test_budget_period_builder_respects_requested_horizon():
    months = month_range(date(2027, 1, 1), date(2027, 12, 31))
    assert len(months) == 12
    assert months[0] == date(2027, 1, 1)
    assert months[-1] == date(2027, 12, 1)


def test_historical_month_shape_normalises_for_target_period():
    weights = monthly_weights_from_history([Decimal('10'), Decimal('20'), Decimal('30')], 6)
    assert len(weights) == 6
    assert abs(sum(weights) - Decimal('1')) < Decimal('0.000001')
