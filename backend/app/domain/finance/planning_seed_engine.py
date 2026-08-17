from __future__ import annotations
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
from dateutil.relativedelta import relativedelta

CENT = Decimal('0.01')


def month_range(start: date, end: date) -> list[date]:
    current = start.replace(day=1)
    last = end.replace(day=1)
    out: list[date] = []
    while current <= last:
        out.append(current)
        current = current + relativedelta(months=1)
    return out


def normalize_weights(values: list[Decimal]) -> list[Decimal]:
    if not values:
        return []
    positive = [abs(Decimal(v or 0)) for v in values]
    total = sum(positive, Decimal('0'))
    if total == 0:
        equal = Decimal('1') / Decimal(len(values))
        return [equal for _ in values]
    return [v / total for v in positive]


def allocate_total(total: Decimal, weights: list[Decimal]) -> list[Decimal]:
    """Allocate while preserving the requested total exactly to cents."""
    if not weights:
        return []
    norm = normalize_weights(weights)
    amounts = [(Decimal(total) * w).quantize(CENT, rounding=ROUND_HALF_UP) for w in norm]
    drift = Decimal(total).quantize(CENT) - sum(amounts, Decimal('0'))
    amounts[-1] += drift
    return amounts


def annualize_history(values: list[Decimal]) -> Decimal:
    if not values:
        return Decimal('0')
    values = [Decimal(v or 0) for v in values]
    if len(values) >= 12:
        return sum(values[-12:], Decimal('0'))
    return (sum(values, Decimal('0')) / Decimal(len(values)) * Decimal('12')).quantize(CENT)


def monthly_weights_from_history(values: list[Decimal], target_count: int) -> list[Decimal]:
    if target_count <= 0:
        return []
    if not values:
        return normalize_weights([Decimal('1')] * target_count)
    recent = [Decimal(v or 0) for v in values[-12:]]
    if len(recent) == 12 and target_count == 12:
        return normalize_weights(recent)
    # For other horizons use the recent monthly shape cyclically, or equal if sparse.
    base = normalize_weights(recent)
    if not base:
        return normalize_weights([Decimal('1')] * target_count)
    cycled = [base[i % len(base)] for i in range(target_count)]
    return normalize_weights(cycled)
