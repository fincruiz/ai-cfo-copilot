from dataclasses import dataclass
from decimal import Decimal
from typing import Iterable

from app.domain.finance.reporting.models import AccountBalance


@dataclass(frozen=True)
class TrialBalance:
    accounts: tuple[AccountBalance, ...]
    total_debit: Decimal
    total_credit: Decimal
    difference: Decimal


def build_trial_balance(accounts: Iterable[AccountBalance]) -> TrialBalance:
    rows = tuple(accounts)
    debit = sum((r.debit for r in rows), Decimal("0"))
    credit = sum((r.credit for r in rows), Decimal("0"))
    return TrialBalance(rows, debit, credit, debit-credit)
