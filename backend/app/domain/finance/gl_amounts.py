from __future__ import annotations

from decimal import Decimal


def canonicalise_debit_credit(
    debit: Decimal,
    credit: Decimal,
) -> tuple[Decimal, Decimal, bool]:
    """Return database-safe debit/credit sides without changing economic direction.

    Some accounting exports represent reversals as a negative value on one side
    (for example, debit=-100 and credit=0). FinCruiz stores a canonical journal
    representation where both sides are non-negative, so the reversal becomes
    debit=0 and credit=100. The signed net amount is unchanged.

    This is intentionally *not* an absolute-value fallback: a signed value is only
    moved when the opposite side is zero. Ambiguous rows are rejected.
    """
    if debit < 0 and credit == 0:
        return Decimal("0"), -debit, True

    if credit < 0 and debit == 0:
        return -credit, Decimal("0"), True

    if debit < 0 or credit < 0:
        raise ValueError(
            "A negative debit or credit can only be treated as a reversal when the opposite side is zero."
        )

    if debit > 0 and credit > 0:
        raise ValueError("A row cannot contain both a debit and a credit amount.")

    return debit, credit, False
