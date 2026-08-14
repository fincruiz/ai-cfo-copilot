from datetime import date
from decimal import Decimal

import pytest

from app.services.finance.assurance_service import FinancialAssuranceService


class ReportingStub:
    async def data_health(self, company_id):
        return {
            "transaction_count": 120,
            "is_trial_balance_balanced": True,
            "trial_balance_difference": Decimal("0"),
            "is_balance_sheet_balanced": True,
            "balance_sheet_difference": Decimal("0"),
            "is_mapping_complete": True,
            "mapped_account_count": 12,
            "account_count": 12,
            "invalid_transaction_count": 0,
            "duplicate_candidate_count": 0,
        }

    async def monthly_actuals(self, company_id):
        return [{"month": date(2026, month, 1)} for month in range(1, 7)]


def test_month_index_is_monotonic():
    assert FinancialAssuranceService._month_index(date(2026, 2, 1)) - FinancialAssuranceService._month_index(date(2026, 1, 1)) == 1
    assert FinancialAssuranceService._month_index(date(2027, 1, 1)) - FinancialAssuranceService._month_index(date(2026, 12, 1)) == 1


@pytest.mark.asyncio
async def test_assurance_returns_full_score_for_reconciled_continuous_data():
    result = await FinancialAssuranceService(ReportingStub()).assess("company")
    assert result["score"] == 100
    assert result["grade"] == "A"
    assert result["status"] == "ready"
    assert all(check["status"] == "pass" for check in result["checks"])
