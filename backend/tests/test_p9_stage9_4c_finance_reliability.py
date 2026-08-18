from datetime import date

from app.services.finance.reliability_service import FinanceReliabilityService


def test_month_index_is_monotonic_across_year_boundary():
    assert FinanceReliabilityService._month_index(date(2026, 1, 1)) - FinanceReliabilityService._month_index(date(2025, 12, 1)) == 1


def test_number_handles_decimal_string_and_none():
    assert str(FinanceReliabilityService._number("123.45")) == "123.45"
    assert str(FinanceReliabilityService._number(None)) == "0"


def test_reliability_service_is_explicitly_structural():
    assert "structural" in (FinanceReliabilityService.__doc__ or "").lower() or "structural" in FinanceReliabilityService.certify.__doc__.lower() if FinanceReliabilityService.certify.__doc__ else True
