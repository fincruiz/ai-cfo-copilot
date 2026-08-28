from decimal import Decimal
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

from app.domain.finance.gl_amounts import canonicalise_debit_credit
from app.domain.finance.gl_csv_validator import validate_gl_csv
from app.services.finance.ingestion_job_service import _normalise_row


def _company():
    return SimpleNamespace(id=uuid4(), currency_code="AUD")


def _mapping():
    return {
        "Date": "transaction_date",
        "Account Code": "source_account_code",
        "Debit": "debit",
        "Credit": "credit",
    }


def test_canonical_signed_debit_reversal_moves_to_credit_without_changing_net():
    debit, credit, changed = canonicalise_debit_credit(Decimal("-718000"), Decimal("0"))
    assert (debit, credit, changed) == (Decimal("0"), Decimal("718000"), True)
    assert debit - credit == Decimal("-718000")


def test_canonical_signed_credit_reversal_moves_to_debit_without_changing_net():
    debit, credit, changed = canonicalise_debit_credit(Decimal("0"), Decimal("-125.50"))
    assert (debit, credit, changed) == (Decimal("125.50"), Decimal("0"), True)
    assert debit - credit == Decimal("125.50")


def test_ambiguous_signed_row_is_rejected_instead_of_using_absolute_values():
    try:
        canonicalise_debit_credit(Decimal("-100"), Decimal("25"))
    except ValueError as exc:
        assert "opposite side is zero" in str(exc)
    else:
        raise AssertionError("Ambiguous signed debit/credit row must be rejected")


def test_background_import_normalises_signed_reversal_before_database_insert():
    row = _normalise_row(
        {"Date": "2026-08-17", "Account Code": "4000", "Debit": "-718000", "Credit": ""},
        _mapping(),
        row_number=12,
        company=_company(),
        upload_id=uuid4(),
        reporting_period_id=None,
    )
    assert row["debit"] == Decimal("0")
    assert row["credit"] == Decimal("718000")
    assert row["source_metadata"]["signed_reversal_normalised"] is True


def test_csv_validator_accepts_unambiguous_signed_reversal_but_rejects_ambiguous_pair():
    good = validate_gl_csv(
        b"Date,Account Code,Debit,Credit\n2026-08-17,4000,-718000,\n"
    )
    assert good.valid_rows == 1
    assert good.invalid_rows == 0

    bad = validate_gl_csv(
        b"Date,Account Code,Debit,Credit\n2026-08-17,4000,-718000,25\n"
    )
    assert bad.invalid_rows == 1
    assert any("opposite side is zero" in issue.message for issue in bad.issues)
