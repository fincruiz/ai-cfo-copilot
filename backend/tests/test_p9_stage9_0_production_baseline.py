from decimal import Decimal

from app.database.models.core.company_member import CompanyMember
from app.database.models.finance.gl_transaction import GLTransaction
from app.repositories.finance.gl_transaction_repository import (
    GENERATED_COLUMNS,
    clean_gl_transaction_row,
)
from scripts.production_preflight import Check, render


def test_gl_cleaner_never_writes_database_generated_net_amount():
    source = {
        "company_id": "not-used-in-unit-test",
        "debit": Decimal("100.00"),
        "credit": Decimal("25.00"),
        "net_amount": Decimal("75.00"),
        "functional_currency_amount": Decimal("75.00"),
    }
    cleaned = clean_gl_transaction_row(source)
    assert "net_amount" not in cleaned
    assert cleaned["functional_currency_amount"] == Decimal("75.00")
    assert GENERATED_COLUMNS == {"net_amount"}


def test_net_amount_remains_declared_as_computed_column():
    column = GLTransaction.__table__.c.net_amount
    assert column.computed is not None
    sqltext = str(column.computed.sqltext).replace(" ", "").lower()
    assert "debit-credit" in sqltext


def test_company_member_model_matches_production_uniqueness():
    constraints = list(CompanyMember.__table__.constraints)
    unique_sets = {
        tuple(column.name for column in constraint.columns)
        for constraint in constraints
        if constraint.__class__.__name__ == "UniqueConstraint"
    }
    assert ("company_id", "user_id") in unique_sets


def test_preflight_exit_code_blocks_only_failures(capsys):
    assert render([Check("schema", "PASS", "ok")]) == 0
    assert render([Check("schema", "FAIL", "bad")]) == 2
    capsys.readouterr()
