from decimal import Decimal
from uuid import uuid4

from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService


def test_demo_dataset_is_balanced():
    rows = WorkspaceLifecycleService._demo_transactions(
        company_id=uuid4(),
        upload_id=uuid4(),
        currency="AUD",
    )
    total_debit = sum((Decimal(str(row["debit"])) for row in rows), Decimal("0"))
    total_credit = sum((Decimal(str(row["credit"])) for row in rows), Decimal("0"))
    assert total_debit == total_credit
    assert len(rows) > 100


def test_demo_dataset_has_confirmed_mapping_for_every_account():
    company_id = uuid4()
    rows = WorkspaceLifecycleService._demo_transactions(
        company_id=company_id,
        upload_id=uuid4(),
        currency="AUD",
    )
    mappings = WorkspaceLifecycleService._demo_mappings(company_id)
    transaction_accounts = {row["source_account_code"] for row in rows}
    mapped_accounts = {row["source_account_code"] for row in mappings}
    assert transaction_accounts == mapped_accounts
    assert all(row["is_confirmed"] for row in mappings)

from collections import defaultdict
from app.domain.finance.reporting.models import AccountBalance
from app.domain.finance.reporting.pnl import build_profit_and_loss


def test_demo_dataset_matches_golden_profit_and_loss():
    company_id = uuid4()
    rows = WorkspaceLifecycleService._demo_transactions(
        company_id=company_id,
        upload_id=uuid4(),
        currency="AUD",
    )
    mappings = {
        row["source_account_code"]: row
        for row in WorkspaceLifecycleService._demo_mappings(company_id)
    }
    totals = defaultdict(lambda: {"debit": Decimal("0"), "credit": Decimal("0")})
    for row in rows:
        totals[row["source_account_code"]]["debit"] += Decimal(row["debit"])
        totals[row["source_account_code"]]["credit"] += Decimal(row["credit"])

    balances = []
    for code, totals_for_code in totals.items():
        mapping = mappings[code]
        if mapping["statement"] != "income_statement":
            continue
        debit = totals_for_code["debit"]
        credit = totals_for_code["credit"]
        signed = credit - debit if mapping["sign_convention"] == "credit" else debit - credit
        balances.append(
            AccountBalance(
                account_code=code,
                account_name=mapping["source_account_name"],
                reporting_group=mapping["reporting_group"],
                reporting_subgroup=mapping["reporting_subgroup"],
                debit=debit,
                credit=credit,
                signed_amount=signed,
            )
        )

    report = build_profit_and_loss(balances)
    assert report.revenue == Decimal("2655000")
    assert report.cost_of_sales == Decimal("1115100")
    assert report.gross_profit == Decimal("1539900")
    assert report.operating_expenses == Decimal("948000")
    assert report.finance_costs == Decimal("28800")
    assert report.net_profit == Decimal("563100")

from app.domain.finance.reporting.balance_sheet import build_balance_sheet


def test_demo_dataset_balance_sheet_reconciles_with_current_earnings():
    company_id = uuid4()
    rows = WorkspaceLifecycleService._demo_transactions(
        company_id=company_id,
        upload_id=uuid4(),
        currency="AUD",
    )
    mappings = {
        row["source_account_code"]: row
        for row in WorkspaceLifecycleService._demo_mappings(company_id)
    }
    totals = defaultdict(lambda: {"debit": Decimal("0"), "credit": Decimal("0")})
    for row in rows:
        totals[row["source_account_code"]]["debit"] += Decimal(row["debit"])
        totals[row["source_account_code"]]["credit"] += Decimal(row["credit"])

    pnl_balances = []
    bs_balances = []
    for code, totals_for_code in totals.items():
        mapping = mappings[code]
        debit = totals_for_code["debit"]
        credit = totals_for_code["credit"]
        signed = credit - debit if mapping["sign_convention"] == "credit" else debit - credit
        balance = AccountBalance(
            account_code=code,
            account_name=mapping["source_account_name"],
            reporting_group=mapping["reporting_group"],
            reporting_subgroup=mapping["reporting_subgroup"],
            debit=debit,
            credit=credit,
            signed_amount=signed,
        )
        (pnl_balances if mapping["statement"] == "income_statement" else bs_balances).append(balance)

    pnl = build_profit_and_loss(pnl_balances)
    bs = build_balance_sheet(bs_balances, current_period_earnings=pnl.net_profit)
    assert abs(bs.balance_difference) <= Decimal("0.01")
