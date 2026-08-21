from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

import pytest

from app.core.config import settings
from app.services.integrations.finance_truth import CanonicalIntegrationGLService
from app.services.integrations.health import integration_health
from app.services.integrations.tally import normalize_tally_record
from app.services.integrations.xero import XeroConnector, xero_scopes
from app.services.integrations.zoho import ZohoConnector


def test_xero_journals_become_balanced_canonical_debit_credit_lines():
    journals = [
        {
            "JournalID": "j-1",
            "JournalNumber": 42,
            "JournalDate": "2026-08-01",
            "Reference": "INV-1",
            "SourceID": "invoice-1",
            "SourceType": "ACCREC",
            "JournalLines": [
                {
                    "JournalLineID": "l-1",
                    "AccountID": "a-110",
                    "AccountCode": "110",
                    "AccountName": "Accounts Receivable",
                    "NetAmount": 110.0,
                    "TrackingCategories": [],
                },
                {
                    "JournalLineID": "l-2",
                    "AccountID": "a-200",
                    "AccountCode": "200",
                    "AccountName": "Sales",
                    "NetAmount": -100.0,
                    "TrackingCategories": [],
                },
                {
                    "JournalLineID": "l-3",
                    "AccountID": "a-820",
                    "AccountCode": "820",
                    "AccountName": "GST Payable",
                    "NetAmount": -10.0,
                    "TrackingCategories": [],
                },
            ],
        }
    ]
    rows = XeroConnector._journal_records(journals)
    assert len(rows) == 3
    finance = [row["payload"]["fincruiz"] for row in rows]
    assert sum(float(row["debit"]) for row in finance) == 110.0
    assert sum(float(row["credit"]) for row in finance) == 110.0
    assert all(row["record_kind"] == "canonical_gl_line" for row in finance)
    assert finance[0]["source_transaction_id"] == "invoice-1"


def test_xero_journal_scope_is_explicitly_gated(monkeypatch):
    monkeypatch.setattr(settings, "xero_journals_enabled", False)
    assert "accounting.journals.read" not in xero_scopes().split()
    monkeypatch.setattr(settings, "xero_journals_enabled", True)
    assert "accounting.journals.read" in xero_scopes().split()


def test_zoho_account_register_transactions_become_gl_lines():
    accounts = [
        {
            "account_id": "a1",
            "account_code": "1000",
            "account_name": "Cash",
        },
        {
            "account_id": "a2",
            "account_code": "4000",
            "account_name": "Revenue",
        },
    ]
    transactions = {
        "a1": [
            {
                "categorized_transaction_id": "c1",
                "transaction_id": "t1",
                "transaction_date": "2026-08-02",
                "entry_number": "INV-1",
                "transaction_type": "invoice",
                "debit_amount": "100",
                "credit_amount": "0",
            }
        ],
        "a2": [
            {
                "categorized_transaction_id": "c2",
                "transaction_id": "t1",
                "transaction_date": "2026-08-02",
                "entry_number": "INV-1",
                "transaction_type": "invoice",
                "debit_amount": "0",
                "credit_amount": "100",
            }
        ],
    }
    rows = ZohoConnector._gl_records(accounts, transactions)
    assert len(rows) == 2
    finance = [row["payload"]["fincruiz"] for row in rows]
    assert {row["account_code"] for row in finance} == {"1000", "4000"}
    assert {row["source_transaction_id"] for row in finance} == {"t1"}
    assert sum(float(row["debit"]) for row in finance) == 100.0
    assert sum(float(row["credit"]) for row in finance) == 100.0


def test_canonical_finance_truth_accepts_balanced_snapshot_and_rejects_imbalance():
    company = SimpleNamespace(id=uuid4(), currency_code="AUD")
    upload_id = uuid4()
    service = CanonicalIntegrationGLService(None)  # pure normalisation path only

    def source(external_id: str, account: str, debit: str, credit: str):
        return {
            "id": uuid4(),
            "external_id": external_id,
            "occurred_at": datetime(2026, 8, 1, tzinfo=timezone.utc),
            "synced_at": datetime(2026, 8, 2, tzinfo=timezone.utc),
            "payload": {
                "fincruiz": {
                    "record_kind": "canonical_gl_line",
                    "transaction_date": "2026-08-01",
                    "account_code": account,
                    "account_name": account,
                    "debit": debit,
                    "credit": credit,
                    "source_transaction_id": "tx-1",
                    "source_line_id": external_id,
                    "source_type": "test",
                }
            },
        }

    rows, summary = service._prepare_rows(
        company=company,
        provider="test",
        source_records=[
            source("1", "100", "100", "0"),
            source("2", "400", "0", "100"),
        ],
        upload_id=upload_id,
    )
    assert len(rows) == 2
    assert summary["balance_difference"] == "0"
    assert summary["data_through"] == "2026-08-01"

    with pytest.raises(ValueError, match="not balanced"):
        service._prepare_rows(
            company=company,
            provider="test",
            source_records=[source("1", "100", "100", "0")],
            upload_id=upload_id,
        )


def test_tally_requires_complete_snapshot_before_finance_truth_activation_contract():
    row = normalize_tally_record(
        {
            "entity_type": "ledger_line",
            "external_id": "line-1",
            "name": "Sales",
            "amount": 250,
            "currency_code": "INR",
            "occurred_at": datetime(2026, 8, 3, tzinfo=timezone.utc),
            "payload": {
                "voucher_id": "v-1",
                "voucher_number": "S-1",
                "voucher_type": "Sales",
                "ledger_code": "4000",
                "ledger_name": "Sales",
                "debit_or_credit": "credit",
                "amount": 250,
                "branch": "North",
            },
        }
    )
    assert row["entity_type"] == "gl_line"
    finance = row["payload"]["fincruiz"]
    assert finance["record_kind"] == "canonical_gl_line"
    assert finance["credit"] == "250"
    assert finance["debit"] == "0"
    assert finance["branch_reference"] == "North"

    schema = Path("app/schemas/integrations.py").read_text(encoding="utf-8")
    router = Path("app/api/v1/integrations/router.py").read_text(encoding="utf-8")
    assert "snapshot_start: bool = False" in schema
    assert "snapshot_complete: bool = False" in schema
    assert "if payload.snapshot_complete:" in router


def test_integration_health_distinguishes_source_sync_from_finance_truth():
    now = datetime.now(timezone.utc)
    activated = integration_health(
        {
            "status": "connected",
            "configured": True,
            "last_sync_status": "success",
            "last_synced_at": now,
            "metadata": {
                "finance_truth": {
                    "status": "activated",
                    "canonical_rows": 123,
                    "data_through": "2026-08-20",
                }
            },
        }
    )
    assert activated["health_status"] == "healthy"
    assert "123 GL lines" in activated["health_message"]

    blocked = integration_health(
        {
            "status": "connected",
            "configured": True,
            "last_sync_status": "success",
            "last_synced_at": now,
            "metadata": {"finance_truth": {"status": "blocked", "message": "TB failed"}},
        }
    )
    assert blocked["health_status"] == "finance_blocked"
    assert "TB failed" in blocked["health_message"]

    source_only = integration_health(
        {
            "status": "connected",
            "configured": True,
            "last_sync_status": "success",
            "last_synced_at": now,
            "metadata": {"finance_truth": {"status": "source_only", "message": "No journal access"}},
        }
    )
    assert source_only["health_status"] == "source_only"


def test_sync_routes_activate_canonical_finance_truth_and_disconnect_purges_copy():
    router = Path("app/api/v1/integrations/router.py").read_text(encoding="utf-8")
    assert 'provider="xero"' in router
    assert 'provider="zoho"' in router
    assert "CanonicalIntegrationGLService(session).activate" in router
    assert "purge_provider" in router


def test_frontend_shows_whether_connection_is_driving_financial_reports():
    page = Path("../frontend/app/dashboard/integrations/page.tsx").read_text(encoding="utf-8")
    service = Path("../frontend/services/integration-service.ts").read_text(encoding="utf-8")
    assert "Financial reporting" in page
    assert "Driving reports" in page
    assert "Activation blocked" in page
    assert "Source only" in page
    assert "IntegrationSyncResult" in service



def test_canonical_finance_truth_blocks_functional_currency_mismatch():
    company = SimpleNamespace(id=uuid4(), currency_code="AUD")
    service = CanonicalIntegrationGLService(None)
    records = [
        {
            "id": uuid4(),
            "external_id": "1",
            "occurred_at": datetime(2026, 8, 1, tzinfo=timezone.utc),
            "synced_at": datetime(2026, 8, 2, tzinfo=timezone.utc),
            "payload": {
                "fincruiz": {
                    "record_kind": "canonical_gl_line",
                    "transaction_date": "2026-08-01",
                    "account_code": "100",
                    "debit": "100",
                    "credit": "0",
                    "source_transaction_id": "t1",
                    "functional_currency_code": "USD",
                }
            },
        },
        {
            "id": uuid4(),
            "external_id": "2",
            "occurred_at": datetime(2026, 8, 1, tzinfo=timezone.utc),
            "synced_at": datetime(2026, 8, 2, tzinfo=timezone.utc),
            "payload": {
                "fincruiz": {
                    "record_kind": "canonical_gl_line",
                    "transaction_date": "2026-08-01",
                    "account_code": "400",
                    "debit": "0",
                    "credit": "100",
                    "source_transaction_id": "t1",
                    "functional_currency_code": "USD",
                }
            },
        },
    ]
    with pytest.raises(ValueError, match="functional currency"):
        service._prepare_rows(
            company=company,
            provider="xero",
            source_records=records,
            upload_id=uuid4(),
        )

def test_stage9_7_migration_protects_integration_tenant_tables_and_indexes_source_path():
    migration = Path(
        "migrations/20260821_p9_stage9_7_canonical_integration_gl.sql"
    ).read_text(encoding="utf-8")
    assert "integration_connections ENABLE ROW LEVEL SECURITY" in migration
    assert "integration_oauth_states ENABLE ROW LEVEL SECURITY" in migration
    assert "integration_records ENABLE ROW LEVEL SECURITY" in migration
    assert "organizational_memory ENABLE ROW LEVEL SECURITY" in migration
    assert "ix_integration_records_company_provider_entity" in migration
    assert "ix_file_uploads_integration_source" in migration
