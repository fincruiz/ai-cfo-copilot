from __future__ import annotations

from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from typing import Any


def _decimal(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError):
        return Decimal("0")


def _datetime(value: Any) -> datetime | None:
    if not value:
        return None
    if isinstance(value, datetime):
        return value if value.tzinfo else value.replace(tzinfo=timezone.utc)
    text = str(value).strip().replace("Z", "+00:00")
    for candidate in (text, text[:10]):
        try:
            result = datetime.fromisoformat(candidate)
            return result if result.tzinfo else result.replace(tzinfo=timezone.utc)
        except ValueError:
            continue
    return None


def normalize_tally_record(record: dict[str, Any]) -> dict[str, Any]:
    """Normalize a Tally bridge record while preserving the original payload.

    The bridge remains backwards-compatible with generic entity records.  When the
    entity type is ``gl_line`` or ``ledger_line``, FinCruiz requires enough fields to
    create a deterministic debit/credit GL line before it can become finance truth.
    """
    entity_type = str(record.get("entity_type") or "").strip().lower()
    payload = dict(record.get("payload") or {})
    occurred_at = record.get("occurred_at") or _datetime(
        payload.get("transaction_date")
        or payload.get("voucher_date")
        or payload.get("date")
    )

    if entity_type not in {"gl_line", "ledger_line"}:
        return {**record, "payload": payload, "occurred_at": occurred_at}

    account_name = (
        payload.get("source_account_name")
        or payload.get("account_name")
        or payload.get("ledger_name")
        or record.get("name")
    )
    account_code = (
        payload.get("source_account_code")
        or payload.get("account_code")
        or payload.get("ledger_code")
        or payload.get("account_id")
    )
    if not account_code and account_name:
        account_code = f"TALLY:{str(account_name).strip()}"

    debit = _decimal(payload.get("debit"))
    credit = _decimal(payload.get("credit"))
    if debit == 0 and credit == 0:
        amount = _decimal(payload.get("amount") or record.get("amount"))
        direction = str(
            payload.get("debit_or_credit")
            or payload.get("dr_cr")
            or payload.get("direction")
            or ""
        ).strip().lower()
        if direction in {"debit", "dr", "d"}:
            debit = abs(amount)
        elif direction in {"credit", "cr", "c"}:
            credit = abs(amount)
        elif amount > 0:
            debit = amount
        elif amount < 0:
            credit = abs(amount)

    source_transaction_id = (
        payload.get("voucher_id")
        or payload.get("transaction_id")
        or payload.get("voucher_number")
        or record.get("external_id")
    )
    source_line_id = (
        payload.get("ledger_entry_id")
        or payload.get("line_id")
        or record.get("external_id")
    )

    finance = {
        "record_kind": "canonical_gl_line",
        "transaction_date": (
            occurred_at.date().isoformat() if isinstance(occurred_at, datetime) else None
        ),
        "account_code": str(account_code or "").strip(),
        "account_name": account_name,
        "debit": str(debit),
        "credit": str(credit),
        "description": payload.get("description") or payload.get("narration"),
        "reference": payload.get("reference") or payload.get("reference_number"),
        "document_number": payload.get("voucher_number"),
        "journal_number": payload.get("voucher_number"),
        "batch_number": payload.get("batch_number"),
        "customer_code": payload.get("customer_code") or payload.get("party_code"),
        "supplier_code": payload.get("supplier_code") or payload.get("party_code"),
        "project_code": payload.get("project_code") or payload.get("job_code"),
        "cost_centre_code": payload.get("cost_centre_code") or payload.get("cost_center_code"),
        "department_code": payload.get("department_code"),
        "branch_reference": (
            payload.get("branch")
            or payload.get("branch_name")
            or payload.get("location")
        ),
        "functional_currency_code": payload.get("functional_currency_code"),
        "transaction_currency_code": record.get("currency_code") or payload.get("currency_code"),
        "exchange_rate": str(payload.get("exchange_rate") or "1"),
        "source_transaction_id": str(source_transaction_id or ""),
        "source_line_id": str(source_line_id or ""),
        "source_type": payload.get("voucher_type") or "tally_voucher",
    }
    payload["fincruiz"] = finance
    return {
        **record,
        "entity_type": "gl_line",
        "occurred_at": occurred_at,
        "name": account_name or record.get("name"),
        "amount": abs(debit or credit),
        "payload": payload,
    }
