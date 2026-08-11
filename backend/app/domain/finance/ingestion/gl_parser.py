import csv
import io
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any
from uuid import UUID

from app.domain.finance.gl_csv_validator import (
    build_column_mapping,
    detect_delimiter,
)


DATE_FORMATS = (
    "%Y-%m-%d",
    "%d/%m/%Y",
    "%m/%d/%Y",
    "%d-%m-%Y",
    "%Y/%m/%d",
)


def _decimal(value: Any) -> Decimal:
    text = (
        str(value or "")
        .strip()
        .replace(",", "")
        .replace("$", "")
        .replace("₹", "")
        .replace("£", "")
        .replace("€", "")
    )

    if not text:
        return Decimal("0")

    is_parenthesised = (
        text.startswith("(")
        and text.endswith(")")
    )

    text = (
        text.replace("(", "")
        .replace(")", "")
        .strip()
    )

    try:
        amount = Decimal(text)
    except InvalidOperation as exc:
        raise ValueError(
            f"Invalid numeric amount: {value}"
        ) from exc

    if is_parenthesised:
        return -amount

    return amount


def _date(
    value: Any,
    *,
    required: bool = False,
) -> date | None:
    text = str(value or "").strip()

    if not text:
        if required:
            raise ValueError(
                "Transaction date is required."
            )

        return None

    for date_format in DATE_FORMATS:
        try:
            return datetime.strptime(
                text,
                date_format,
            ).date()
        except ValueError:
            continue

    try:
        return datetime.fromisoformat(
            text
        ).date()
    except ValueError as exc:
        raise ValueError(
            f"Unsupported date: {text}"
        ) from exc


def parse_validated_gl_csv(
    file_bytes: bytes,
    *,
    default_currency: str,
    company_id: UUID,
    file_upload_id: UUID,
    reporting_period_id: UUID | None = None,
) -> list[dict]:
    try:
        text = file_bytes.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError(
            "The CSV file must use UTF-8 encoding."
        ) from exc

    delimiter = detect_delimiter(text)

    reader = csv.DictReader(
        io.StringIO(text),
        delimiter=delimiter,
    )

    if not reader.fieldnames:
        raise ValueError(
            "CSV header is missing."
        )

    headers = [
        header.strip()
        for header in reader.fieldnames
        if header
    ]

    mapping, _, _ = build_column_mapping(
        headers
    )

    rows: list[dict] = []

    for row_number, raw_row in enumerate(
        reader,
        start=2,
    ):
        if not any(
            str(value or "").strip()
            for value in raw_row.values()
        ):
            continue

        row = {
            mapping.get(key, key): (
                value.strip()
                if isinstance(value, str)
                else value
            )
            for key, value in raw_row.items()
            if key
        }

        debit = _decimal(
            row.get("debit")
        )

        credit = _decimal(
            row.get("credit")
        )

        transaction_date = _date(
            row.get("transaction_date"),
            required=True,
        )

        exchange_rate = _decimal(
            row.get("exchange_rate") or "1"
        )

        source_account_code = str(
            row.get("source_account_code")
            or ""
        ).strip()

        if not source_account_code:
            raise ValueError(
                f"Account code is missing on row "
                f"{row_number}."
            )

        currency_code = str(
            row.get("currency_code")
            or default_currency
        ).strip().upper()

        rows.append(
            {
                "company_id": company_id,
                "_branch_reference": str(row.get("branch") or "").strip() or None,
                "reporting_period_id":
                    reporting_period_id,
                "file_upload_id":
                    file_upload_id,
                "transaction_date":
                    transaction_date,
                "posting_date": _date(
                    row.get("posting_date")
                ),
                "document_date": _date(
                    row.get("document_date")
                ),
                "document_number":
                    row.get("document_number")
                    or None,
                "journal_number":
                    row.get("journal_code")
                    or row.get("journal_number")
                    or None,
                "batch_number":
                    row.get("batch_number")
                    or None,
                "source_account_code":
                    source_account_code,
                "source_account_name":
                    row.get(
                        "source_account_name"
                    )
                    or None,
                "description":
                    row.get("description")
                    or None,
                "reference":
                    row.get("reference")
                    or None,
                "customer_code":
                    row.get("customer")
                    or row.get("customer_code")
                    or None,
                "supplier_code":
                    row.get("supplier")
                    or row.get("supplier_code")
                    or None,
                "project_code":
                    row.get("project")
                    or row.get("project_code")
                    or None,
                "cost_centre_code":
                    row.get("cost_centre")
                    or row.get(
                        "cost_centre_code"
                    )
                    or None,
                "department_code":
                    row.get("department")
                    or row.get(
                        "department_code"
                    )
                    or None,
                "debit": debit,
                "credit": credit,
                "currency_code":
                    currency_code,
                "exchange_rate":
                    exchange_rate,
                "external_reference":
                    row.get(
                        "external_reference"
                    )
                    or None,
                "source_row_number":
                    row_number,
                "validation_status":
                    "valid",
                "validation_messages":
                    [],
                "source_metadata": {
                    "raw_columns": list(
                        raw_row.keys()
                    )
                },
            }
        )

    return rows