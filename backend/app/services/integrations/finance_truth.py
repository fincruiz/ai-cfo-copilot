from __future__ import annotations

import re
from collections import defaultdict
from datetime import UTC, date, datetime
from decimal import Decimal, InvalidOperation
from typing import Any
from uuid import UUID, uuid4

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.repositories.core.branch_repository import BranchRepository
from app.repositories.finance.file_upload_repository import FileUploadRepository
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.integrations.base import IntegrationStore

BALANCE_TOLERANCE = Decimal("0.01")


def _decimal(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as exc:
        raise ValueError(f"Invalid financial amount: {value!r}") from exc


def _date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        raise ValueError("Transaction date is required for canonical integration GL rows.")
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        try:
            return date.fromisoformat(text[:10])
        except ValueError as exc:
            raise ValueError(f"Unsupported integration transaction date: {text}") from exc


def _branch_code(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", value.upper())[:12] or "BRANCH"


def _normalise_sides(debit_value: Any, credit_value: Any) -> tuple[Decimal, Decimal]:
    debit = _decimal(debit_value)
    credit = _decimal(credit_value)

    # Preserve economic direction if a source expresses a reversal as a negative side.
    if debit < 0 and credit == 0:
        credit = abs(debit)
        debit = Decimal("0")
    if credit < 0 and debit == 0:
        debit = abs(credit)
        credit = Decimal("0")

    if debit < 0 or credit < 0:
        raise ValueError("Canonical GL debit/credit values must not both contain signed negatives.")
    if debit > 0 and credit > 0:
        raise ValueError("A canonical GL line cannot contain both a debit and a credit amount.")
    return debit, credit


class CanonicalIntegrationGLService:
    """Promote provider source records into the governed FinCruiz GL dataset.

    Source synchronization and financial truth activation are deliberately separate.
    A provider can be connected and synchronized without becoming the active ledger.  Only a
    complete, balanced, journal-grade source snapshot can supersede the current active GL.
    """

    def __init__(self, session: AsyncSession):
        self.session = session
        self.store = IntegrationStore(session)
        self.uploads = FileUploadRepository(session)
        self.transactions = GLTransactionRepository(session)
        self.branches = BranchRepository(session)

    async def _branch_id(
        self,
        *,
        company_id: UUID,
        upload_id: UUID,
        branch_reference: str | None,
        mapping: dict[str, Any],
    ) -> UUID | None:
        if not branch_reference:
            return None
        clean = branch_reference.strip()
        if not clean:
            return None
        key = clean.lower()
        branch = mapping.get(key)
        if branch is None:
            base_code = _branch_code(clean)
            code = base_code
            suffix = 2
            while await self.branches.find_by_code_or_name(company_id, code):
                code = f"{base_code[:9]}{suffix}"
                suffix += 1
            branch = await self.branches.create(
                {
                    "company_id": company_id,
                    "branch_code": code,
                    "branch_name": clean,
                    "region": None,
                    "review_status": "pending",
                    "source_value": clean,
                    "discovered_from_upload_id": upload_id,
                    "is_active": True,
                }
            )
            mapping[key] = branch
            mapping[code.lower()] = branch
        return branch.id

    def _prepare_rows(
        self,
        *,
        company: Company,
        provider: str,
        source_records: list[dict[str, Any]],
        upload_id: UUID,
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        rows: list[dict[str, Any]] = []
        total_debit = Decimal("0")
        total_credit = Decimal("0")
        source_group_totals: dict[str, list[Decimal]] = defaultdict(
            lambda: [Decimal("0"), Decimal("0")]
        )
        dates: list[date] = []
        skipped_zero = 0

        for index, record in enumerate(source_records, start=1):
            payload = record.get("payload") or {}
            finance = payload.get("fincruiz") or {}
            if finance.get("record_kind") != "canonical_gl_line":
                raise ValueError(
                    f"Integration record {record.get('external_id')} is not marked as a canonical GL line."
                )

            transaction_date = _date(
                finance.get("transaction_date") or record.get("occurred_at")
            )
            account_code = str(finance.get("account_code") or "").strip()
            if not account_code:
                raise ValueError(
                    f"Integration GL line {record.get('external_id')} has no account code."
                )

            debit, credit = _normalise_sides(
                finance.get("debit"), finance.get("credit")
            )
            if debit == 0 and credit == 0:
                skipped_zero += 1
                continue

            total_debit += debit
            total_credit += credit
            dates.append(transaction_date)

            source_transaction_id = str(
                finance.get("source_transaction_id")
                or finance.get("journal_number")
                or record.get("external_id")
            )
            source_group_totals[source_transaction_id][0] += debit
            source_group_totals[source_transaction_id][1] += credit

            functional_currency = str(
                finance.get("functional_currency_code") or company.currency_code
            ).strip().upper()
            company_currency = str(company.currency_code).strip().upper()
            if functional_currency != company_currency:
                raise ValueError(
                    "Provider functional currency does not match the FinCruiz company currency: "
                    f"provider {functional_currency}, FinCruiz {company_currency}. "
                    "Update the company currency or reconnect the correct source organisation before activation."
                )

            rows.append(
                {
                    "company_id": company.id,
                    "branch_id": None,  # assigned after source validation
                    "reporting_period_id": None,
                    "file_upload_id": upload_id,
                    "transaction_date": transaction_date,
                    "posting_date": transaction_date,
                    "document_date": transaction_date,
                    "document_number": finance.get("document_number"),
                    "journal_number": finance.get("journal_number"),
                    "batch_number": finance.get("batch_number"),
                    "source_account_code": account_code,
                    "source_account_name": finance.get("account_name") or record.get("name"),
                    "description": finance.get("description") or record.get("name"),
                    "reference": finance.get("reference"),
                    "customer_code": finance.get("customer_code"),
                    "supplier_code": finance.get("supplier_code"),
                    "project_code": finance.get("project_code"),
                    "cost_centre_code": finance.get("cost_centre_code"),
                    "department_code": finance.get("department_code"),
                    "debit": debit,
                    "credit": credit,
                    # Provider canonical debit/credit values are already functional-currency
                    # ledger amounts. Foreign/source currency remains in source metadata.
                    "currency_code": company_currency,
                    "exchange_rate": Decimal("1"),
                    "external_reference": str(record.get("external_id") or "") or None,
                    "source_row_number": index,
                    "is_adjustment": bool(finance.get("is_adjustment", False)),
                    "is_elimination": bool(finance.get("is_elimination", False)),
                    "is_intercompany": bool(finance.get("is_intercompany", False)),
                    "validation_status": "valid",
                    "validation_messages": [],
                    "source_metadata": {
                        "integration_provider": provider,
                        "integration_record_id": str(record.get("id") or ""),
                        "source_external_id": str(record.get("external_id") or ""),
                        "source_transaction_id": source_transaction_id,
                        "source_line_id": finance.get("source_line_id"),
                        "source_type": finance.get("source_type"),
                        "source_synced_at": str(record.get("synced_at") or ""),
                        "branch_reference": finance.get("branch_reference"),
                        "tracking": finance.get("tracking") or [],
                        "transaction_currency_code": finance.get("transaction_currency_code")
                        or record.get("currency_code"),
                        "functional_currency_code": functional_currency,
                    },
                    "_branch_reference": finance.get("branch_reference"),
                }
            )

        if not rows:
            raise ValueError("The provider did not return any non-zero journal-grade GL lines.")

        difference = total_debit - total_credit
        if abs(difference) > BALANCE_TOLERANCE:
            raise ValueError(
                "Provider GL snapshot is not balanced: "
                f"debits {total_debit}, credits {total_credit}, difference {difference}. "
                "The existing active FinCruiz ledger has been left unchanged."
            )

        unbalanced_groups = []
        for source_id, (debits, credits) in source_group_totals.items():
            group_diff = debits - credits
            if abs(group_diff) > BALANCE_TOLERANCE:
                unbalanced_groups.append(
                    {"source_transaction_id": source_id, "difference": str(group_diff)}
                )
                if len(unbalanced_groups) >= 10:
                    break

        summary = {
            "source_record_count": len(source_records),
            "canonical_row_count": len(rows),
            "skipped_zero_rows": skipped_zero,
            "total_debit": str(total_debit),
            "total_credit": str(total_credit),
            "balance_difference": str(difference),
            "data_start": min(dates).isoformat(),
            "data_through": max(dates).isoformat(),
            "unbalanced_source_transactions_sample": unbalanced_groups,
            "source_transaction_balance_warning_count": len(unbalanced_groups),
        }
        return rows, summary

    async def purge_provider(self, *, company_id: UUID, provider: str) -> dict[str, int]:
        """Delete canonical GL datasets created from one integration provider.

        Disconnecting a provider promises removal of the synchronized FinCruiz copy.
        User-uploaded CSV datasets are not touched, even when their source_system label
        happens to contain the same provider name.
        """
        deleted_transactions = int(
            (
                await self.session.execute(
                    text(
                        """
                        DELETE FROM public.gl_transactions
                        WHERE company_id=:company_id
                          AND file_upload_id IN (
                              SELECT id FROM public.file_uploads
                              WHERE company_id=:company_id
                                AND storage_bucket='integration-sync'
                                AND source_system=:provider
                          )
                        """
                    ),
                    {"company_id": company_id, "provider": provider},
                )
            ).rowcount
            or 0
        )
        deleted_datasets = int(
            (
                await self.session.execute(
                    text(
                        """
                        DELETE FROM public.file_uploads
                        WHERE company_id=:company_id
                          AND storage_bucket='integration-sync'
                          AND source_system=:provider
                        """
                    ),
                    {"company_id": company_id, "provider": provider},
                )
            ).rowcount
            or 0
        )
        return {
            "deleted_gl_transactions": deleted_transactions,
            "deleted_gl_datasets": deleted_datasets,
        }

    async def activate(
        self,
        *,
        company: Company,
        provider: str,
        activated_by: UUID | None,
    ) -> dict[str, Any]:
        source_records = await self.store.records(company.id, provider, "gl_line")
        if not source_records:
            result = {
                "status": "source_only",
                "provider": provider,
                "message": (
                    "Source records synchronized, but no journal-grade GL snapshot is available. "
                    "The current active FinCruiz ledger was not changed."
                ),
                "active_ledger_changed": False,
                "canonical_rows": 0,
            }
            await self.store.merge_metadata(company.id, provider, {"finance_truth": result})
            return result

        upload_id = uuid4()
        try:
            rows, summary = self._prepare_rows(
                company=company,
                provider=provider,
                source_records=source_records,
                upload_id=upload_id,
            )
        except ValueError as exc:
            await self.session.rollback()
            result = {
                "status": "blocked",
                "provider": provider,
                "message": str(exc),
                "active_ledger_changed": False,
                "canonical_rows": 0,
            }
            await self.store.merge_metadata(company.id, provider, {"finance_truth": result})
            return result

        now = datetime.now(UTC)
        filename = f"{provider}-canonical-gl-{now.strftime('%Y%m%dT%H%M%SZ')}.json"
        upload = await self.uploads.create(
            {
                "id": upload_id,
                "company_id": company.id,
                "reporting_period_id": None,
                "file_name": filename,
                "original_file_name": filename,
                "storage_bucket": "integration-sync",
                "storage_path": f"{company.id}/integrations/{provider}/{filename}",
                "mime_type": "application/json",
                "file_size_bytes": None,
                "document_type": "general_ledger",
                "source_system": provider,
                "processing_status": "validated",
                "is_active": False,
                "row_count": summary["source_record_count"],
                "valid_row_count": summary["canonical_row_count"],
                "invalid_row_count": 0,
                "validation_summary": {
                    "is_valid": True,
                    "source": "canonical_integration_gl",
                    **summary,
                },
                "column_mapping": {},
                "processing_metadata": {
                    "integration_provider": provider,
                    "finance_truth_version": "1.0",
                    "dataset_status": "pending_activation",
                    **summary,
                },
                "uploaded_by": activated_by,
                "processed_at": now,
            }
        )

        branch_mapping = await self.branches.mapping_by_code_and_name(company.id)
        for row in rows:
            branch_reference = row.pop("_branch_reference", None)
            row["branch_id"] = await self._branch_id(
                company_id=company.id,
                upload_id=upload.id,
                branch_reference=branch_reference,
                mapping=branch_mapping,
            )

        try:
            inserted = await self.transactions.bulk_create(rows)
            await self.uploads.deactivate_active_datasets(
                company_id=company.id,
                document_type="general_ledger",
                reporting_period_id=None,
                exclude_upload_id=upload.id,
            )
            await self.uploads.update(
                upload,
                {
                    "is_active": True,
                    "superseded_at": None,
                    "processing_metadata": {
                        **upload.processing_metadata,
                        "inserted_transaction_count": inserted,
                        "dataset_status": "active",
                        "activated_at": now.isoformat(),
                    },
                },
            )
            result = {
                "status": "activated",
                "provider": provider,
                "message": (
                    f"{provider.title()} is now the active FinCruiz General Ledger source "
                    f"with {inserted:,} validated journal lines."
                ),
                "active_ledger_changed": True,
                "active_upload_id": str(upload.id),
                "canonical_rows": inserted,
                "data_start": summary["data_start"],
                "data_through": summary["data_through"],
                "total_debit": summary["total_debit"],
                "total_credit": summary["total_credit"],
                "balance_difference": summary["balance_difference"],
                "source_transaction_balance_warning_count": summary[
                    "source_transaction_balance_warning_count"
                ],
            }
            await self.store.merge_metadata(
                company.id,
                provider,
                {"finance_truth": result},
                commit=False,
            )
            await self.session.commit()
            return result
        except Exception:
            await self.session.rollback()
            raise
