from __future__ import annotations

import csv
import io
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from uuid import UUID

from sqlalchemy import delete, select, text
from sqlalchemy.dialects.postgresql import insert
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.branch import Branch
from app.database.models.finance.account_mapping import FinanceAccountMapping
from app.schemas.finance.imports import FinanceImportResponse, ImportIssue


def _normalise_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_")


def _decode_csv(content: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return content.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("The file must be a readable UTF-8 CSV.")


def _decimal(value: object) -> Decimal:
    raw = str(value or "").strip().replace(",", "").replace("$", "").replace("₹", "")
    if not raw:
        return Decimal("0")
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid numeric value: {value}") from exc


def _parse_date(value: object) -> date | None:
    raw = str(value or "").strip()
    if not raw:
        return None
    for pattern in ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(raw, pattern).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(raw).date()
    except ValueError:
        return None


def _bucket(due_date: date | None, existing: str | None) -> tuple[str, int | None]:
    if existing and str(existing).strip():
        label = str(existing).strip()
        days = (date.today() - due_date).days if due_date else None
        return label, days
    if due_date is None:
        return "Unknown", None
    days = (date.today() - due_date).days
    if days <= 0:
        return "Current", days
    if days <= 30:
        return "1-30", days
    if days <= 60:
        return "31-60", days
    if days <= 90:
        return "61-90", days
    return "90+", days


class FinanceImportService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def _branch_lookup(self, company_id: UUID) -> dict[str, UUID]:
        rows = (
            await self.session.execute(
                select(Branch).where(
                    Branch.company_id == company_id,
                    Branch.is_active.is_(True),
                )
            )
        ).scalars().all()

        result: dict[str, UUID] = {}
        for branch in rows:
            result[branch.branch_code.strip().lower()] = branch.id
            result[branch.branch_name.strip().lower()] = branch.id
            if getattr(branch, "source_value", None):
                result[str(branch.source_value).strip().lower()] = branch.id
        return result

    async def import_coa(
        self,
        *,
        company_id: UUID,
        file_name: str,
        content: bytes,
        source_system: str | None,
    ) -> FinanceImportResponse:
        reader = csv.DictReader(io.StringIO(_decode_csv(content)))
        if not reader.fieldnames:
            raise ValueError("The COA file does not contain a header row.")

        aliases = {
            "account_code": ["account_code", "account code", "code", "gl_code", "gl code"],
            "account_name": ["account_name", "account name", "description", "gl_name", "gl name"],
            "reporting_group": ["reporting_group", "reporting group", "group"],
            "reporting_subgroup": ["reporting_subgroup", "reporting subgroup", "subgroup"],
            "statement": ["statement", "statement_type", "statement type"],
            "sign_convention": ["sign_convention", "sign convention", "sign"],
            "display_order": ["display_order", "display order", "report_order", "report order"],
        }

        normalised = {_normalise_header(name): name for name in reader.fieldnames}

        def source_column(key: str) -> str | None:
            for alias in aliases[key]:
                match = normalised.get(_normalise_header(alias))
                if match:
                    return match
            return None

        columns = {key: source_column(key) for key in aliases}
        required = ["account_code", "reporting_group", "statement"]
        missing = [key for key in required if not columns[key]]
        if missing:
            raise ValueError(
                "Missing required COA columns: " + ", ".join(missing)
            )

        issues: list[ImportIssue] = []
        rows_by_code: dict[str, dict] = {}
        total = 0

        for row_number, row in enumerate(reader, start=2):
            total += 1
            try:
                code = str(row.get(columns["account_code"] or "", "")).strip()
                if not code:
                    raise ValueError("Account code is required.")

                statement_raw = str(row.get(columns["statement"] or "", "")).strip().lower()
                statement = (
                    "income_statement"
                    if statement_raw in {"income statement", "profit and loss", "p&l", "pnl", "income_statement"}
                    else "balance_sheet"
                    if statement_raw in {"balance sheet", "bs", "balance_sheet"}
                    else statement_raw
                )
                if statement not in {"income_statement", "balance_sheet"}:
                    raise ValueError("Statement must be Income Statement or Balance Sheet.")

                sign = str(row.get(columns["sign_convention"] or "", "") or "positive").strip().lower()
                if sign not in {"debit", "credit", "positive"}:
                    sign = "positive"

                display_raw = str(row.get(columns["display_order"] or "", "") or "").strip()
                display_order = int(float(display_raw)) if display_raw else None

                rows_by_code[code] = {
                    "company_id": company_id,
                    "source_account_code": code,
                    "source_account_name": str(row.get(columns["account_name"] or "", "") or "").strip() or None,
                    "statement": statement,
                    "reporting_group": str(row.get(columns["reporting_group"] or "", "")).strip(),
                    "reporting_subgroup": str(row.get(columns["reporting_subgroup"] or "", "") or "").strip() or None,
                    "sign_convention": sign,
                    "display_order": display_order,
                    "is_confirmed": True,
                }
            except Exception as exc:
                issues.append(
                    ImportIssue(
                        row_number=row_number,
                        message=str(exc),
                    )
                )

        final_rows = list(rows_by_code.values())
        if final_rows:
            statement = insert(FinanceAccountMapping).values(final_rows)
            excluded = statement.excluded
            statement = statement.on_conflict_do_update(
                constraint="uq_finance_mapping_company_account",
                set_={
                    "source_account_name": excluded.source_account_name,
                    "statement": excluded.statement,
                    "reporting_group": excluded.reporting_group,
                    "reporting_subgroup": excluded.reporting_subgroup,
                    "sign_convention": excluded.sign_convention,
                    "display_order": excluded.display_order,
                    "is_confirmed": True,
                    "updated_at": text("now()"),
                },
            )
            await self.session.execute(statement)

        await self.session.execute(
            text(
                """
                INSERT INTO public.finance_import_batches
                (company_id, import_type, original_file_name, source_system,
                 row_count, valid_row_count, invalid_row_count, status, validation_summary)
                VALUES
                (:company_id, 'coa', :file_name, :source_system,
                 :row_count, :valid_rows, :invalid_rows, 'completed',
                 CAST(:summary AS jsonb))
                """
            ),
            {
                "company_id": company_id,
                "file_name": file_name,
                "source_system": source_system,
                "row_count": total,
                "valid_rows": len(final_rows),
                "invalid_rows": len(issues),
                "summary": __import__("json").dumps({"issues": [item.model_dump() for item in issues]}),
            },
        )
        await self.session.commit()

        return FinanceImportResponse(
            import_type="coa",
            original_file_name=file_name,
            total_rows=total,
            valid_rows=len(final_rows),
            invalid_rows=len(issues),
            inserted_rows=len(final_rows),
            issues=issues,
            metadata={"duplicates_resolved": total - len(final_rows) - len(issues)},
        )

    async def import_ageing(
        self,
        *,
        company_id: UUID,
        ageing_type: str,
        file_name: str,
        content: bytes,
        source_system: str | None,
        replace_existing: bool,
    ) -> FinanceImportResponse:
        reader = csv.DictReader(io.StringIO(_decode_csv(content)))
        if not reader.fieldnames:
            raise ValueError("The ageing file does not contain a header row.")

        aliases = {
            "party_name": ["party_name", "party name", "customer", "customer_name", "supplier", "supplier_name", "vendor", "vendor_name"],
            "document_number": ["document_number", "document number", "invoice_number", "invoice number", "invoice_no", "bill_number", "bill number"],
            "document_date": ["document_date", "document date", "invoice_date", "invoice date", "bill_date", "bill date"],
            "due_date": ["due_date", "due date"],
            "outstanding_amount": ["outstanding_amount", "outstanding amount", "outstanding", "outstanding_balance", "balance", "amount"],
            "original_amount": ["original_amount", "original amount", "invoice_amount", "invoice amount", "bill_amount", "bill amount"],
            "paid_amount": ["paid_amount", "paid amount", "payment_amount", "payment amount"],
            "branch": ["branch", "branch_code", "branch code", "location", "business_unit", "business unit"],
            "age_bucket": ["age_bucket", "age bucket", "ageing_bucket", "aging_bucket"],
            "currency_code": ["currency_code", "currency code", "currency"],
        }
        normalised = {_normalise_header(name): name for name in reader.fieldnames}

        def source_column(key: str) -> str | None:
            for alias in aliases[key]:
                match = normalised.get(_normalise_header(alias))
                if match:
                    return match
            return None

        columns = {key: source_column(key) for key in aliases}
        missing = [key for key in ("party_name", "outstanding_amount") if not columns[key]]
        if missing:
            raise ValueError("Missing required ageing columns: " + ", ".join(missing))

        branch_lookup = await self._branch_lookup(company_id)
        issues: list[ImportIssue] = []
        rows: list[dict] = []
        total = 0

        for row_number, row in enumerate(reader, start=2):
            total += 1
            try:
                party = str(row.get(columns["party_name"] or "", "")).strip()
                if not party:
                    raise ValueError("Party name is required.")

                outstanding = _decimal(row.get(columns["outstanding_amount"] or "", ""))
                document_date = _parse_date(row.get(columns["document_date"] or "", ""))
                due_date = _parse_date(row.get(columns["due_date"] or "", ""))
                bucket, days_overdue = _bucket(
                    due_date,
                    str(row.get(columns["age_bucket"] or "", "") or "").strip() or None,
                )
                branch_value = str(row.get(columns["branch"] or "", "") or "").strip()
                branch_id = branch_lookup.get(branch_value.lower()) if branch_value else None

                rows.append(
                    {
                        "company_id": company_id,
                        "ageing_type": ageing_type,
                        "party_name": party,
                        "document_number": str(row.get(columns["document_number"] or "", "") or "").strip() or None,
                        "document_date": document_date,
                        "due_date": due_date,
                        "outstanding_amount": outstanding,
                        "original_amount": _decimal(row.get(columns["original_amount"] or "", "")) if columns["original_amount"] else None,
                        "paid_amount": _decimal(row.get(columns["paid_amount"] or "", "")) if columns["paid_amount"] else None,
                        "branch_id": branch_id,
                        "branch_source_value": branch_value or None,
                        "age_bucket": bucket,
                        "days_overdue": days_overdue,
                        "currency_code": str(row.get(columns["currency_code"] or "", "") or "").strip() or None,
                        "source_row_number": row_number,
                    }
                )
            except Exception as exc:
                issues.append(ImportIssue(row_number=row_number, message=str(exc)))

        if replace_existing:
            await self.session.execute(
                text(
                    """
                    DELETE FROM public.finance_ageing_documents
                    WHERE company_id = :company_id AND ageing_type = :ageing_type
                    """
                ),
                {"company_id": company_id, "ageing_type": ageing_type},
            )

        if rows:
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.finance_ageing_documents
                    (
                        company_id, ageing_type, party_name, document_number,
                        document_date, due_date, outstanding_amount, original_amount,
                        paid_amount, branch_id, branch_source_value, age_bucket,
                        days_overdue, currency_code, source_row_number
                    )
                    VALUES
                    (
                        :company_id, :ageing_type, :party_name, :document_number,
                        :document_date, :due_date, :outstanding_amount, :original_amount,
                        :paid_amount, :branch_id, :branch_source_value, :age_bucket,
                        :days_overdue, :currency_code, :source_row_number
                    )
                    """
                ),
                rows,
            )

        await self.session.execute(
            text(
                """
                INSERT INTO public.finance_import_batches
                (company_id, import_type, original_file_name, source_system,
                 row_count, valid_row_count, invalid_row_count, status, validation_summary)
                VALUES
                (:company_id, :import_type, :file_name, :source_system,
                 :row_count, :valid_rows, :invalid_rows, 'completed',
                 CAST(:summary AS jsonb))
                """
            ),
            {
                "company_id": company_id,
                "import_type": ageing_type.lower(),
                "file_name": file_name,
                "source_system": source_system,
                "row_count": total,
                "valid_rows": len(rows),
                "invalid_rows": len(issues),
                "summary": __import__("json").dumps({"issues": [item.model_dump() for item in issues]}),
            },
        )
        await self.session.commit()

        return FinanceImportResponse(
            import_type=ageing_type.lower(),
            original_file_name=file_name,
            total_rows=total,
            valid_rows=len(rows),
            invalid_rows=len(issues),
            inserted_rows=len(rows),
            issues=issues,
            metadata={"replace_existing": replace_existing},
        )
