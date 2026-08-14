from __future__ import annotations

from datetime import date
from decimal import Decimal
from pathlib import Path
import shutil
from uuid import UUID, uuid4

from sqlalchemy import func, select, text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.branch import Branch
from app.database.models.core.company import Company
from app.database.models.finance.account_mapping import FinanceAccountMapping
from app.database.models.finance.file_upload import FileUpload
from app.database.models.finance.gl_transaction import GLTransaction


class WorkspaceLifecycleService:
    """Privacy-safe lifecycle operations for a single company workspace."""

    DATA_TABLES_IN_DELETE_ORDER = (
        "generated_artifacts",
        "board_pack_runs",
        "board_pack_templates",
        "scenario_model_runs",
        "forecast_model_runs",
        "native_plan_lines",
        "planning_versions",
        "finance_plan_lines",
        "finance_ageing_documents",
        "finance_import_batches",
        "finance_account_mappings",
        "gl_transactions",
        "file_uploads",
    )

    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def _table_exists(self, table_name: str) -> bool:
        result = await self.session.execute(
            text("SELECT to_regclass(:table_name)"),
            {"table_name": f"public.{table_name}"},
        )
        return result.scalar_one_or_none() is not None

    async def status(self, *, company_id: UUID) -> dict:
        uploads = int(
            (
                await self.session.execute(
                    select(func.count())
                    .select_from(FileUpload)
                    .where(FileUpload.company_id == company_id)
                )
            ).scalar_one()
        )
        transactions = int(
            (
                await self.session.execute(
                    select(func.count())
                    .select_from(GLTransaction)
                    .where(GLTransaction.company_id == company_id)
                )
            ).scalar_one()
        )
        mappings = int(
            (
                await self.session.execute(
                    select(func.count())
                    .select_from(FinanceAccountMapping)
                    .where(FinanceAccountMapping.company_id == company_id)
                )
            ).scalar_one()
        )
        demo_uploads = int(
            (
                await self.session.execute(
                    select(func.count())
                    .select_from(FileUpload)
                    .where(
                        FileUpload.company_id == company_id,
                        FileUpload.processing_metadata["is_demo"].astext == "true",
                        FileUpload.is_active.is_(True),
                    )
                )
            ).scalar_one()
        )
        return {
            "has_financial_data": transactions > 0 or uploads > 0,
            "demo_data_active": demo_uploads > 0,
            "upload_count": uploads,
            "transaction_count": transactions,
            "mapping_count": mappings,
        }


    RESET_SCOPES = {
        "general_ledger": ("gl_transactions", "file_uploads"),
        "account_mappings": ("finance_account_mappings",),
        "coa": ("finance_account_mappings",),
        "ar_ageing": ("finance_ageing_documents",),
        "ap_ageing": ("finance_ageing_documents",),
        "planning": ("native_plan_lines", "planning_versions", "finance_plan_lines"),
        "forecasts": ("scenario_model_runs", "forecast_model_runs"),
        "board_packs": ("generated_artifacts", "board_pack_runs", "board_pack_templates"),
        "branches": ("branches",),
    }

    async def reset_scope(self, *, company_id: UUID, scope: str) -> dict[str, int]:
        if scope not in self.RESET_SCOPES:
            raise ApplicationError(message="Unknown reset scope.", error_code="INVALID_RESET_SCOPE", status_code=422)
        deleted: dict[str, int] = {}
        for table_name in self.RESET_SCOPES[scope]:
            if not await self._table_exists(table_name):
                continue
            if scope in {"ar_ageing", "ap_ageing"} and table_name == "finance_ageing_documents":
                ageing_type = "AR" if scope == "ar_ageing" else "AP"
                result = await self.session.execute(text(f"DELETE FROM public.{table_name} WHERE company_id=:company_id AND ageing_type=:ageing_type"), {"company_id": company_id, "ageing_type": ageing_type})
            elif scope == "general_ledger" and table_name == "file_uploads":
                result = await self.session.execute(text("DELETE FROM public.file_uploads WHERE company_id=:company_id AND document_type='general_ledger'"), {"company_id": company_id})
            else:
                result = await self.session.execute(text(f"DELETE FROM public.{table_name} WHERE company_id=:company_id"), {"company_id": company_id})
            deleted[table_name] = int(result.rowcount or 0)
        await self.session.commit()
        if scope == "board_packs":
            shutil.rmtree(Path("generated_artifacts") / str(company_id), ignore_errors=True)
        return deleted

    async def reset_financial_data(self, *, company_id: UUID) -> dict[str, int]:
        """Delete imported/generated finance data while preserving company/profile/settings."""
        deleted: dict[str, int] = {}

        # Branches discovered from uploaded data belong to the imported dataset.
        branch_result = await self.session.execute(
            text(
                "DELETE FROM public.branches "
                "WHERE company_id=:company_id AND discovered_from_upload_id IS NOT NULL"
            ),
            {"company_id": company_id},
        )
        deleted["discovered_branches"] = int(branch_result.rowcount or 0)

        for table_name in self.DATA_TABLES_IN_DELETE_ORDER:
            if not await self._table_exists(table_name):
                continue
            result = await self.session.execute(
                text(f"DELETE FROM public.{table_name} WHERE company_id=:company_id"),
                {"company_id": company_id},
            )
            deleted[table_name] = int(result.rowcount or 0)

        await self.session.commit()

        # Board-pack exports are written to the local generated_artifacts folder.
        # Remove the company-scoped folder as part of the privacy reset.
        shutil.rmtree(Path("generated_artifacts") / str(company_id), ignore_errors=True)
        return deleted

    async def seed_demo_data(
        self,
        *,
        company: Company,
        user_id: UUID,
        replace_existing: bool = False,
    ) -> dict:
        current = await self.status(company_id=company.id)
        if current["has_financial_data"]:
            if not replace_existing:
                raise ApplicationError(
                    message=(
                        "This workspace already contains financial data. "
                        "Reset it first, or explicitly replace it with demo data."
                    ),
                    error_code="WORKSPACE_NOT_EMPTY",
                    status_code=409,
                )
            await self.reset_financial_data(company_id=company.id)

        upload_id = uuid4()
        upload = FileUpload(
            id=upload_id,
            company_id=company.id,
            file_name=f"{upload_id}_fincruiz-demo-ledger.csv",
            original_file_name="fincruiz-demo-ledger.csv",
            storage_bucket="demo",
            storage_path=f"{company.id}/demo/{upload_id}.csv",
            mime_type="text/csv",
            file_size_bytes=0,
            document_type="general_ledger",
            source_system="FinCruiz Demo",
            processing_status="validated",
            is_active=True,
            row_count=0,
            valid_row_count=0,
            invalid_row_count=0,
            validation_summary={"demo": True, "message": "Synthetic FinCruiz demonstration dataset."},
            column_mapping={},
            processing_metadata={
                "is_demo": True,
                "dataset_status": "active",
                "storage_status": "synthetic_not_stored",
                "gl_transactions_inserted": True,
            },
            uploaded_by=user_id,
        )
        self.session.add(upload)
        await self.session.flush()

        mappings = self._demo_mappings(company.id)
        self.session.add_all([FinanceAccountMapping(**row) for row in mappings])

        rows = self._demo_transactions(
            company_id=company.id,
            upload_id=upload_id,
            currency=company.currency_code,
        )
        self.session.add_all([GLTransaction(**row) for row in rows])

        upload.row_count = len(rows)
        upload.valid_row_count = len(rows)
        upload.processing_metadata = {
            **upload.processing_metadata,
            "inserted_transaction_count": len(rows),
        }

        await self.session.commit()
        return {
            "upload_id": str(upload_id),
            "months": 12,
            "transactions_created": len(rows),
            "mappings_created": len(mappings),
        }

    @staticmethod
    def _demo_mappings(company_id: UUID) -> list[dict]:
        mapping_specs = [
            ("1000", "Bank", "balance_sheet", "Current Assets", "Cash and Cash Equivalents", "debit", 10),
            ("1100", "Accounts Receivable", "balance_sheet", "Current Assets", "Trade Receivables", "debit", 20),
            ("1200", "Inventory", "balance_sheet", "Current Assets", "Inventory", "debit", 30),
            ("1500", "Plant and Equipment", "balance_sheet", "Non Current Assets", "Property Plant and Equipment", "debit", 40),
            ("2000", "Accounts Payable", "balance_sheet", "Current Liabilities", "Trade Payables", "credit", 50),
            ("2500", "Business Loan", "balance_sheet", "Non Current Liabilities", "Borrowings", "credit", 60),
            ("3000", "Owner Equity", "balance_sheet", "Equity", "Retained Earnings", "credit", 70),
            ("4000", "Sales Revenue", "income_statement", "Revenue", "Sales", "credit", 80),
            ("5000", "Cost of Sales", "income_statement", "Cost of Sales", None, "debit", 90),
            ("6100", "Salaries and Wages", "income_statement", "Operating Expenses", "People", "debit", 100),
            ("6200", "Rent", "income_statement", "Operating Expenses", "Occupancy", "debit", 110),
            ("6300", "Marketing", "income_statement", "Operating Expenses", "Marketing", "debit", 120),
            ("6400", "Software", "income_statement", "Operating Expenses", "Technology", "debit", 130),
            ("7000", "Interest Expense", "income_statement", "Finance Costs", None, "debit", 140),
        ]
        return [
            {
                "company_id": company_id,
                "source_account_code": code,
                "source_account_name": name,
                "statement": statement,
                "reporting_group": group,
                "reporting_subgroup": subgroup,
                "sign_convention": sign,
                "display_order": order,
                "is_confirmed": True,
            }
            for code, name, statement, group, subgroup, sign, order in mapping_specs
        ]

    @staticmethod
    def _demo_transactions(*, company_id: UUID, upload_id: UUID, currency: str) -> list[dict]:
        today = date.today()
        current_month = date(today.year, today.month, 1)

        def shift_months(d: date, offset: int) -> date:
            month_index = d.year * 12 + (d.month - 1) + offset
            return date(month_index // 12, month_index % 12 + 1, 1)

        rows: list[dict] = []
        row_no = 1

        names = {
            "1000": "Bank",
            "1100": "Accounts Receivable",
            "1200": "Inventory",
            "1500": "Plant and Equipment",
            "2000": "Accounts Payable",
            "2500": "Business Loan",
            "3000": "Owner Equity",
            "4000": "Sales Revenue",
            "5000": "Cost of Sales",
            "6100": "Salaries and Wages",
            "6200": "Rent",
            "6300": "Marketing",
            "6400": "Software",
            "7000": "Interest Expense",
        }

        def add(period: date, account: str, debit: Decimal = Decimal("0"), credit: Decimal = Decimal("0"), description: str = "") -> None:
            nonlocal row_no
            rows.append(
                {
                    "company_id": company_id,
                    "file_upload_id": upload_id,
                    "transaction_date": period,
                    "source_account_code": account,
                    "source_account_name": names[account],
                    "description": description,
                    "debit": debit,
                    "credit": credit,
                    "currency_code": currency,
                    "exchange_rate": Decimal("1"),
                    "source_row_number": row_no,
                    "validation_status": "valid",
                    "validation_messages": [],
                    "source_metadata": {"is_demo": True},
                }
            )
            row_no += 1

        opening = shift_months(current_month, -12)
        add(opening, "1000", Decimal("300000"), description="Demo opening balance")
        add(opening, "1100", Decimal("150000"), description="Demo opening balance")
        add(opening, "1200", Decimal("100000"), description="Demo opening balance")
        add(opening, "1500", Decimal("250000"), description="Demo opening balance")
        add(opening, "2000", credit=Decimal("100000"), description="Demo opening balance")
        add(opening, "2500", credit=Decimal("200000"), description="Demo opening balance")
        add(opening, "3000", credit=Decimal("500000"), description="Demo opening balance")

        for i in range(12):
            period = shift_months(current_month, -11 + i)
            revenue = Decimal(180000 + i * 7500)
            cogs = (revenue * Decimal("0.42")).quantize(Decimal("1"))
            salaries = Decimal(42000 + (i // 4) * 2000)
            rent = Decimal("15000")
            marketing = Decimal(11000 + (i % 3) * 2500)
            software = Decimal("6500")
            interest = Decimal("2400")
            cash_sale = (revenue * Decimal("0.35")).quantize(Decimal("1"))
            credit_sale = revenue - cash_sale
            collected = (credit_sale * Decimal("0.88")).quantize(Decimal("1"))
            supplier_payment = (cogs * Decimal("0.82")).quantize(Decimal("1"))

            # Sales and collections.
            add(period, "1000", cash_sale, description="Cash sales")
            add(period, "1100", credit_sale, description="Credit sales")
            add(period, "4000", credit=revenue, description="Monthly sales revenue")
            add(period, "1000", collected, description="Receivables collected")
            add(period, "1100", credit=collected, description="Receivables collected")

            # Inventory purchase and COGS recognition keep stock broadly stable.
            add(period, "1200", cogs, description="Inventory purchases")
            add(period, "2000", credit=cogs, description="Inventory purchases")
            add(period, "5000", cogs, description="Cost of sales")
            add(period, "1200", credit=cogs, description="Cost of sales")
            add(period, "2000", supplier_payment, description="Supplier payment")
            add(period, "1000", credit=supplier_payment, description="Supplier payment")

            for account, amount, description in (
                ("6100", salaries, "Payroll"),
                ("6200", rent, "Premises rent"),
                ("6300", marketing, "Marketing spend"),
                ("6400", software, "Software subscriptions"),
                ("7000", interest, "Loan interest"),
            ):
                add(period, account, amount, description=description)
                add(period, "1000", credit=amount, description=description)

        return rows
