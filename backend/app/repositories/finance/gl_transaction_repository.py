from __future__ import annotations

from datetime import date
from uuid import UUID

from sqlalchemy import case, delete, func, select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.branch import Branch
from app.database.models.finance.account_mapping import FinanceAccountMapping
from app.database.models.finance.file_upload import FileUpload
from app.database.models.finance.gl_transaction import GLTransaction

GENERATED_COLUMNS = {"net_amount"}


def clean_gl_transaction_row(row: dict) -> dict:
    """Return a database-safe GL row.

    PostgreSQL owns GENERATED ALWAYS columns such as ``net_amount``.  Keeping
    the sanitisation in one helper protects both the legacy synchronous upload
    path and the Stage 8.6 background/chunked importer.
    """
    cleaned = row.copy()
    for column in GENERATED_COLUMNS:
        cleaned.pop(column, None)
    return cleaned


class GLTransactionRepository:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def bulk_create(self, rows: list[dict]) -> int:
        if not rows:
            return 0
        transactions = []
        for row in rows:
            transactions.append(GLTransaction(**clean_gl_transaction_row(row)))
        self.session.add_all(transactions)
        await self.session.flush()
        return len(transactions)

    async def delete_by_upload(
        self,
        *,
        company_id: UUID,
        upload_id: UUID,
    ) -> int:
        result = await self.session.execute(
            delete(GLTransaction).where(
                GLTransaction.company_id == company_id,
                GLTransaction.file_upload_id == upload_id,
            )
        )
        return int(result.rowcount or 0)

    def _active_base_conditions(self, company_id: UUID):
        return (
            GLTransaction.company_id == company_id,
            GLTransaction.validation_status == "valid",
            FileUpload.is_active.is_(True),
            FileUpload.processing_status == "validated",
            GLTransaction.is_elimination.is_(False),
        )

    async def account_balances(
        self,
        *,
        company_id: UUID,
        start_date: date | None = None,
        end_date: date | None = None,
        statement: str | None = None,
        branch_id: UUID | None = None,
    ):
        query = (
            select(
                GLTransaction.source_account_code,
                func.max(GLTransaction.source_account_name).label("account_name"),
                func.sum(GLTransaction.debit).label("debit"),
                func.sum(GLTransaction.credit).label("credit"),
                FinanceAccountMapping.reporting_group,
                FinanceAccountMapping.reporting_subgroup,
                FinanceAccountMapping.statement,
                FinanceAccountMapping.sign_convention,
                FinanceAccountMapping.display_order,
            )
            .join(FileUpload, FileUpload.id == GLTransaction.file_upload_id)
            .outerjoin(
                FinanceAccountMapping,
                (
                    FinanceAccountMapping.company_id == GLTransaction.company_id
                )
                & (
                    FinanceAccountMapping.source_account_code
                    == GLTransaction.source_account_code
                ),
            )
            .where(*self._active_base_conditions(company_id))
        )
        if start_date is not None:
            query = query.where(GLTransaction.transaction_date >= start_date)
        if end_date is not None:
            query = query.where(GLTransaction.transaction_date <= end_date)
        if statement is not None:
            query = query.where(FinanceAccountMapping.statement == statement)
        if branch_id is not None:
            query = query.where(GLTransaction.branch_id == branch_id)

        query = query.group_by(
            GLTransaction.source_account_code,
            FinanceAccountMapping.reporting_group,
            FinanceAccountMapping.reporting_subgroup,
            FinanceAccountMapping.statement,
            FinanceAccountMapping.sign_convention,
            FinanceAccountMapping.display_order,
        )
        return (await self.session.execute(query)).all()

    async def monthly_actuals(
        self,
        *,
        company_id: UUID,
        branch_id: UUID | None = None,
        start_date: date | None = None,
        end_date: date | None = None,
    ):
        month = func.date_trunc(
            "month",
            GLTransaction.transaction_date,
        ).label("month")
        signed_amount = func.sum(
            case(
                (
                    FinanceAccountMapping.sign_convention.in_(("credit", "negative", "invert", "reverse")),
                    GLTransaction.credit - GLTransaction.debit,
                ),
                (
                    FinanceAccountMapping.sign_convention == "debit",
                    GLTransaction.debit - GLTransaction.credit,
                ),
                (
                    FinanceAccountMapping.reporting_group.in_(("Revenue", "Sales", "Other Income")),
                    GLTransaction.credit - GLTransaction.debit,
                ),
                else_=GLTransaction.debit - GLTransaction.credit,
            )
        ).label("amount")

        query = (
            select(
                month,
                FinanceAccountMapping.reporting_group,
                FinanceAccountMapping.reporting_subgroup,
                signed_amount,
            )
            .join(FileUpload, FileUpload.id == GLTransaction.file_upload_id)
            .join(
                FinanceAccountMapping,
                (
                    FinanceAccountMapping.company_id == GLTransaction.company_id
                )
                & (
                    FinanceAccountMapping.source_account_code
                    == GLTransaction.source_account_code
                ),
            )
            .where(
                *self._active_base_conditions(company_id),
                FinanceAccountMapping.statement == "income_statement",
            )
        )
        if branch_id is not None:
            query = query.where(GLTransaction.branch_id == branch_id)
        if start_date is not None:
            query = query.where(GLTransaction.transaction_date >= start_date)
        if end_date is not None:
            query = query.where(GLTransaction.transaction_date <= end_date)
        query = query.group_by(
            month,
            FinanceAccountMapping.reporting_group,
            FinanceAccountMapping.reporting_subgroup,
        ).order_by(month)
        return (await self.session.execute(query)).all()


    async def latest_transaction_date(
        self,
        *,
        company_id: UUID,
        branch_id: UUID | None = None,
    ) -> date | None:
        query = (
            select(func.max(GLTransaction.transaction_date))
            .join(FileUpload, FileUpload.id == GLTransaction.file_upload_id)
            .where(*self._active_base_conditions(company_id))
        )
        if branch_id is not None:
            query = query.where(GLTransaction.branch_id == branch_id)
        return (await self.session.execute(query)).scalar_one_or_none()

    async def monthly_group_totals(
        self,
        *,
        company_id: UUID,
        reporting_group: str,
        branch_id: UUID | None = None,
    ):
        rows = await self.monthly_actuals(
            company_id=company_id,
            branch_id=branch_id,
        )
        return [
            type("MonthlyRow", (), {"month": row.month, "net": row.amount})
            for row in rows
            if row.reporting_group == reporting_group
        ]

    async def branch_ids_with_activity(
        self,
        company_id: UUID,
    ) -> list[tuple[UUID, str, str]]:
        query = (
            select(Branch.id, Branch.branch_code, Branch.branch_name)
            .join(GLTransaction, GLTransaction.branch_id == Branch.id)
            .join(FileUpload, FileUpload.id == GLTransaction.file_upload_id)
            .where(
                Branch.company_id == company_id,
                Branch.is_active.is_(True),
                *self._active_base_conditions(company_id),
            )
            .distinct()
            .order_by(Branch.branch_code)
        )
        return list((await self.session.execute(query)).all())
