from __future__ import annotations

from uuid import UUID

from sqlalchemy import select
from sqlalchemy.dialects.postgresql import insert
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.finance.account_mapping import (
    FinanceAccountMapping,
)
from app.database.models.finance.gl_transaction import (
    GLTransaction,
)
from app.database.models.finance.file_upload import FileUpload


class AccountMappingRepository:
    def __init__(
        self,
        session: AsyncSession,
    ) -> None:
        self.session = session

    async def list_mappings(
        self,
        company_id: UUID,
    ) -> list[FinanceAccountMapping]:
        statement = (
            select(FinanceAccountMapping)
            .where(
                FinanceAccountMapping.company_id
                == company_id
            )
            .order_by(
                FinanceAccountMapping.display_order.asc(),
                FinanceAccountMapping.source_account_code.asc(),
            )
        )

        result = await self.session.execute(
            statement
        )

        return list(result.scalars().all())

    async def upsert_many(
        self,
        rows: list[dict],
    ) -> int:
        if not rows:
            return 0

        statement = insert(
            FinanceAccountMapping
        ).values(rows)

        excluded = statement.excluded

        statement = statement.on_conflict_do_update(
            constraint=(
                "uq_finance_mapping_company_account"
            ),
            set_={
                "source_account_name":
                    excluded.source_account_name,
                "statement":
                    excluded.statement,
                "reporting_group":
                    excluded.reporting_group,
                "reporting_subgroup":
                    excluded.reporting_subgroup,
                "sign_convention":
                    excluded.sign_convention,
                "display_order":
                    excluded.display_order,
                "is_confirmed":
                    excluded.is_confirmed,
                "updated_at":
                    excluded.updated_at,
            },
        )

        await self.session.execute(statement)

        return len(rows)

    async def unmapped_accounts(
        self,
        company_id: UUID,
    ) -> list[tuple[str, str | None]]:
        mapped_accounts = (
            select(
                FinanceAccountMapping.source_account_code
            )
            .where(
                FinanceAccountMapping.company_id
                == company_id
            )
        )

        statement = (
            select(
                GLTransaction.source_account_code,
                GLTransaction.source_account_name,
            )
            .join(FileUpload, FileUpload.id == GLTransaction.file_upload_id)
            .where(
                GLTransaction.company_id == company_id,
                GLTransaction.validation_status == "valid",
                GLTransaction.is_elimination.is_(False),
                FileUpload.is_active.is_(True),
                FileUpload.processing_status == "validated",
                FileUpload.document_type == "general_ledger",
                GLTransaction.source_account_code.not_in(mapped_accounts),
            )
            .distinct()
            .order_by(GLTransaction.source_account_code.asc())
        )

        result = await self.session.execute(
            statement
        )

        return [
            (row[0], row[1])
            for row in result.all()
        ]