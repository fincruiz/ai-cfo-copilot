from __future__ import annotations

from datetime import UTC, datetime
from uuid import UUID

from sqlalchemy import select, update
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.finance.file_upload import FileUpload
from app.repositories.base import BaseRepository


class FileUploadRepository(BaseRepository[FileUpload]):
    def __init__(self, session: AsyncSession) -> None:
        super().__init__(session=session, model=FileUpload)

    async def list_company_uploads(
        self,
        *,
        company_id: UUID,
        document_type: str | None = None,
        limit: int = 50,
        offset: int = 0,
    ) -> list[FileUpload]:
        statement = (
            select(FileUpload)
            .where(FileUpload.company_id == company_id)
            .order_by(FileUpload.created_at.desc())
            .limit(limit)
            .offset(offset)
        )

        if document_type:
            statement = statement.where(
                FileUpload.document_type == document_type
            )

        result = await self.session.execute(statement)
        return list(result.scalars().all())

    async def deactivate_active_datasets(
        self,
        *,
        company_id: UUID,
        document_type: str,
        reporting_period_id: UUID | None,
        exclude_upload_id: UUID | None = None,
    ) -> int:
        conditions = [
            FileUpload.company_id == company_id,
            FileUpload.document_type == document_type,
            FileUpload.is_active.is_(True),
        ]

        if reporting_period_id is None:
            conditions.append(FileUpload.reporting_period_id.is_(None))
        else:
            conditions.append(
                FileUpload.reporting_period_id == reporting_period_id
            )

        if exclude_upload_id is not None:
            conditions.append(FileUpload.id != exclude_upload_id)

        statement = (
            update(FileUpload)
            .where(*conditions)
            .values(
                is_active=False,
                superseded_at=datetime.now(UTC),
                updated_at=datetime.now(UTC),
            )
        )

        result = await self.session.execute(statement)
        return int(result.rowcount or 0)
