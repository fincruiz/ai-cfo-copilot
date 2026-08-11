from datetime import datetime
from uuid import UUID

from sqlalchemy import BigInteger, Boolean, Integer, Text, text
from sqlalchemy.dialects.postgresql import JSONB, TIMESTAMP
from sqlalchemy.dialects.postgresql import UUID as PostgreSQLUUID
from sqlalchemy.orm import Mapped, mapped_column

from app.database.base import Base


class FileUpload(Base):
    __tablename__ = "file_uploads"
    __table_args__ = {"schema": "public"}

    id: Mapped[UUID] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        primary_key=True,
        server_default=text("gen_random_uuid()"),
    )

    company_id: Mapped[UUID] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        nullable=False,
    )

    reporting_period_id: Mapped[UUID | None] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        nullable=True,
    )

    file_name: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    original_file_name: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    storage_bucket: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'company-uploads'"),
    )

    storage_path: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    mime_type: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    file_size_bytes: Mapped[int | None] = mapped_column(
        BigInteger,
        nullable=True,
    )

    document_type: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    source_system: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    processing_status: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'uploaded'"),
    )

    is_active: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        server_default=text("false"),
        index=True,
    )

    superseded_at: Mapped[datetime | None] = mapped_column(
        TIMESTAMP(timezone=True),
        nullable=True,
    )

    row_count: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    valid_row_count: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    invalid_row_count: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    validation_summary: Mapped[dict] = mapped_column(
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )

    column_mapping: Mapped[dict] = mapped_column(
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )

    processing_metadata: Mapped[dict] = mapped_column(
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )

    uploaded_by: Mapped[UUID | None] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        nullable=True,
    )

    processed_at: Mapped[datetime | None] = mapped_column(
        TIMESTAMP(timezone=True),
        nullable=True,
    )

    created_at: Mapped[datetime] = mapped_column(
        TIMESTAMP(timezone=True),
        nullable=False,
        server_default=text("now()"),
    )

    updated_at: Mapped[datetime] = mapped_column(
        TIMESTAMP(timezone=True),
        nullable=False,
        server_default=text("now()"),
    )