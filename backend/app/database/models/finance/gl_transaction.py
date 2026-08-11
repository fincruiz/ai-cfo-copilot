from datetime import date, datetime
from decimal import Decimal
from uuid import UUID

from sqlalchemy import (
    Boolean,
    Computed,
    Date,
    Integer,
    Numeric,
    Text,
    text,
)
from sqlalchemy.dialects.postgresql import JSONB, TIMESTAMP
from sqlalchemy.dialects.postgresql import UUID as PostgreSQLUUID
from sqlalchemy.orm import Mapped, mapped_column

from app.database.base import Base


class GLTransaction(Base):
    __tablename__ = "gl_transactions"
    __table_args__ = {"schema": "public"}

    id: Mapped[UUID] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        primary_key=True,
        server_default=text("gen_random_uuid()"),
    )

    company_id: Mapped[UUID] = mapped_column(PostgreSQLUUID(as_uuid=True), nullable=False, index=True)
    branch_id: Mapped[UUID | None] = mapped_column(PostgreSQLUUID(as_uuid=True))
    reporting_period_id: Mapped[UUID | None] = mapped_column(PostgreSQLUUID(as_uuid=True), index=True)
    file_upload_id: Mapped[UUID | None] = mapped_column(PostgreSQLUUID(as_uuid=True), index=True)

    transaction_date: Mapped[date] = mapped_column(Date, nullable=False, index=True)
    posting_date: Mapped[date | None] = mapped_column(Date)
    document_date: Mapped[date | None] = mapped_column(Date)

    document_number: Mapped[str | None] = mapped_column(Text)
    journal_number: Mapped[str | None] = mapped_column(Text)
    batch_number: Mapped[str | None] = mapped_column(Text)

    source_account_code: Mapped[str] = mapped_column(Text, nullable=False, index=True)
    source_account_name: Mapped[str | None] = mapped_column(Text)

    chart_account_id: Mapped[UUID | None] = mapped_column(PostgreSQLUUID(as_uuid=True))

    description: Mapped[str | None] = mapped_column(Text)
    reference: Mapped[str | None] = mapped_column(Text)

    customer_code: Mapped[str | None] = mapped_column(Text)
    supplier_code: Mapped[str | None] = mapped_column(Text)
    project_code: Mapped[str | None] = mapped_column(Text)
    cost_centre_code: Mapped[str | None] = mapped_column(Text)
    department_code: Mapped[str | None] = mapped_column(Text)

    debit: Mapped[Decimal] = mapped_column(
        Numeric,
        nullable=False,
        server_default=text("0"),
    )

    credit: Mapped[Decimal] = mapped_column(
        Numeric,
        nullable=False,
        server_default=text("0"),
    )

    # GENERATED ALWAYS in PostgreSQL
    net_amount: Mapped[Decimal | None] = mapped_column(
        Numeric,
        Computed("debit - credit", persisted=True),
    )

    currency_code: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'AUD'"),
    )

    exchange_rate: Mapped[Decimal] = mapped_column(
        Numeric,
        nullable=False,
        server_default=text("1"),
    )

    # GENERATED ALWAYS in PostgreSQL
    functional_currency_amount: Mapped[Decimal | None] = mapped_column(
        Numeric,
        Computed("net_amount * exchange_rate", persisted=True),
    )

    external_reference: Mapped[str | None] = mapped_column(Text)

    source_row_number: Mapped[int | None] = mapped_column(Integer)

    is_adjustment: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        server_default=text("false"),
    )

    is_elimination: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        server_default=text("false"),
    )

    is_intercompany: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        server_default=text("false"),
    )

    validation_status: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'valid'"),
    )

    validation_messages: Mapped[list] = mapped_column(
        JSONB,
        nullable=False,
        server_default=text("'[]'::jsonb"),
    )

    source_metadata: Mapped[dict] = mapped_column(
        JSONB,
        nullable=False,
        server_default=text("'{}'::jsonb"),
    )

    created_at: Mapped[datetime] = mapped_column(
        TIMESTAMP(timezone=True),
        nullable=False,
        server_default=text("now()"),
    )