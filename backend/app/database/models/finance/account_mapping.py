from datetime import datetime
from uuid import UUID

from sqlalchemy import Boolean, Integer, Text, UniqueConstraint, text
from sqlalchemy.dialects.postgresql import TIMESTAMP
from sqlalchemy.dialects.postgresql import UUID as PostgreSQLUUID
from sqlalchemy.orm import Mapped, mapped_column

from app.database.base import Base


class FinanceAccountMapping(Base):
    __tablename__ = "finance_account_mappings"
    __table_args__ = (
        UniqueConstraint("company_id", "source_account_code", name="uq_finance_mapping_company_account"),
        {"schema": "public"},
    )

    id: Mapped[UUID] = mapped_column(PostgreSQLUUID(as_uuid=True), primary_key=True, server_default=text("gen_random_uuid()"))
    company_id: Mapped[UUID] = mapped_column(PostgreSQLUUID(as_uuid=True), nullable=False, index=True)
    source_account_code: Mapped[str] = mapped_column(Text, nullable=False)
    source_account_name: Mapped[str | None] = mapped_column(Text, nullable=True)
    statement: Mapped[str] = mapped_column(Text, nullable=False)
    reporting_group: Mapped[str] = mapped_column(Text, nullable=False)
    reporting_subgroup: Mapped[str | None] = mapped_column(Text, nullable=True)
    sign_convention: Mapped[str] = mapped_column(Text, nullable=False, server_default=text("'positive'"))
    display_order: Mapped[int | None] = mapped_column(Integer, nullable=True)
    is_confirmed: Mapped[bool] = mapped_column(Boolean, nullable=False, server_default=text("false"))
    created_at: Mapped[datetime] = mapped_column(TIMESTAMP(timezone=True), nullable=False, server_default=text("now()"))
    updated_at: Mapped[datetime] = mapped_column(TIMESTAMP(timezone=True), nullable=False, server_default=text("now()"))
