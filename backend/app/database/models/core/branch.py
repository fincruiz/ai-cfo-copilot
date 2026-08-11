from datetime import datetime
from uuid import UUID

from sqlalchemy import Boolean, Text, UniqueConstraint, text
from sqlalchemy.dialects.postgresql import TIMESTAMP
from sqlalchemy.dialects.postgresql import UUID as PostgreSQLUUID
from sqlalchemy.orm import Mapped, mapped_column

from app.database.base import Base


class Branch(Base):
    __tablename__ = "branches"
    __table_args__ = (
        UniqueConstraint("company_id", "branch_code", name="uq_branches_company_code"),
        {"schema": "public"},
    )

    id: Mapped[UUID] = mapped_column(PostgreSQLUUID(as_uuid=True), primary_key=True, server_default=text("gen_random_uuid()"))
    company_id: Mapped[UUID] = mapped_column(PostgreSQLUUID(as_uuid=True), nullable=False, index=True)
    branch_code: Mapped[str] = mapped_column(Text, nullable=False)
    branch_name: Mapped[str] = mapped_column(Text, nullable=False)
    region: Mapped[str | None] = mapped_column(Text)
    review_status: Mapped[str] = mapped_column(Text, nullable=False, server_default=text("'accepted'"))
    source_value: Mapped[str | None] = mapped_column(Text)
    discovered_from_upload_id: Mapped[UUID | None] = mapped_column(PostgreSQLUUID(as_uuid=True))
    is_active: Mapped[bool] = mapped_column(Boolean, nullable=False, server_default=text("true"))
    created_at: Mapped[datetime] = mapped_column(TIMESTAMP(timezone=True), nullable=False, server_default=text("now()"))
    updated_at: Mapped[datetime] = mapped_column(TIMESTAMP(timezone=True), nullable=False, server_default=text("now()"))
