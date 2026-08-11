from datetime import datetime
from decimal import Decimal
from uuid import UUID

from sqlalchemy import Boolean, Integer, Numeric, SmallInteger, Text, text
from sqlalchemy.dialects.postgresql import TIMESTAMP
from sqlalchemy.dialects.postgresql import UUID as PostgreSQLUUID
from sqlalchemy.orm import Mapped, mapped_column

from app.database.base import Base


class Company(Base):
    __tablename__ = "companies"
    __table_args__ = {"schema": "public"}

    id: Mapped[UUID] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
        primary_key=True,
        server_default=text("gen_random_uuid()"),
    )

    legal_name: Mapped[str] = mapped_column(
        Text,
        nullable=False,
    )

    trading_name: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    abn: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    country_code: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'AU'"),
    )

    currency_code: Mapped[str] = mapped_column(
        Text,
        nullable=False,
        server_default=text("'AUD'"),
    )

    financial_year_end_month: Mapped[int] = mapped_column(
        SmallInteger,
        nullable=False,
        server_default=text("6"),
    )

    industry: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    business_model: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    employee_count: Mapped[int | None] = mapped_column(
        Integer,
        nullable=True,
    )

    annual_revenue: Mapped[Decimal | None] = mapped_column(
        Numeric,
        nullable=True,
    )

    logo_path: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    website_url: Mapped[str | None] = mapped_column(
        Text,
        nullable=True,
    )

    is_active: Mapped[bool] = mapped_column(
        Boolean,
        nullable=False,
        server_default=text("true"),
    )

    # The database already enforces the foreign key to auth.users.
    # We avoid mapping auth.users at this stage.
    created_by: Mapped[UUID | None] = mapped_column(
        PostgreSQLUUID(as_uuid=True),
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