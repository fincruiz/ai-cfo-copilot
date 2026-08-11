from collections.abc import Mapping
from typing import Any, Generic, TypeVar
from uuid import UUID

from sqlalchemy import func, select
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.orm import DeclarativeBase


ModelT = TypeVar(
    "ModelT",
    bound=DeclarativeBase,
)


class BaseRepository(Generic[ModelT]):
    def __init__(
        self,
        session: AsyncSession,
        model: type[ModelT],
    ) -> None:
        self.session = session
        self.model = model

    async def get_by_id(
        self,
        record_id: UUID,
    ) -> ModelT | None:
        statement = select(self.model).where(
            self.model.id == record_id
        )

        result = await self.session.execute(statement)

        return result.scalar_one_or_none()

    async def list_records(
        self,
        *,
        limit: int = 100,
        offset: int = 0,
        order_by: Any | None = None,
        filters: Mapping[str, Any] | None = None,
    ) -> list[ModelT]:
        statement = select(self.model)

        if filters:
            for field_name, value in filters.items():
                model_field = getattr(
                    self.model,
                    field_name,
                    None,
                )

                if model_field is None:
                    raise ValueError(
                        f"Invalid filter field: {field_name}"
                    )

                statement = statement.where(
                    model_field == value
                )

        if order_by is not None:
            statement = statement.order_by(order_by)

        statement = statement.limit(limit).offset(offset)

        result = await self.session.execute(statement)

        return list(result.scalars().all())

    async def count_records(
        self,
        *,
        filters: Mapping[str, Any] | None = None,
    ) -> int:
        statement = select(
            func.count()
        ).select_from(self.model)

        if filters:
            for field_name, value in filters.items():
                model_field = getattr(
                    self.model,
                    field_name,
                    None,
                )

                if model_field is None:
                    raise ValueError(
                        f"Invalid filter field: {field_name}"
                    )

                statement = statement.where(
                    model_field == value
                )

        result = await self.session.execute(statement)

        return int(result.scalar_one())

    async def create(
        self,
        values: Mapping[str, Any],
    ) -> ModelT:
        record = self.model(**dict(values))

        self.session.add(record)

        await self.session.flush()
        await self.session.refresh(record)

        return record

    async def update(
        self,
        record: ModelT,
        values: Mapping[str, Any],
    ) -> ModelT:
        for field_name, value in values.items():
            if not hasattr(record, field_name):
                raise ValueError(
                    f"Invalid update field: {field_name}"
                )

            setattr(record, field_name, value)

        await self.session.flush()
        await self.session.refresh(record)

        return record

    async def delete(
        self,
        record: ModelT,
    ) -> None:
        await self.session.delete(record)
        await self.session.flush()