from typing import Generic, TypeVar
from uuid import UUID

from sqlalchemy.orm import DeclarativeBase

from app.core.exceptions import ResourceNotFoundError
from app.repositories.base import BaseRepository


ModelT = TypeVar(
    "ModelT",
    bound=DeclarativeBase,
)


class BaseService(Generic[ModelT]):
    def __init__(
        self,
        repository: BaseRepository[ModelT],
        resource_name: str,
    ) -> None:
        self.repository = repository
        self.resource_name = resource_name

    async def get_or_404(
        self,
        record_id: UUID,
    ) -> ModelT:
        record = await self.repository.get_by_id(
            record_id
        )

        if record is None:
            raise ResourceNotFoundError(
                resource_name=self.resource_name,
                resource_id=str(record_id),
            )

        return record