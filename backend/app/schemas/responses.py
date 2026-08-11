from typing import Generic, TypeVar

from pydantic import BaseModel, Field


DataT = TypeVar("DataT")


class APIResponse(BaseModel, Generic[DataT]):
    success: bool = True
    message: str
    data: DataT | None = None


class PaginatedResponse(BaseModel, Generic[DataT]):
    success: bool = True
    message: str
    count: int = Field(ge=0)
    limit: int = Field(ge=1)
    offset: int = Field(ge=0)
    data: list[DataT]


class ErrorResponse(BaseModel):
    success: bool = False
    message: str
    error_code: str
    details: dict | list | str | None = None