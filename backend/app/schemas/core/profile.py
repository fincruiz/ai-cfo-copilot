from uuid import UUID

from pydantic import BaseModel, ConfigDict


class ProfileResponse(BaseModel):
    model_config = ConfigDict(from_attributes=True)

    id: UUID
    full_name: str | None = None
    job_title: str | None = None
    phone: str | None = None
    avatar_path: str | None = None


class UpdateProfileRequest(BaseModel):
    full_name: str | None = None
    job_title: str | None = None
    phone: str | None = None
    avatar_path: str |None = None