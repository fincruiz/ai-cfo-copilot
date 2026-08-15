from __future__ import annotations

from typing import Any

from pydantic import BaseModel, Field


class UsageEventCreate(BaseModel):
    event_name: str = Field(min_length=2, max_length=80)
    path: str = Field(default="", max_length=180)
    session_id: str = Field(default="", max_length=100)
    properties: dict[str, Any] = Field(default_factory=dict)
