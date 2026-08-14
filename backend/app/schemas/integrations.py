from __future__ import annotations
from datetime import datetime
from typing import Any, Literal
from pydantic import BaseModel, Field

Provider = Literal["xero", "zoho", "tally"]

class IntegrationConnectionOut(BaseModel):
    provider: Provider
    status: str
    external_tenant_id: str | None = None
    external_tenant_name: str | None = None
    last_synced_at: datetime | None = None
    last_sync_status: str | None = None
    last_sync_message: str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    configured: bool = True

class OAuthStartOut(BaseModel):
    authorization_url: str

class TenantSelection(BaseModel):
    tenant_id: str

class TallyBridgeTokenOut(BaseModel):
    bridge_token: str
    ingest_url: str

class TallyRecord(BaseModel):
    entity_type: str
    external_id: str
    name: str | None = None
    amount: float | None = None
    currency_code: str | None = None
    occurred_at: datetime | None = None
    source_updated_at: datetime | None = None
    payload: dict[str, Any] = Field(default_factory=dict)

class TallyPushRequest(BaseModel):
    records: list[TallyRecord] = Field(min_length=1, max_length=5000)

class MemoryCreate(BaseModel):
    title: str = Field(min_length=2, max_length=160)
    content: str = Field(min_length=2, max_length=4000)
    memory_type: str = Field(default="management_context", max_length=80)
    importance: str = Field(default="normal", max_length=32)

class MemoryOut(MemoryCreate):
    id: str
    created_at: datetime
