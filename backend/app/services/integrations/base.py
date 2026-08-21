from __future__ import annotations

import hashlib
import json
import secrets
from datetime import datetime, timedelta, timezone
from typing import Any
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.services.integrations.crypto import decrypt_secret, encrypt_secret


class IntegrationStore:
    def __init__(self, session: AsyncSession):
        self.session = session

    async def list_connections(self, company_id: UUID):
        rows = (
            await self.session.execute(
                text(
                    """
                    SELECT provider,status,external_tenant_id,external_tenant_name,last_synced_at,
                           last_sync_status,last_sync_message,metadata
                    FROM public.integration_connections
                    WHERE company_id=:company_id
                    ORDER BY provider
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().all()
        return [dict(row) for row in rows]

    async def get(self, company_id: UUID, provider: str):
        row = (
            await self.session.execute(
                text(
                    "SELECT * FROM public.integration_connections "
                    "WHERE company_id=:c AND provider=:p"
                ),
                {"c": company_id, "p": provider},
            )
        ).mappings().first()
        return dict(row) if row else None

    async def upsert_connection(
        self,
        *,
        company_id: UUID,
        provider: str,
        user_id: UUID | None = None,
        status: str = "connected",
        external_tenant_id=None,
        external_tenant_name=None,
        access_token=None,
        refresh_token=None,
        expires_in=None,
        metadata=None,
    ):
        expires_at = (
            datetime.now(timezone.utc)
            + timedelta(seconds=max(int(expires_in or 0) - 60, 0))
            if expires_in
            else None
        )
        await self.session.execute(
            text(
                """
                INSERT INTO public.integration_connections(
                    company_id,provider,status,external_tenant_id,external_tenant_name,
                    access_token_encrypted,refresh_token_encrypted,token_expires_at,
                    metadata,connected_by,updated_at
                )
                VALUES(:c,:p,:s,:tid,:tname,:at,:rt,:exp,CAST(:meta AS jsonb),:uid,now())
                ON CONFLICT(company_id,provider) DO UPDATE SET
                    status=EXCLUDED.status,
                    external_tenant_id=COALESCE(EXCLUDED.external_tenant_id,integration_connections.external_tenant_id),
                    external_tenant_name=COALESCE(EXCLUDED.external_tenant_name,integration_connections.external_tenant_name),
                    access_token_encrypted=COALESCE(EXCLUDED.access_token_encrypted,integration_connections.access_token_encrypted),
                    refresh_token_encrypted=COALESCE(EXCLUDED.refresh_token_encrypted,integration_connections.refresh_token_encrypted),
                    token_expires_at=COALESCE(EXCLUDED.token_expires_at,integration_connections.token_expires_at),
                    metadata=CASE
                        WHEN EXCLUDED.metadata='{}'::jsonb THEN integration_connections.metadata
                        ELSE EXCLUDED.metadata
                    END,
                    connected_by=COALESCE(EXCLUDED.connected_by,integration_connections.connected_by),
                    updated_at=now()
                """
            ),
            {
                "c": company_id,
                "p": provider,
                "s": status,
                "tid": external_tenant_id,
                "tname": external_tenant_name,
                "at": encrypt_secret(access_token),
                "rt": encrypt_secret(refresh_token),
                "exp": expires_at,
                "meta": json.dumps(metadata or {}),
                "uid": user_id,
            },
        )
        await self.session.commit()

    async def merge_metadata(
        self,
        company_id: UUID,
        provider: str,
        metadata: dict[str, Any],
        *,
        commit: bool = True,
    ) -> None:
        """Merge operational metadata without destroying OAuth tenant metadata."""
        await self.session.execute(
            text(
                """
                UPDATE public.integration_connections
                SET metadata = COALESCE(metadata,'{}'::jsonb) || CAST(:meta AS jsonb),
                    updated_at=now()
                WHERE company_id=:c AND provider=:p
                """
            ),
            {"c": company_id, "p": provider, "meta": json.dumps(metadata)},
        )
        if commit:
            await self.session.commit()

    async def save_oauth_state(self, company_id: UUID, user_id: UUID, provider: str):
        state = secrets.token_urlsafe(36)
        await self.session.execute(
            text(
                "INSERT INTO public.integration_oauth_states"
                "(state,company_id,user_id,provider,expires_at) "
                "VALUES(:s,:c,:u,:p,now()+interval '10 minutes')"
            ),
            {"s": state, "c": company_id, "u": user_id, "p": provider},
        )
        await self.session.commit()
        return state

    async def consume_oauth_state(self, state: str, provider: str):
        row = (
            await self.session.execute(
                text(
                    "DELETE FROM public.integration_oauth_states "
                    "WHERE state=:s AND provider=:p AND expires_at>now() "
                    "RETURNING company_id,user_id"
                ),
                {"s": state, "p": provider},
            )
        ).mappings().first()
        await self.session.commit()
        return dict(row) if row else None

    async def credentials(self, company_id: UUID, provider: str):
        connection = await self.get(company_id, provider)
        if not connection:
            return None
        connection["access_token"] = decrypt_secret(connection.get("access_token_encrypted"))
        connection["refresh_token"] = decrypt_secret(connection.get("refresh_token_encrypted"))
        return connection

    async def replace_records_snapshot(
        self,
        company_id: UUID,
        provider: str,
        entity_type: str,
        records: list[dict[str, Any]],
        *,
        commit: bool = False,
    ) -> int:
        """Replace one fully-fetched entity snapshot atomically in the current transaction.

        Integration source tables are a synchronized copy, not an append-only ledger.  Deleting
        rows that disappeared from the provider prevents stale/deleted source transactions from
        leaking into later finance-truth activation.
        """
        await self.session.execute(
            text(
                "DELETE FROM public.integration_records "
                "WHERE company_id=:c AND provider=:p AND entity_type=:e"
            ),
            {"c": company_id, "p": provider, "e": entity_type},
        )

        if records:
            parameters = [
                {
                    "c": company_id,
                    "p": provider,
                    "e": entity_type,
                    "x": str(record["external_id"]),
                    "o": record.get("occurred_at"),
                    "n": record.get("name"),
                    "a": record.get("amount"),
                    "cc": record.get("currency_code"),
                    "payload": json.dumps(record.get("payload") or {}, default=str),
                    "su": record.get("source_updated_at"),
                }
                for record in records
            ]
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.integration_records(
                        company_id,provider,entity_type,external_id,occurred_at,name,
                        amount,currency_code,payload,source_updated_at,synced_at
                    )
                    VALUES(:c,:p,:e,:x,:o,:n,:a,:cc,CAST(:payload AS jsonb),:su,now())
                    """
                ),
                parameters,
            )

        if commit:
            await self.session.commit()
        return len(records)

    async def replace_records(
        self,
        company_id: UUID,
        provider: str,
        entity_type: str,
        records: list[dict[str, Any]],
    ):
        """Upsert records without deleting unseen rows.

        This is retained for streaming/chunked sources such as the Tally bridge.
        Full-provider synchronizations should use ``replace_records_snapshot`` so
        deleted source rows cannot remain stale in FinCruiz.
        """
        if records:
            parameters = [
                {
                    "c": company_id,
                    "p": provider,
                    "e": entity_type,
                    "x": str(record["external_id"]),
                    "o": record.get("occurred_at"),
                    "n": record.get("name"),
                    "a": record.get("amount"),
                    "cc": record.get("currency_code"),
                    "payload": json.dumps(record.get("payload") or {}, default=str),
                    "su": record.get("source_updated_at"),
                }
                for record in records
            ]
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.integration_records(
                        company_id,provider,entity_type,external_id,occurred_at,name,
                        amount,currency_code,payload,source_updated_at,synced_at
                    )
                    VALUES(:c,:p,:e,:x,:o,:n,:a,:cc,CAST(:payload AS jsonb),:su,now())
                    ON CONFLICT(company_id,provider,entity_type,external_id) DO UPDATE SET
                        occurred_at=EXCLUDED.occurred_at, name=EXCLUDED.name, amount=EXCLUDED.amount,
                        currency_code=EXCLUDED.currency_code, payload=EXCLUDED.payload,
                        source_updated_at=EXCLUDED.source_updated_at, synced_at=now()
                    """
                ),
                parameters,
            )
        await self.session.commit()

    async def clear_records(
        self,
        company_id: UUID,
        provider: str,
        entity_type: str,
        *,
        commit: bool = False,
    ) -> int:
        result = await self.session.execute(
            text(
                "DELETE FROM public.integration_records "
                "WHERE company_id=:c AND provider=:p AND entity_type=:e"
            ),
            {"c": company_id, "p": provider, "e": entity_type},
        )
        if commit:
            await self.session.commit()
        return int(result.rowcount or 0)

    async def records(
        self,
        company_id: UUID,
        provider: str,
        entity_type: str,
    ) -> list[dict[str, Any]]:
        rows = (
            await self.session.execute(
                text(
                    """
                    SELECT id, external_id, occurred_at, name, amount, currency_code,
                           payload, source_updated_at, synced_at
                    FROM public.integration_records
                    WHERE company_id=:c AND provider=:p AND entity_type=:e
                    ORDER BY occurred_at NULLS LAST, external_id
                    """
                ),
                {"c": company_id, "p": provider, "e": entity_type},
            )
        ).mappings().all()
        return [dict(row) for row in rows]

    async def mark_sync(
        self,
        company_id: UUID,
        provider: str,
        status: str,
        message: str,
        *,
        commit: bool = True,
    ):
        await self.session.execute(
            text(
                """
                UPDATE public.integration_connections
                SET last_synced_at=now(),last_sync_status=:s,last_sync_message=:m,updated_at=now()
                WHERE company_id=:c AND provider=:p
                """
            ),
            {"s": status, "m": message[:1000], "c": company_id, "p": provider},
        )
        if commit:
            await self.session.commit()

    async def disconnect(self, company_id: UUID, provider: str, delete_records: bool = True):
        if delete_records:
            await self.session.execute(
                text(
                    "DELETE FROM public.integration_records "
                    "WHERE company_id=:c AND provider=:p"
                ),
                {"c": company_id, "p": provider},
            )
        await self.session.execute(
            text(
                "DELETE FROM public.integration_connections "
                "WHERE company_id=:c AND provider=:p"
            ),
            {"c": company_id, "p": provider},
        )
        await self.session.commit()

    async def tally_bridge_token(self, company_id: UUID, user_id: UUID):
        token = "fct_" + secrets.token_urlsafe(36)
        digest = hashlib.sha256(token.encode()).hexdigest()
        await self.session.execute(
            text(
                """
                INSERT INTO public.integration_connections(
                    company_id,provider,status,bridge_token_hash,connected_by,metadata,updated_at
                )
                VALUES(:c,'tally','awaiting_bridge',:h,:u,'{}'::jsonb,now())
                ON CONFLICT(company_id,provider) DO UPDATE SET
                    status='awaiting_bridge',bridge_token_hash=EXCLUDED.bridge_token_hash,
                    connected_by=:u,updated_at=now()
                """
            ),
            {"c": company_id, "h": digest, "u": user_id},
        )
        await self.session.commit()
        return token

    async def company_for_bridge(self, token: str):
        digest = hashlib.sha256(token.encode()).hexdigest()
        row = (
            await self.session.execute(
                text(
                    "SELECT company_id FROM public.integration_connections "
                    "WHERE provider='tally' AND bridge_token_hash=:h"
                ),
                {"h": digest},
            )
        ).first()
        return row[0] if row else None
