from __future__ import annotations

import json
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


class AuditService:
    def __init__(self, session: AsyncSession):
        self.session = session

    async def _ensure_table(self) -> None:
        await self.session.execute(
            text(
                """
                CREATE TABLE IF NOT EXISTS public.audit_events (
                    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
                    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
                    user_id uuid NULL,
                    action text NOT NULL,
                    module text NOT NULL,
                    summary text NOT NULL,
                    metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
                    created_at timestamptz NOT NULL DEFAULT now()
                )
                """
            )
        )
        await self.session.execute(
            text(
                "CREATE INDEX IF NOT EXISTS ix_audit_events_company_created "
                "ON public.audit_events(company_id, created_at DESC)"
            )
        )

    async def record(
        self,
        *,
        company_id: UUID,
        user_id: UUID | None,
        action: str,
        module: str,
        summary: str,
        metadata: dict | None = None,
        commit: bool = False,
    ):
        try:
            await self._ensure_table()
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.audit_events(
                        company_id,user_id,action,module,summary,metadata
                    )
                    VALUES (
                        :company_id,:user_id,:action,:module,:summary,
                        CAST(:metadata AS jsonb)
                    )
                    """
                ),
                {
                    "company_id": company_id,
                    "user_id": user_id,
                    "action": action,
                    "module": module,
                    "summary": summary,
                    "metadata": json.dumps(metadata or {}),
                },
            )
            if commit:
                await self.session.commit()
        except Exception:
            # Audit logging is valuable, but it must never turn an otherwise valid
            # finance action into a customer-facing 500 response.
            await self.session.rollback()
            return None

    async def list(self, *, company_id: UUID, limit: int = 100):
        try:
            await self._ensure_table()
            result = await self.session.execute(
                text(
                    """
                    SELECT id,user_id,action,module,summary,metadata,created_at
                    FROM public.audit_events
                    WHERE company_id=:company_id
                    ORDER BY created_at DESC
                    LIMIT :limit
                    """
                ),
                {"company_id": company_id, "limit": min(max(limit, 1), 250)},
            )
            await self.session.commit()
            return [dict(row._mapping) for row in result]
        except Exception:
            await self.session.rollback()
            return []
