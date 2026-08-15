from __future__ import annotations

import json
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


# Privacy boundary: only product-behaviour metadata is accepted. Customer finance,
# names, free-form AI prompts and ERP payloads are deliberately not valid properties.
ALLOWED_PROPERTY_KEYS = {
    "area", "feature", "group", "source", "view", "expanded",
    "widgets", "result_count", "role", "plan", "step", "status",
}


class UsageService:
    def __init__(self, session: AsyncSession):
        self.session = session

    async def record(
        self,
        *,
        company_id: UUID,
        user_id: UUID | None,
        event_name: str,
        path: str,
        session_id: str,
        properties: dict | None = None,
    ) -> None:
        safe = {
            key: value
            for key, value in (properties or {}).items()
            if key in ALLOWED_PROPERTY_KEYS
            and isinstance(value, (str, int, float, bool, list, type(None)))
        }
        try:
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.product_usage_events(
                        company_id, user_id, event_name, path, session_id, properties
                    ) VALUES (
                        :company_id, :user_id, :event_name, :path, :session_id,
                        CAST(:properties AS jsonb)
                    )
                    """
                ),
                {
                    "company_id": company_id,
                    "user_id": user_id,
                    "event_name": event_name[:80],
                    "path": path[:180],
                    "session_id": session_id[:100],
                    "properties": json.dumps(safe),
                },
            )
            await self.session.commit()
        except Exception:
            # Telemetry must never break the customer product experience.
            await self.session.rollback()

    async def summary(self, *, company_id: UUID, days: int = 30):
        try:
            result = await self.session.execute(
                text(
                    """
                    SELECT event_name, count(*)::int AS count,
                           count(DISTINCT user_id)::int AS users
                    FROM public.product_usage_events
                    WHERE company_id=:company_id
                      AND created_at >= now() - (:days || ' days')::interval
                    GROUP BY event_name
                    ORDER BY count(*) DESC
                    LIMIT 100
                    """
                ),
                {"company_id": company_id, "days": max(1, min(days, 365))},
            )
            return [dict(row._mapping) for row in result]
        except Exception:
            await self.session.rollback()
            return []
