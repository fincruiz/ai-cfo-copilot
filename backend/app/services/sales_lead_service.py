from __future__ import annotations

from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


class SalesLeadService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def create_demo_lead(
        self,
        *,
        name: str,
        work_email: str,
        company_name: str,
        role: str | None,
        persona: str | None,
        country: str | None,
        team_size: str | None,
        message: str | None,
        source_path: str | None,
        referrer_host: str | None,
    ) -> UUID:
        lead_id = (
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.sales_leads(
                        lead_type,name,work_email,company_name,role,persona,country,
                        team_size,message,source_path,referrer_host,status
                    ) VALUES (
                        'book_demo',:name,:work_email,:company_name,:role,:persona,:country,
                        :team_size,:message,:source_path,:referrer_host,'new'
                    )
                    RETURNING id
                    """
                ),
                {
                    "name": name[:120],
                    "work_email": work_email[:254],
                    "company_name": company_name[:180],
                    "role": (role or "")[:120] or None,
                    "persona": (persona or "")[:40] or None,
                    "country": (country or "")[:100] or None,
                    "team_size": (team_size or "")[:80] or None,
                    "message": (message or "")[:1200] or None,
                    "source_path": (source_path or "")[:250] or None,
                    "referrer_host": (referrer_host or "")[:200] or None,
                },
            )
        ).scalar_one()
        await self.session.commit()
        return lead_id
