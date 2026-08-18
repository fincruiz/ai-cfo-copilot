from __future__ import annotations

from uuid import UUID
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


class BetaFeedbackService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def create(
        self, *, company_id: UUID, user_id: UUID, category: str, severity: str,
        title: str, description: str, path: str, app_version: str | None,
        browser: str | None, viewport: str | None, request_id: str | None,
        attachment_mime: str | None, attachment_bytes: bytes | None,
    ) -> dict:
        role = await self.session.scalar(text(
            "SELECT role::text FROM public.company_members WHERE company_id=:c AND user_id=:u AND is_active=true"
        ), {"c": company_id, "u": user_id})
        row = (await self.session.execute(text("""
            INSERT INTO public.beta_feedback(
              company_id,user_id,category,severity,title,description,path,user_role,
              app_version,browser,viewport,request_id,attachment_mime,attachment_bytes
            ) VALUES(
              :c,:u,:category,:severity,:title,:description,:path,:role,
              :version,:browser,:viewport,:request_id,:attachment_mime,:attachment_bytes
            )
            RETURNING id,category,severity,status,title,description,path,user_role,
                      app_version,browser,viewport,request_id,
                      (attachment_bytes IS NOT NULL) AS has_attachment,created_at,updated_at
        """), {
            "c": company_id,"u": user_id,"category": category,"severity": severity,
            "title": title[:180],"description": description[:6000],"path": path[:250],
            "role": role,"version": (app_version or "")[:80] or None,
            "browser": (browser or "")[:500] or None,"viewport": (viewport or "")[:80] or None,
            "request_id": (request_id or "")[:120] or None,
            "attachment_mime": attachment_mime,"attachment_bytes": attachment_bytes,
        })).mappings().one()
        await self.session.commit()
        return dict(row)

    async def list(self, *, company_id: UUID, status: str | None = None, severity: str | None = None, limit: int = 200) -> list[dict]:
        rows=(await self.session.execute(text("""
            SELECT bf.id,bf.user_id,bf.category,bf.severity,bf.status,bf.title,bf.description,
                   bf.path,bf.user_role,bf.app_version,bf.browser,bf.viewport,bf.request_id,
                   (bf.attachment_bytes IS NOT NULL) AS has_attachment,
                   bf.resolution_notes,bf.created_at,bf.updated_at,
                   COALESCE(p.full_name,'Beta tester') AS reporter_name
            FROM public.beta_feedback bf
            LEFT JOIN public.profiles p ON p.id=bf.user_id
            WHERE bf.company_id=:c
              AND (:status IS NULL OR bf.status=:status)
              AND (:severity IS NULL OR bf.severity=:severity)
            ORDER BY CASE bf.severity WHEN 'p0' THEN 0 WHEN 'p1' THEN 1 ELSE 2 END,
                     bf.created_at DESC
            LIMIT :limit
        """), {"c":company_id,"status":status,"severity":severity,"limit":max(1,min(limit,500))})).mappings().all()
        return [dict(x) for x in rows]

    async def update(self, *, company_id: UUID, feedback_id: UUID, status: str, resolution_notes: str | None) -> dict | None:
        row=(await self.session.execute(text("""
            UPDATE public.beta_feedback
            SET status=:status,resolution_notes=:notes,updated_at=now()
            WHERE id=:id AND company_id=:c
            RETURNING id,status,resolution_notes,updated_at
        """), {"id":feedback_id,"c":company_id,"status":status,"notes":resolution_notes})).mappings().first()
        await self.session.commit()
        return dict(row) if row else None

    async def attachment(self, *, company_id: UUID, feedback_id: UUID):
        return (await self.session.execute(text("""
            SELECT attachment_mime,attachment_bytes
            FROM public.beta_feedback
            WHERE id=:id AND company_id=:c AND attachment_bytes IS NOT NULL
        """), {"id":feedback_id,"c":company_id})).mappings().first()

    async def summary(self, *, company_id: UUID) -> dict:
        row=(await self.session.execute(text("""
            SELECT count(*)::int total,
              count(*) FILTER(WHERE status='open')::int open,
              count(*) FILTER(WHERE severity='p0' AND status IN ('open','reviewing'))::int p0_open,
              count(*) FILTER(WHERE severity='p1' AND status IN ('open','reviewing'))::int p1_open,
              count(*) FILTER(WHERE status='fixed')::int fixed
            FROM public.beta_feedback WHERE company_id=:c
        """), {"c":company_id})).mappings().one()
        return dict(row)
