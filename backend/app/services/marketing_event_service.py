from __future__ import annotations
from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

ALLOWED_EVENTS = {
    "homepage_viewed", "homepage_hero_demo_clicked", "homepage_hero_signup_clicked",
    "homepage_ai_question_submitted", "homepage_ai_signup_clicked",
    "homepage_product_tour_clicked", "homepage_capability_demo_clicked",
    "homepage_pricing_clicked", "homepage_final_demo_clicked",
    "homepage_final_signup_clicked",
}

class MarketingEventService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def record(self, *, event_name: str, session_id: str, path: str, referrer_host: str | None, properties: dict) -> None:
        if event_name not in ALLOWED_EVENTS:
            return
        safe_properties = {
            str(k)[:60]: v for k, v in (properties or {}).items()
            if isinstance(v, (str, int, float, bool)) and str(k) not in {"question", "email", "name", "description"}
        }
        import json
        await self.session.execute(text("""
            INSERT INTO public.marketing_events(event_name,session_id,path,referrer_host,properties)
            VALUES(:event_name,:session_id,:path,:referrer_host,CAST(:properties AS jsonb))
        """), {
            "event_name": event_name, "session_id": session_id[:120], "path": path[:250],
            "referrer_host": (referrer_host or "")[:200] or None,
            "properties": json.dumps(safe_properties),
        })
        await self.session.commit()

    async def funnel(self, *, days: int = 30) -> dict:
        row=(await self.session.execute(text("""
          SELECT
            count(DISTINCT session_id)::int AS visitors,
            count(*) FILTER(WHERE event_name='homepage_hero_demo_clicked')::int AS hero_demo,
            count(*) FILTER(WHERE event_name='homepage_hero_signup_clicked')::int AS hero_signup,
            count(*) FILTER(WHERE event_name='homepage_ai_question_submitted')::int AS ai_questions,
            count(*) FILTER(WHERE event_name='homepage_ai_signup_clicked')::int AS ai_signup,
            count(*) FILTER(WHERE event_name='homepage_pricing_clicked')::int AS pricing,
            count(*) FILTER(WHERE event_name='homepage_final_signup_clicked')::int AS final_signup
          FROM public.marketing_events
          WHERE created_at >= now() - make_interval(days => :days)
        """), {"days": max(1,min(days,365))})).mappings().one()
        return dict(row)
