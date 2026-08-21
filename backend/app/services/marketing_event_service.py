from __future__ import annotations

import json

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession


ALLOWED_EVENTS = {
    # Homepage baseline
    "homepage_viewed",
    "homepage_hero_demo_clicked",
    "homepage_hero_signup_clicked",
    "homepage_ai_question_submitted",
    "homepage_ai_signup_clicked",
    "homepage_product_tour_clicked",
    "homepage_capability_demo_clicked",
    "homepage_pricing_clicked",  # legacy name retained for existing events
    "homepage_pricing_cta_clicked",
    "homepage_reporting_cta_clicked",
    "homepage_forecasting_cta_clicked",
    "homepage_persona_changed",
    "homepage_final_demo_clicked",
    "homepage_final_signup_clicked",
    "homepage_book_demo_clicked",
    "demo_book_demo_clicked",
    "demo_lead_submitted",
    # Public guided demo
    "demo_viewed",
    "demo_audience_changed",
    "demo_presenter_mode_toggled",
    "demo_guided_scene_clicked",
    "demo_scenario_clicked",
    "demo_question_submitted",
    "demo_signup_clicked",
    "demo_pricing_clicked",
}

# Anonymous marketing telemetry must never become a back door for storing
# prospect questions, contact details or other uncontrolled free text.
DENIED_PROPERTY_KEYS = {"question", "email", "name", "description", "message", "notes"}


class MarketingEventService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def record(
        self,
        *,
        event_name: str,
        session_id: str,
        path: str,
        referrer_host: str | None,
        properties: dict,
    ) -> None:
        if event_name not in ALLOWED_EVENTS:
            return

        safe_properties = {
            str(key)[:60]: value
            for key, value in (properties or {}).items()
            if isinstance(value, (str, int, float, bool))
            and str(key).lower() not in DENIED_PROPERTY_KEYS
        }

        await self.session.execute(
            text(
                """
                INSERT INTO public.marketing_events(
                    event_name, session_id, path, referrer_host, properties
                )
                VALUES(
                    :event_name, :session_id, :path, :referrer_host,
                    CAST(:properties AS jsonb)
                )
                """
            ),
            {
                "event_name": event_name,
                "session_id": session_id[:120],
                "path": path[:250],
                "referrer_host": (referrer_host or "")[:200] or None,
                "properties": json.dumps(safe_properties),
            },
        )
        await self.session.commit()

    async def funnel(self, *, days: int = 30) -> dict:
        row = (
            await self.session.execute(
                text(
                    """
                    SELECT
                        count(DISTINCT session_id)::int AS visitors,
                        count(*) FILTER(
                            WHERE event_name='homepage_hero_demo_clicked'
                        )::int AS hero_demo,
                        count(*) FILTER(
                            WHERE event_name='homepage_hero_signup_clicked'
                        )::int AS hero_signup,
                        count(*) FILTER(
                            WHERE event_name='homepage_ai_question_submitted'
                        )::int AS ai_questions,
                        count(*) FILTER(
                            WHERE event_name='homepage_ai_signup_clicked'
                        )::int AS ai_signup,
                        count(*) FILTER(
                            WHERE event_name IN (
                                'homepage_pricing_clicked',
                                'homepage_pricing_cta_clicked'
                            )
                        )::int AS pricing,
                        count(*) FILTER(
                            WHERE event_name='homepage_final_signup_clicked'
                        )::int AS final_signup,
                        count(*) FILTER(
                            WHERE event_name='demo_viewed'
                        )::int AS demo_views,
                        count(*) FILTER(
                            WHERE event_name='demo_question_submitted'
                        )::int AS demo_questions,
                        count(*) FILTER(
                            WHERE event_name='demo_signup_clicked'
                        )::int AS demo_signup
                    FROM public.marketing_events
                    WHERE created_at >= now() - make_interval(days => :days)
                    """
                ),
                {"days": max(1, min(days, 365))},
            )
        ).mappings().one()
        return dict(row)
