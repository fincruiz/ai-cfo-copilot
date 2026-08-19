-- P9 Stage 9.5.1 — privacy-safe public homepage conversion telemetry.
CREATE TABLE IF NOT EXISTS public.marketing_events (
  id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
  event_name text NOT NULL,
  session_id text NOT NULL,
  path text NOT NULL DEFAULT '/',
  referrer_host text NULL,
  properties jsonb NOT NULL DEFAULT '{}'::jsonb,
  created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_marketing_events_name_created
  ON public.marketing_events(event_name,created_at DESC);
CREATE INDEX IF NOT EXISTS ix_marketing_events_session_created
  ON public.marketing_events(session_id,created_at DESC);
ALTER TABLE public.marketing_events ENABLE ROW LEVEL SECURITY;
-- No browser policies: events are accepted only through the validated backend endpoint.
