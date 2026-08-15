-- P8.1 privacy-safe product analytics.
-- Stores feature-usage metadata only. Do not store finance values, transaction text,
-- customer/vendor names, uploaded filenames, ERP payloads, or AI prompt/answer content.

CREATE TABLE IF NOT EXISTS public.product_usage_events (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    user_id uuid NULL,
    event_name text NOT NULL,
    path text NOT NULL DEFAULT '',
    session_id text NOT NULL DEFAULT '',
    properties jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_product_usage_company_created
    ON public.product_usage_events(company_id, created_at DESC);
CREATE INDEX IF NOT EXISTS ix_product_usage_company_event
    ON public.product_usage_events(company_id, event_name, created_at DESC);

ALTER TABLE public.product_usage_events ENABLE ROW LEVEL SECURITY;
-- App access is through the authenticated backend service. No browser-direct policies are added.
