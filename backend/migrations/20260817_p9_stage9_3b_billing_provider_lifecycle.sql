-- P9 Stage 9.3B — Provider-neutral recurring billing lifecycle.

ALTER TABLE public.company_subscriptions
    ADD COLUMN IF NOT EXISTS last_checkout_id text,
    ADD COLUMN IF NOT EXISTS payment_failure_at timestamptz,
    ADD COLUMN IF NOT EXISTS last_billing_event_at timestamptz;

CREATE TABLE IF NOT EXISTS public.billing_events (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    provider text NOT NULL CHECK (provider IN ('stripe','razorpay')),
    provider_event_id text NOT NULL,
    event_type text NOT NULL,
    company_id uuid NULL REFERENCES public.companies(id) ON DELETE SET NULL,
    payload jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE(provider, provider_event_id)
);

CREATE INDEX IF NOT EXISTS idx_billing_events_company_created
    ON public.billing_events(company_id, created_at DESC);

ALTER TABLE public.billing_events ENABLE ROW LEVEL SECURITY;

COMMENT ON TABLE public.billing_events IS
'Idempotency/audit log for verified payment-provider webhooks. Raw provider payloads are server-only and not exposed directly to browser clients.';
