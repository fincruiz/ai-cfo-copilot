CREATE TABLE IF NOT EXISTS public.audit_events (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    user_id uuid NULL,
    action text NOT NULL,
    module text NOT NULL,
    summary text NOT NULL,
    metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_audit_events_company_created ON public.audit_events(company_id, created_at DESC);
