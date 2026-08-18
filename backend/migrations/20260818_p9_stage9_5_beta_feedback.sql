-- P9 Stage 9.5 — Beta feedback and testing console.
-- Screenshots are stored as authenticated DB attachments during controlled beta;
-- they are never mounted into the public /uploads static path.

CREATE TABLE IF NOT EXISTS public.beta_feedback (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    user_id uuid NOT NULL,
    category text NOT NULL CHECK (category IN ('bug','incorrect_number','ai_answer','confusing_ux','feature_request','other')),
    severity text NOT NULL CHECK (severity IN ('p0','p1','p2')),
    status text NOT NULL DEFAULT 'open' CHECK (status IN ('open','reviewing','fixed','closed')),
    title text NOT NULL,
    description text NOT NULL,
    path text NOT NULL DEFAULT '',
    user_role text NULL,
    app_version text NULL,
    browser text NULL,
    viewport text NULL,
    request_id text NULL,
    attachment_mime text NULL,
    attachment_bytes bytea NULL,
    resolution_notes text NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_beta_feedback_company_status_created
ON public.beta_feedback(company_id,status,created_at DESC);

CREATE INDEX IF NOT EXISTS ix_beta_feedback_company_severity_created
ON public.beta_feedback(company_id,severity,created_at DESC);

ALTER TABLE public.beta_feedback ENABLE ROW LEVEL SECURITY;

COMMENT ON TABLE public.beta_feedback IS
'Controlled-beta feedback. Access is through authenticated company-scoped backend endpoints; screenshots may contain finance UI and are not publicly exposed.';
