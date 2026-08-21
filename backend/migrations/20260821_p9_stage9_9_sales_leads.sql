-- FinCruiz Stage 9.9: public Book a Demo lead capture.
CREATE TABLE IF NOT EXISTS public.sales_leads (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    lead_type text NOT NULL DEFAULT 'book_demo',
    name text NOT NULL,
    work_email text NOT NULL,
    company_name text NOT NULL,
    role text,
    persona text,
    country text,
    team_size text,
    message text,
    source_path text,
    referrer_host text,
    status text NOT NULL DEFAULT 'new',
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT sales_leads_status_check CHECK (status IN ('new','contacted','qualified','closed','spam'))
);

CREATE INDEX IF NOT EXISTS ix_sales_leads_created_at ON public.sales_leads(created_at DESC);
CREATE INDEX IF NOT EXISTS ix_sales_leads_status_created ON public.sales_leads(status, created_at DESC);
CREATE INDEX IF NOT EXISTS ix_sales_leads_email ON public.sales_leads(lower(work_email));

ALTER TABLE public.sales_leads ENABLE ROW LEVEL SECURITY;

COMMENT ON TABLE public.sales_leads IS
'First-party sales enquiries submitted through FinCruiz public Book a Demo forms. No anonymous browser RLS policy is granted; writes occur through the backend only.';
