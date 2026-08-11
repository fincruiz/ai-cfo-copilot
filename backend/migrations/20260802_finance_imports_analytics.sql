CREATE TABLE IF NOT EXISTS public.finance_ageing_documents (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    file_upload_id uuid NULL REFERENCES public.file_uploads(id) ON DELETE SET NULL,
    ageing_type text NOT NULL CHECK (ageing_type IN ('AR', 'AP')),
    party_name text NOT NULL,
    document_number text NULL,
    document_date date NULL,
    due_date date NULL,
    outstanding_amount numeric NOT NULL DEFAULT 0,
    original_amount numeric NULL,
    paid_amount numeric NULL,
    branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
    branch_source_value text NULL,
    age_bucket text NOT NULL DEFAULT 'Unknown',
    days_overdue integer NULL,
    currency_code text NULL,
    source_row_number integer NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_finance_ageing_company_type
    ON public.finance_ageing_documents(company_id, ageing_type);

CREATE INDEX IF NOT EXISTS ix_finance_ageing_party
    ON public.finance_ageing_documents(company_id, ageing_type, party_name);

CREATE INDEX IF NOT EXISTS ix_finance_ageing_due_date
    ON public.finance_ageing_documents(company_id, ageing_type, due_date);

CREATE TABLE IF NOT EXISTS public.finance_import_batches (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    import_type text NOT NULL,
    original_file_name text NOT NULL,
    source_system text NULL,
    row_count integer NOT NULL DEFAULT 0,
    valid_row_count integer NOT NULL DEFAULT 0,
    invalid_row_count integer NOT NULL DEFAULT 0,
    status text NOT NULL DEFAULT 'completed',
    validation_summary jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_by uuid NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_finance_import_batches_company
    ON public.finance_import_batches(company_id, created_at DESC);
