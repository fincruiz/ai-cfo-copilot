-- P9 Stage 8.6: durable ingestion-job state for streamed/background GL imports.
CREATE TABLE IF NOT EXISTS public.ingestion_jobs (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL,
    uploaded_by uuid NULL,
    reporting_period_id uuid NULL,
    job_type text NOT NULL DEFAULT 'general_ledger',
    original_file_name text NOT NULL,
    staged_path text NOT NULL,
    file_size_bytes bigint NOT NULL DEFAULT 0,
    mime_type text NULL,
    source_system text NULL,
    status text NOT NULL DEFAULT 'queued',
    progress_percent integer NOT NULL DEFAULT 0,
    phase text NOT NULL DEFAULT 'queued',
    total_rows integer NULL,
    valid_rows integer NULL,
    invalid_rows integer NULL,
    inserted_rows integer NOT NULL DEFAULT 0,
    file_upload_id uuid NULL,
    error_message text NULL,
    attempts integer NOT NULL DEFAULT 0,
    created_at timestamptz NOT NULL DEFAULT now(),
    started_at timestamptz NULL,
    completed_at timestamptz NULL,
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_ingestion_jobs_company_created
ON public.ingestion_jobs (company_id, created_at DESC);
CREATE INDEX IF NOT EXISTS ix_ingestion_jobs_status_created
ON public.ingestion_jobs (status, created_at);
ALTER TABLE public.ingestion_jobs ENABLE ROW LEVEL SECURITY;
