-- P9 Stage 9.4E — operational/read-path indexes.
-- Supabase SQL Editor compatible. Run during a quiet beta/testing window.

CREATE INDEX IF NOT EXISTS ix_file_upload_company_active_document
ON public.file_uploads(company_id, document_type, is_active)
WHERE is_active=true;

CREATE INDEX IF NOT EXISTS ix_ingestion_jobs_company_status_updated
ON public.ingestion_jobs(company_id, status, updated_at DESC);

CREATE INDEX IF NOT EXISTS ix_mapping_company_confirmed_account
ON public.finance_account_mappings(company_id, source_account_code)
WHERE is_confirmed=true;

ANALYZE public.file_uploads;
ANALYZE public.ingestion_jobs;
ANALYZE public.finance_account_mappings;
