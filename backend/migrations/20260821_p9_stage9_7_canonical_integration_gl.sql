-- FinCruiz Stage 9.7: canonical integration -> GL hardening
-- No data rewrite. Adds tenant-table RLS barriers and read-path indexes used by finance activation.

ALTER TABLE IF EXISTS public.integration_connections ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS public.integration_oauth_states ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS public.integration_records ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS public.organizational_memory ENABLE ROW LEVEL SECURITY;

CREATE INDEX IF NOT EXISTS ix_integration_records_company_provider_entity
    ON public.integration_records(company_id, provider, entity_type, occurred_at);

CREATE INDEX IF NOT EXISTS ix_integration_connections_company_provider_status
    ON public.integration_connections(company_id, provider, status);

CREATE INDEX IF NOT EXISTS ix_file_uploads_integration_source
    ON public.file_uploads(company_id, source_system, created_at DESC)
    WHERE storage_bucket = 'integration-sync';
