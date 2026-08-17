-- P9 Stage 9.0 — Production baseline guardrails.
-- Production has already been manually verified.  This migration makes the
-- expected branch schema reproducible for fresh/stale environments.

ALTER TABLE IF EXISTS public.branches
    ADD COLUMN IF NOT EXISTS region text NULL;

ALTER TABLE IF EXISTS public.branches
    ADD COLUMN IF NOT EXISTS review_status text NOT NULL DEFAULT 'accepted';

ALTER TABLE IF EXISTS public.branches
    ADD COLUMN IF NOT EXISTS source_value text NULL;

ALTER TABLE IF EXISTS public.branches
    ADD COLUMN IF NOT EXISTS discovered_from_upload_id uuid NULL;

ALTER TABLE IF EXISTS public.audit_events ENABLE ROW LEVEL SECURITY;
ALTER TABLE IF EXISTS public.ingestion_jobs ENABLE ROW LEVEL SECURITY;

-- Intentionally do not modify gl_transactions.net_amount here.
-- Production verification confirms that it is GENERATED ALWAYS AS (debit-credit),
-- and the application repository now strips that field before every insert.
