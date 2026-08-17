-- P9 Stage 8.4 / 8.5C reporting-path indexes.
--
-- This version is intentionally compatible with manual execution in the
-- Supabase SQL Editor. CREATE INDEX (without CONCURRENTLY) may briefly block
-- writes while an index is built, so run during a quiet beta/testing period.
--
-- For a large live production table, create the same indexes CONCURRENTLY
-- through a direct PostgreSQL connection that is not inside a transaction.

CREATE INDEX IF NOT EXISTS ix_gl_company_date_valid
ON public.gl_transactions (company_id, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

CREATE INDEX IF NOT EXISTS ix_gl_company_branch_date_valid
ON public.gl_transactions (company_id, branch_id, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

CREATE INDEX IF NOT EXISTS ix_gl_company_account_date_valid
ON public.gl_transactions (company_id, source_account_code, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

ANALYZE public.gl_transactions;
