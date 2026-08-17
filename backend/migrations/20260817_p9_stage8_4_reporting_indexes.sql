-- P9 Stage 8.4: reporting-path indexes for large GL datasets.
-- CREATE INDEX CONCURRENTLY avoids long table locks; run this migration outside a transaction.
CREATE INDEX CONCURRENTLY IF NOT EXISTS ix_gl_company_date_valid
ON public.gl_transactions (company_id, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

CREATE INDEX CONCURRENTLY IF NOT EXISTS ix_gl_company_branch_date_valid
ON public.gl_transactions (company_id, branch_id, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

CREATE INDEX CONCURRENTLY IF NOT EXISTS ix_gl_company_account_date_valid
ON public.gl_transactions (company_id, source_account_code, transaction_date)
WHERE validation_status = 'valid' AND is_elimination = false;

ANALYZE public.gl_transactions;
