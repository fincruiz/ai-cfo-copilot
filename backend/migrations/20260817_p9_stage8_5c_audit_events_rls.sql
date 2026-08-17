-- P9 Stage 8.5C security baseline.
-- FinCruiz uses FastAPI as the trusted application data boundary.
-- Enabling RLS without browser-facing policies preserves PostgreSQL's
-- deny-by-default behavior for roles subject to RLS.

ALTER TABLE IF EXISTS public.audit_events ENABLE ROW LEVEL SECURITY;
