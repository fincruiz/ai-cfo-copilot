-- Run after 20260802_planning_profile_board.sql.
-- This script is safe to run more than once.

CREATE TABLE IF NOT EXISTS public.finance_plan_lines (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    plan_type text NOT NULL CHECK (plan_type IN ('budget', 'forecast')),
    version_name text NOT NULL DEFAULT 'Default',
    period date NOT NULL,
    source_account_code text NULL,
    reporting_group text NOT NULL,
    reporting_subgroup text NULL,
    branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
    branch_source_value text NULL,
    amount numeric NOT NULL DEFAULT 0,
    notes text NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_finance_plan_company_type_period
    ON public.finance_plan_lines(company_id, plan_type, period);

SELECT
    table_schema,
    table_name
FROM information_schema.tables
WHERE table_schema = 'public'
  AND table_name = 'finance_plan_lines';
