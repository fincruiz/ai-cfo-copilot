-- FinCruiz P4 customer-beta migration. Safe to run repeatedly.
CREATE TABLE IF NOT EXISTS public.planning_versions (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    plan_type text NOT NULL CHECK (plan_type IN ('budget','forecast')),
    version_name text NOT NULL,
    financial_year_start date NOT NULL,
    financial_year_end date NOT NULL,
    status text NOT NULL DEFAULT 'draft' CHECK (status IN ('draft','submitted','approved','locked')),
    source_type text NOT NULL DEFAULT 'native',
    assumptions jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_by uuid NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_planning_version UNIQUE(company_id, plan_type, version_name)
);
CREATE TABLE IF NOT EXISTS public.native_plan_lines (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    version_id uuid NOT NULL REFERENCES public.planning_versions(id) ON DELETE CASCADE,
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    period date NOT NULL,
    branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
    reporting_group text NOT NULL,
    reporting_subgroup text NULL,
    source_account_code text NULL,
    amount numeric NOT NULL DEFAULT 0,
    driver_type text NOT NULL DEFAULT 'manual',
    driver_value numeric NULL,
    notes text NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_native_plan_line ON public.native_plan_lines(
    version_id, period, COALESCE(branch_id::text,''), reporting_group,
    COALESCE(reporting_subgroup,''), COALESCE(source_account_code,'')
);
CREATE INDEX IF NOT EXISTS ix_native_plan_version_period ON public.native_plan_lines(version_id, period);
CREATE TABLE IF NOT EXISTS public.audit_events (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    user_id uuid NULL,
    action text NOT NULL,
    module text NOT NULL,
    summary text NOT NULL,
    metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_audit_events_company_created ON public.audit_events(company_id, created_at DESC);
