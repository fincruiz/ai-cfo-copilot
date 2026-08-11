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

CREATE TABLE IF NOT EXISTS public.forecast_model_runs (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
    run_name text NOT NULL,
    forecast_start date NOT NULL,
    forecast_months integer NOT NULL,
    budget_version_id uuid NULL REFERENCES public.planning_versions(id) ON DELETE SET NULL,
    configuration jsonb NOT NULL,
    summary jsonb NOT NULL DEFAULT '{}'::jsonb,
    result_payload jsonb NOT NULL DEFAULT '{}'::jsonb,
    status text NOT NULL DEFAULT 'completed',
    created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_forecast_model_runs_company ON public.forecast_model_runs(company_id, created_at DESC);

CREATE TABLE IF NOT EXISTS public.scenario_model_runs (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    forecast_run_id uuid NULL REFERENCES public.forecast_model_runs(id) ON DELETE CASCADE,
    scenario_name text NOT NULL,
    assumptions jsonb NOT NULL,
    impact_summary jsonb NOT NULL,
    result_payload jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.board_pack_templates (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    template_name text NOT NULL,
    selected_sections jsonb NOT NULL DEFAULT '[]'::jsonb,
    branding jsonb NOT NULL DEFAULT '{}'::jsonb,
    is_default boolean NOT NULL DEFAULT false,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_board_pack_template UNIQUE(company_id, template_name)
);

CREATE TABLE IF NOT EXISTS public.board_pack_runs (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    template_id uuid NULL REFERENCES public.board_pack_templates(id) ON DELETE SET NULL,
    forecast_run_id uuid NULL REFERENCES public.forecast_model_runs(id) ON DELETE SET NULL,
    pack_name text NOT NULL,
    reporting_period text NOT NULL,
    selected_sections jsonb NOT NULL,
    commentary jsonb NOT NULL DEFAULT '{}'::jsonb,
    status text NOT NULL DEFAULT 'completed',
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.generated_artifacts (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    board_pack_run_id uuid NULL REFERENCES public.board_pack_runs(id) ON DELETE CASCADE,
    artifact_type text NOT NULL CHECK (artifact_type IN ('pptx','pdf','xlsx')),
    file_name text NOT NULL,
    storage_path text NOT NULL,
    file_size_bytes bigint NOT NULL DEFAULT 0,
    created_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_generated_artifacts_company ON public.generated_artifacts(company_id, created_at DESC);
