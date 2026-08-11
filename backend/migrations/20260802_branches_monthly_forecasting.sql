CREATE TABLE IF NOT EXISTS public.branches (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    branch_code text NOT NULL,
    branch_name text NOT NULL,
    region text NULL,
    is_active boolean NOT NULL DEFAULT true,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_branches_company_code UNIQUE (company_id, branch_code)
);

CREATE INDEX IF NOT EXISTS ix_branches_company_id
    ON public.branches(company_id);

CREATE INDEX IF NOT EXISTS ix_gl_transactions_branch_id
    ON public.gl_transactions(branch_id);

ALTER TABLE public.gl_transactions
    DROP CONSTRAINT IF EXISTS fk_gl_transactions_branch_id;

ALTER TABLE public.gl_transactions
    ADD CONSTRAINT fk_gl_transactions_branch_id
    FOREIGN KEY (branch_id)
    REFERENCES public.branches(id)
    ON DELETE SET NULL;

CREATE TABLE IF NOT EXISTS public.forecast_runs (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
    method text NOT NULL,
    future_months integer NOT NULL,
    downside_factor numeric NOT NULL DEFAULT 0.90,
    upside_factor numeric NOT NULL DEFAULT 1.10,
    history_periods integer NOT NULL DEFAULT 0,
    status text NOT NULL DEFAULT 'completed',
    assumptions jsonb NOT NULL DEFAULT '{}'::jsonb,
    created_by uuid NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS ix_forecast_runs_company_created
    ON public.forecast_runs(company_id, created_at DESC);

CREATE TABLE IF NOT EXISTS public.forecast_lines (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    forecast_run_id uuid NOT NULL REFERENCES public.forecast_runs(id) ON DELETE CASCADE,
    period date NOT NULL,
    reporting_group text NOT NULL,
    base_amount numeric NOT NULL,
    downside_amount numeric NOT NULL,
    upside_amount numeric NOT NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    CONSTRAINT uq_forecast_line_run_period_group
        UNIQUE (forecast_run_id, period, reporting_group)
);

CREATE INDEX IF NOT EXISTS ix_forecast_lines_run
    ON public.forecast_lines(forecast_run_id);
