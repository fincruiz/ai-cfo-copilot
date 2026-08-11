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

CREATE TABLE IF NOT EXISTS public.company_preferences (
    company_id uuid PRIMARY KEY REFERENCES public.companies(id) ON DELETE CASCADE,
    theme_preference text NOT NULL DEFAULT 'system',
    number_format text NOT NULL DEFAULT 'international',
    reporting_frequency text NOT NULL DEFAULT 'monthly',
    default_report_view text NOT NULL DEFAULT 'consolidated',
    show_ai_assistant boolean NOT NULL DEFAULT true,
    email_notifications boolean NOT NULL DEFAULT true,
    variance_warning_percent numeric NOT NULL DEFAULT 10,
    updated_at timestamptz NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS uq_finance_plan_line
ON public.finance_plan_lines (
    company_id,
    plan_type,
    version_name,
    period,
    reporting_group,
    COALESCE(source_account_code, ''),
    COALESCE(branch_id::text, '')
);
