-- FinCruiz P9 Stage 5: beta subscription / entitlement foundation.
-- Billing provider integration is intentionally separate. This table is the
-- application entitlement source of truth that Stripe/Razorpay/etc. can update later.

CREATE TABLE IF NOT EXISTS public.company_subscriptions (
    company_id uuid PRIMARY KEY REFERENCES public.companies(id) ON DELETE CASCADE,
    plan text NOT NULL DEFAULT 'trial'
        CHECK (plan IN ('trial','founding','growth','enterprise')),
    status text NOT NULL DEFAULT 'trialing'
        CHECK (status IN ('trialing','active','past_due','cancelled','expired')),
    trial_started_at timestamptz,
    trial_ends_at timestamptz,
    current_period_ends_at timestamptz,
    entitlements jsonb NOT NULL DEFAULT '{}'::jsonb,
    provider text,
    provider_customer_id text,
    provider_subscription_id text,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);

-- Existing beta workspaces become founding customers so this migration never
-- unexpectedly locks an existing tester out of the product.
INSERT INTO public.company_subscriptions(company_id, plan, status)
SELECT id, 'founding', 'active'
FROM public.companies
ON CONFLICT (company_id) DO NOTHING;

CREATE OR REPLACE FUNCTION public.fincruiz_seed_company_subscription()
RETURNS trigger
LANGUAGE plpgsql
SECURITY DEFINER
SET search_path = public
AS $$
BEGIN
    INSERT INTO public.company_subscriptions(
        company_id, plan, status, trial_started_at, trial_ends_at
    ) VALUES (
        NEW.id, 'trial', 'trialing', now(), now() + interval '30 days'
    ) ON CONFLICT (company_id) DO NOTHING;
    RETURN NEW;
END;
$$;

DROP TRIGGER IF EXISTS trg_fincruiz_seed_company_subscription ON public.companies;
CREATE TRIGGER trg_fincruiz_seed_company_subscription
AFTER INSERT ON public.companies
FOR EACH ROW EXECUTE FUNCTION public.fincruiz_seed_company_subscription();

ALTER TABLE public.company_subscriptions ENABLE ROW LEVEL SECURITY;
-- No browser-side policy is created. The FastAPI service accesses this table
-- through the server database connection and applies company membership checks.
