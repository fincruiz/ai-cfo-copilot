ALTER TABLE public.company_subscriptions
    ADD COLUMN IF NOT EXISTS billing_country_code text,
    ADD COLUMN IF NOT EXISTS billing_interval text NOT NULL DEFAULT 'monthly',
    ADD COLUMN IF NOT EXISTS requested_plan text,
    ADD COLUMN IF NOT EXISTS requested_interval text,
    ADD COLUMN IF NOT EXISTS change_requested_at timestamptz,
    ADD COLUMN IF NOT EXISTS cancellation_requested_at timestamptz;

UPDATE public.company_subscriptions s
SET billing_country_code = c.country_code
FROM public.companies c
WHERE c.id = s.company_id AND s.billing_country_code IS NULL;

ALTER TABLE public.company_subscriptions
    DROP CONSTRAINT IF EXISTS company_subscriptions_billing_interval_check;
ALTER TABLE public.company_subscriptions
    ADD CONSTRAINT company_subscriptions_billing_interval_check CHECK (billing_interval IN ('monthly','annual'));

ALTER TABLE public.company_subscriptions
    DROP CONSTRAINT IF EXISTS company_subscriptions_requested_interval_check;
ALTER TABLE public.company_subscriptions
    ADD CONSTRAINT company_subscriptions_requested_interval_check CHECK (requested_interval IS NULL OR requested_interval IN ('monthly','annual'));

ALTER TABLE public.company_subscriptions
    DROP CONSTRAINT IF EXISTS company_subscriptions_requested_plan_check;
ALTER TABLE public.company_subscriptions
    ADD CONSTRAINT company_subscriptions_requested_plan_check CHECK (requested_plan IS NULL OR requested_plan IN ('founding','growth','enterprise'));
