-- P9 Stage 9.3A — Commercial subscription foundation & entitlement guardrails.
-- Existing plan keys are preserved for backward compatibility:
-- founding = Starter, growth = Growth, enterprise = Scale / Enterprise.
-- Payment-provider state remains separate and will be connected in Stage 9.3B.

CREATE INDEX IF NOT EXISTS idx_company_subscriptions_status
    ON public.company_subscriptions(status);

CREATE INDEX IF NOT EXISTS idx_company_subscriptions_provider_customer
    ON public.company_subscriptions(provider, provider_customer_id)
    WHERE provider_customer_id IS NOT NULL;

CREATE INDEX IF NOT EXISTS idx_company_subscriptions_provider_subscription
    ON public.company_subscriptions(provider, provider_subscription_id)
    WHERE provider_subscription_id IS NOT NULL;

COMMENT ON TABLE public.company_subscriptions IS
'Application source of truth for FinCruiz subscription status and entitlement overrides. External billing providers update this record through verified server-side billing events.';

COMMENT ON COLUMN public.company_subscriptions.entitlements IS
'Per-company entitlement overrides only. Default plan entitlements are defined server-side.';
