-- FinCruiz P6: integration hub + organizational brain foundation
CREATE EXTENSION IF NOT EXISTS pgcrypto;

CREATE TABLE IF NOT EXISTS public.integration_connections (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    provider text NOT NULL CHECK (provider IN ('xero','zoho','tally')),
    status text NOT NULL DEFAULT 'disconnected',
    external_tenant_id text NULL,
    external_tenant_name text NULL,
    access_token_encrypted text NULL,
    refresh_token_encrypted text NULL,
    token_expires_at timestamptz NULL,
    bridge_token_hash text NULL,
    metadata jsonb NOT NULL DEFAULT '{}'::jsonb,
    last_synced_at timestamptz NULL,
    last_sync_status text NULL,
    last_sync_message text NULL,
    connected_by uuid NULL,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE(company_id, provider)
);
CREATE INDEX IF NOT EXISTS ix_integration_connections_company ON public.integration_connections(company_id);

CREATE TABLE IF NOT EXISTS public.integration_oauth_states (
    state text PRIMARY KEY,
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    user_id uuid NOT NULL,
    provider text NOT NULL,
    expires_at timestamptz NOT NULL,
    created_at timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS public.integration_records (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    provider text NOT NULL,
    entity_type text NOT NULL,
    external_id text NOT NULL,
    occurred_at timestamptz NULL,
    name text NULL,
    amount numeric NULL,
    currency_code text NULL,
    payload jsonb NOT NULL DEFAULT '{}'::jsonb,
    source_updated_at timestamptz NULL,
    synced_at timestamptz NOT NULL DEFAULT now(),
    UNIQUE(company_id, provider, entity_type, external_id)
);
CREATE INDEX IF NOT EXISTS ix_integration_records_company_entity ON public.integration_records(company_id, entity_type);
CREATE INDEX IF NOT EXISTS ix_integration_records_occurred ON public.integration_records(company_id, occurred_at);

CREATE TABLE IF NOT EXISTS public.organizational_memory (
    id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
    company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
    memory_type text NOT NULL DEFAULT 'management_context',
    title text NOT NULL,
    content text NOT NULL,
    importance text NOT NULL DEFAULT 'normal',
    effective_from date NULL,
    effective_to date NULL,
    created_by uuid NULL,
    is_active boolean NOT NULL DEFAULT true,
    created_at timestamptz NOT NULL DEFAULT now(),
    updated_at timestamptz NOT NULL DEFAULT now()
);
CREATE INDEX IF NOT EXISTS ix_org_memory_company_active ON public.organizational_memory(company_id, is_active);
