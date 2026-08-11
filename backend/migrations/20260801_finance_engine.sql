    -- FinCruiz finance-engine integration.
    create table if not exists public.finance_account_mappings (
    id uuid primary key default gen_random_uuid(), company_id uuid not null references public.companies(id) on delete cascade,
    source_account_code text not null, source_account_name text, statement text not null check (statement in ('income_statement','balance_sheet')),
    reporting_group text not null, reporting_subgroup text, sign_convention text not null default 'positive', display_order integer,
    is_confirmed boolean not null default false, created_at timestamptz not null default now(), updated_at timestamptz not null default now(),
    constraint uq_finance_mapping_company_account unique(company_id,source_account_code)
    );
    create index if not exists ix_finance_mapping_company on public.finance_account_mappings(company_id);
    create index if not exists ix_gl_transactions_company_date on public.gl_transactions(company_id,transaction_date);
    create index if not exists ix_gl_transactions_upload on public.gl_transactions(file_upload_id);
    -- Run in Supabase SQL editor before enabling report/mapping endpoints.
