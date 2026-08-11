-- FinCruiz active finance dataset and report-version foundation.
-- Keeps every upload, but ensures live reports use only the latest active GL dataset
-- for each company/reporting period.

alter table public.file_uploads
    add column if not exists is_active boolean not null default false;

alter table public.file_uploads
    add column if not exists superseded_at timestamptz null;

-- Activate the latest successfully validated GL upload for each company/period.
with ranked_uploads as (
    select
        id,
        row_number() over (
            partition by company_id, document_type, reporting_period_id
            order by created_at desc, id desc
        ) as upload_rank
    from public.file_uploads
    where document_type = 'general_ledger'
      and processing_status = 'validated'
)
update public.file_uploads as uploads
set
    is_active = (ranked.upload_rank = 1),
    superseded_at = case
        when ranked.upload_rank = 1 then null
        else coalesce(uploads.superseded_at, now())
    end
from ranked_uploads as ranked
where uploads.id = ranked.id;

create index if not exists ix_file_uploads_active_finance_dataset
    on public.file_uploads(company_id, document_type, reporting_period_id, is_active);
