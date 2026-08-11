ALTER TABLE public.branches
    ADD COLUMN IF NOT EXISTS review_status text NOT NULL DEFAULT 'accepted',
    ADD COLUMN IF NOT EXISTS source_value text NULL,
    ADD COLUMN IF NOT EXISTS discovered_from_upload_id uuid NULL REFERENCES public.file_uploads(id) ON DELETE SET NULL;

CREATE INDEX IF NOT EXISTS ix_branches_company_review_status
    ON public.branches(company_id, review_status);
