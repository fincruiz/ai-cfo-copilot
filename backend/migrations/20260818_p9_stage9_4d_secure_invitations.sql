-- P9 Stage 9.4D — secure invitation boundary
CREATE TABLE IF NOT EXISTS public.company_invitations(
 id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
 company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
 email text NOT NULL,
 role public.company_role NOT NULL,
 token_hash text NOT NULL UNIQUE,
 status text NOT NULL DEFAULT 'pending' CHECK(status IN ('pending','accepted','completed','revoked','expired')),
 invited_by uuid NOT NULL,
 accepted_by uuid NULL,
 expires_at timestamptz NOT NULL,
 accepted_at timestamptz NULL,
 completed_at timestamptz NULL,
 revoked_at timestamptz NULL,
 created_at timestamptz NOT NULL DEFAULT now(),
 updated_at timestamptz NOT NULL DEFAULT now()
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_company_invitation_pending_email
 ON public.company_invitations(company_id,lower(email)) WHERE status='pending';
CREATE INDEX IF NOT EXISTS ix_company_invitations_email_status
 ON public.company_invitations(lower(email),status,expires_at);
ALTER TABLE public.company_invitations ENABLE ROW LEVEL SECURITY;
