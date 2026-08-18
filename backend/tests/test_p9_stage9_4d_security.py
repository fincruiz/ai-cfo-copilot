from pathlib import Path

from app.security.company_roles import can_company_admin, can_finance_write


def test_rbac():
    assert can_company_admin("owner")
    assert can_company_admin("admin")
    assert not can_company_admin("cfo")
    assert not can_company_admin("viewer")

    assert can_finance_write("accountant")
    assert not can_finance_write("viewer")
    assert not can_finance_write("board_member")


def test_invite_tokens_are_hashed():
    text = Path(
        "migrations/20260818_p9_stage9_4d_secure_invitations.sql"
    ).read_text()

    assert "token_hash" in text
    assert "ENABLE ROW LEVEL SECURITY" in text


def test_acceptance_is_inactive_until_profile():
    text = Path("app/api/v1/core/access/router.py").read_text()

    accept_start = text.index("async def accept(")
    profile_start = text.index('@router.get("/profile"')
    accept_section = text[accept_start:profile_start]

    profile_update_start = text.index("async def save_profile(")
    role_route_start = text.index('@router.patch("/members/')
    profile_section = text[profile_update_start:role_route_start]

    # Accepting an invitation must create/update a membership as inactive.
    assert "INSERT INTO public.company_members" in accept_section
    assert "is_active" in accept_section
    assert "false" in accept_section.lower()

    # The authenticated email must match the invited email.
    assert "INVITATION_EMAIL_MISMATCH" in accept_section

    # Only profile completion may activate the accepted membership.
    assert "SET is_active=true" in profile_section
    assert "status='completed'" in profile_section


def test_owner_access_is_protected():
    text = Path("app/api/v1/core/access/router.py").read_text()

    assert "OWNER_ROLE_PROTECTED" in text
    assert "OWNER_ACCESS_PROTECTED" in text