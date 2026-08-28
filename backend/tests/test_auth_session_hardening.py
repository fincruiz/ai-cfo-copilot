from pathlib import Path

from app.schemas.auth import LogoutRequest


def _frontend(relative: str) -> str:
    return Path("../frontend", relative).read_text(encoding="utf-8")


def test_logout_scope_is_fail_closed_to_supported_values():
    assert LogoutRequest().scope == "global"
    assert LogoutRequest(scope="local").scope == "local"


def test_confirmation_email_redirect_returns_to_explicit_fincruiz_callback():
    source = Path("app/services/auth_service.py").read_text(encoding="utf-8")
    assert "/auth/callback?confirmation=1&next=/onboarding" in source


def test_frontend_enforces_idle_and_absolute_session_limits():
    source = _frontend("lib/session-security.ts")
    assert "NEXT_PUBLIC_SESSION_IDLE_MINUTES" in source
    assert "NEXT_PUBLIC_SESSION_MAX_HOURS" in source
    assert "30" in source
    assert "12" in source


def test_logout_propagates_across_tabs_and_calls_server_logout():
    service = _frontend("services/auth-service.ts")
    guard = _frontend("components/session-security-guard.tsx")
    assert '"/auth/logout"' in service
    assert 'window.addEventListener("storage", onStorage)' in guard
    assert "logoutEverywhere" in guard


def test_email_callback_has_visible_success_state_before_redirect():
    callback = _frontend("app/auth/callback/page.tsx")
    assert "Email verified successfully" in callback
    assert "Continue to FinCruiz" in callback
    assert "2200" in callback
