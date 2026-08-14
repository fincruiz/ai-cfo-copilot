from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService

def test_scoped_reset_surface_is_explicit_and_safe():
    assert set(WorkspaceLifecycleService.RESET_SCOPES) == {
        "general_ledger", "account_mappings", "coa", "ar_ageing", "ap_ageing", "planning", "forecasts", "board_packs", "branches"
    }
    assert "companies" not in {t for tables in WorkspaceLifecycleService.RESET_SCOPES.values() for t in tables}
    assert "profiles" not in {t for tables in WorkspaceLifecycleService.RESET_SCOPES.values() for t in tables}
