from pathlib import Path

from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService


def test_branch_reset_is_scoped_to_branch_records():
    assert WorkspaceLifecycleService.RESET_SCOPES["branches"] == ("branches",)
    assert "gl_transactions" not in WorkspaceLifecycleService.RESET_SCOPES["branches"]


def test_p4_repair_migration_is_present():
    migration = Path(__file__).parents[1] / "migrations" / "20260814_p4_customer_beta.sql"
    sql = migration.read_text(encoding="utf-8")
    assert "CREATE TABLE IF NOT EXISTS public.planning_versions" in sql
    assert "CREATE TABLE IF NOT EXISTS public.native_plan_lines" in sql
    assert "CREATE TABLE IF NOT EXISTS public.audit_events" in sql
