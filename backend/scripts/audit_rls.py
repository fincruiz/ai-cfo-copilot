"""Audit tenant-table RLS coverage without changing any policies.

Run from backend with DATABASE_URL configured:
  python scripts/audit_rls.py

Exit code 2 means one or more tenant tables require review. This is a launch audit,
not an automatic policy migration.
"""
from __future__ import annotations
import asyncio
from sqlalchemy import text
from app.database.session import engine

TENANT_TABLES = [
    "companies", "company_members", "branches", "file_uploads", "gl_transactions",
    "finance_account_mappings", "finance_ageing_documents", "planning_versions",
    "native_plan_lines", "integration_connections", "integration_records",
    "audit_events", "product_usage_events",
]

async def main() -> int:
    if engine is None:
        print("DATABASE_URL is not configured.")
        return 2
    async with engine.connect() as conn:
        rows=(await conn.execute(text("""
            SELECT c.relname AS table_name, c.relrowsecurity AS rls_enabled,
                   count(p.policyname) AS policy_count
            FROM pg_class c
            JOIN pg_namespace n ON n.oid=c.relnamespace
            LEFT JOIN pg_policies p ON p.schemaname=n.nspname AND p.tablename=c.relname
            WHERE n.nspname='public' AND c.relname = ANY(CAST(:tables AS text[]))
            GROUP BY c.relname,c.relrowsecurity ORDER BY c.relname
        """), {"tables": TENANT_TABLES})).mappings().all()
    found={r["table_name"]:r for r in rows}; issues=0
    print("table                              rls     policies")
    print("-"*56)
    for name in TENANT_TABLES:
        row=found.get(name)
        if not row:
            print(f"{name:34} missing  -"); continue
        ok=bool(row["rls_enabled"] and int(row["policy_count"])>0)
        if not ok: issues+=1
        print(f"{name:34} {'yes' if row['rls_enabled'] else 'NO':7} {row['policy_count']}")
    if issues:
        print(f"\nReview required for {issues} table(s).")
    else:
        print("\nRLS is enabled and at least one policy exists on every audited table.")
    return 2 if issues else 0

if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
