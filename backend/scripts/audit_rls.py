"""Audit FinCruiz tenant-table RLS posture without changing database policies.

Run either way from the backend directory:
  .\.venv\Scripts\python.exe -m scripts.audit_rls
  .\.venv\Scripts\python.exe scripts\audit_rls.py

Interpretation:
- RLS disabled/missing -> CRITICAL review.
- RLS enabled + zero policies -> DENY-BY-DEFAULT for roles subject to RLS.
  FinCruiz currently uses the trusted FastAPI backend as its data-access boundary,
  so zero direct-client policies is intentional unless a table is explicitly exposed.
- RLS enabled + policies -> POLICY-PROTECTED; policies still require review.
"""
from __future__ import annotations

import asyncio
from pathlib import Path
import sys

from sqlalchemy import text

BACKEND_ROOT = Path(__file__).resolve().parents[1]
if str(BACKEND_ROOT) not in sys.path:
    sys.path.insert(0, str(BACKEND_ROOT))

from app.database.session import engine

TENANT_TABLES = [
    "companies",
    "company_members",
    "branches",
    "file_uploads",
    "gl_transactions",
    "finance_account_mappings",
    "finance_ageing_documents",
    "planning_versions",
    "native_plan_lines",
    "integration_connections",
    "integration_records",
    "audit_events",
    "product_usage_events",
]


def classify_rls(*, exists: bool, rls_enabled: bool, policy_count: int) -> tuple[str, bool]:
    if not exists:
        return "MISSING", True
    if not rls_enabled:
        return "RLS DISABLED", True
    if policy_count == 0:
        return "DENY-BY-DEFAULT", False
    return "POLICY-PROTECTED", False


async def main() -> int:
    if engine is None:
        print("DATABASE_URL is not configured.")
        return 2

    async with engine.connect() as conn:
        rows = (
            await conn.execute(
                text(
                    """
                    SELECT
                        c.relname AS table_name,
                        c.relrowsecurity AS rls_enabled,
                        count(p.policyname) AS policy_count
                    FROM pg_class c
                    JOIN pg_namespace n
                      ON n.oid = c.relnamespace
                    LEFT JOIN pg_policies p
                      ON p.schemaname = n.nspname
                     AND p.tablename = c.relname
                    WHERE n.nspname = 'public'
                      AND c.relname = ANY(CAST(:tables AS text[]))
                    GROUP BY c.relname, c.relrowsecurity
                    ORDER BY c.relname
                    """
                ),
                {"tables": TENANT_TABLES},
            )
        ).mappings().all()

    found = {row["table_name"]: row for row in rows}
    critical = 0

    print("table                              rls     policies  posture")
    print("-" * 82)

    for name in TENANT_TABLES:
        row = found.get(name)
        exists = row is not None
        enabled = bool(row["rls_enabled"]) if row else False
        policies = int(row["policy_count"]) if row else 0
        posture, is_critical = classify_rls(
            exists=exists,
            rls_enabled=enabled,
            policy_count=policies,
        )
        critical += int(is_critical)

        rls_label = "yes" if enabled else ("NO" if exists else "missing")
        policy_label = str(policies) if exists else "-"
        print(f"{name:34} {rls_label:7} {policy_label:8} {posture}")

    if critical:
        print(f"\nCRITICAL: {critical} audited table(s) are missing or have RLS disabled.")
        return 2

    print(
        "\nNo critical RLS gaps detected. "
        "Tables with zero policies are deny-by-default for roles subject to RLS."
    )
    print(
        "FinCruiz backend authorization remains the primary tenant boundary; "
        "do not expose service-role/database credentials to the browser."
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
