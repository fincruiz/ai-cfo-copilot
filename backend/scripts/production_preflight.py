"""FinCruiz production database preflight.

Run from backend:
  .\.venv\Scripts\python.exe -m scripts.production_preflight

This is read-only.  It validates the database assumptions that have previously
caused production failures before a deployment is promoted.
"""
from __future__ import annotations

import asyncio
from dataclasses import dataclass
from pathlib import Path
import sys
from typing import Any

from sqlalchemy import text

BACKEND_ROOT = Path(__file__).resolve().parents[1]
if str(BACKEND_ROOT) not in sys.path:
    sys.path.insert(0, str(BACKEND_ROOT))

from app.database.session import engine


REQUIRED_TABLES = {
    "audit_events",
    "branches",
    "company_members",
    "file_uploads",
    "gl_transactions",
    "ingestion_jobs",
}

REQUIRED_BRANCH_COLUMNS = {
    "id",
    "company_id",
    "branch_code",
    "branch_name",
    "region",
    "review_status",
    "is_active",
}

RLS_TABLES = {
    "audit_events",
    "branches",
    "company_members",
    "file_uploads",
    "gl_transactions",
    "ingestion_jobs",
}


@dataclass(frozen=True)
class Check:
    name: str
    status: str
    detail: str

    @property
    def failed(self) -> bool:
        return self.status == "FAIL"

    @property
    def warning(self) -> bool:
        return self.status == "WARN"


def _check(name: str, ok: bool, detail_ok: str, detail_fail: str) -> Check:
    return Check(name=name, status="PASS" if ok else "FAIL", detail=detail_ok if ok else detail_fail)


async def run_preflight() -> list[Check]:
    if engine is None:
        return [Check("database configuration", "FAIL", "DATABASE_URL is not configured.")]

    checks: list[Check] = []
    async with engine.connect() as conn:
        table_rows = (
            await conn.execute(
                text("""
                    SELECT table_name
                    FROM information_schema.tables
                    WHERE table_schema='public'
                      AND table_name = ANY(CAST(:tables AS text[]))
                """),
                {"tables": sorted(REQUIRED_TABLES)},
            )
        ).scalars().all()
        found_tables = set(table_rows)
        missing_tables = sorted(REQUIRED_TABLES - found_tables)
        checks.append(_check(
            "required tables",
            not missing_tables,
            "All required production tables exist.",
            "Missing table(s): " + ", ".join(missing_tables),
        ))

        branch_cols = set(
            (
                await conn.execute(
                    text("""
                        SELECT column_name
                        FROM information_schema.columns
                        WHERE table_schema='public' AND table_name='branches'
                    """)
                )
            ).scalars().all()
        )
        missing_branch_cols = sorted(REQUIRED_BRANCH_COLUMNS - branch_cols)
        checks.append(_check(
            "branches schema",
            not missing_branch_cols,
            "Branch schema includes region/review fields required by the application.",
            "Missing branch column(s): " + ", ".join(missing_branch_cols),
        ))

        generated = (
            await conn.execute(
                text("""
                    SELECT is_generated, generation_expression
                    FROM information_schema.columns
                    WHERE table_schema='public'
                      AND table_name='gl_transactions'
                      AND column_name='net_amount'
                """)
            )
        ).mappings().one_or_none()
        net_generated = bool(
            generated
            and str(generated["is_generated"]).upper() == "ALWAYS"
            and "debit" in str(generated["generation_expression"] or "").lower()
            and "credit" in str(generated["generation_expression"] or "").lower()
        )
        checks.append(_check(
            "GL generated net_amount",
            net_generated,
            "net_amount is GENERATED ALWAYS by PostgreSQL.",
            "net_amount is missing or is no longer GENERATED ALWAYS from debit/credit.",
        ))

        duplicate_count = int(
            (
                await conn.execute(
                    text("""
                        SELECT count(*)
                        FROM (
                            SELECT company_id, user_id
                            FROM public.company_members
                            WHERE is_active = true
                            GROUP BY company_id, user_id
                            HAVING count(*) > 1
                        ) duplicates
                    """)
                )
            ).scalar_one()
        )
        checks.append(_check(
            "active membership duplicates",
            duplicate_count == 0,
            "No duplicate active company memberships detected.",
            f"{duplicate_count} duplicate active company/user membership pair(s) detected.",
        ))

        unique_indexes = (
            await conn.execute(
                text("""
                    SELECT indexdef
                    FROM pg_indexes
                    WHERE schemaname='public' AND tablename='company_members'
                """)
            )
        ).scalars().all()
        normalized = [" ".join(str(v).lower().split()) for v in unique_indexes]
        has_membership_unique = any(
            "unique" in idx
            and "company_id" in idx
            and "user_id" in idx
            for idx in normalized
        )
        checks.append(_check(
            "membership uniqueness index",
            has_membership_unique,
            "A unique company_id/user_id membership index is present.",
            "No unique company_id/user_id index was detected.",
        ))

        rls_rows = (
            await conn.execute(
                text("""
                    SELECT c.relname, c.relrowsecurity
                    FROM pg_class c
                    JOIN pg_namespace n ON n.oid=c.relnamespace
                    WHERE n.nspname='public'
                      AND c.relname = ANY(CAST(:tables AS text[]))
                """),
                {"tables": sorted(RLS_TABLES)},
            )
        ).all()
        rls_map = {name: bool(enabled) for name, enabled in rls_rows}
        missing_rls = sorted(name for name in RLS_TABLES if not rls_map.get(name, False))
        checks.append(_check(
            "tenant-table RLS",
            not missing_rls,
            "RLS is enabled on all audited launch-critical tables.",
            "RLS is disabled/missing on: " + ", ".join(missing_rls),
        ))

    return checks


def render(checks: list[Check]) -> int:
    print("\nFinCruiz Production Preflight")
    print("=" * 78)
    for item in checks:
        print(f"{item.status:5}  {item.name:31} {item.detail}")
    failures = sum(item.failed for item in checks)
    warnings = sum(item.warning for item in checks)
    print("-" * 78)
    if failures:
        print(f"FAIL: {failures} blocker(s), {warnings} warning(s). Do not promote this build.")
        return 2
    print(f"PASS: production baseline is compatible ({warnings} warning(s)).")
    return 0


async def main() -> int:
    return render(await run_preflight())


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
