"""Read-only PostgreSQL benchmark for FinCruiz reporting access paths.

This script does not modify data. It runs EXPLAIN (ANALYZE, BUFFERS) for
representative GL queries against one explicitly supplied TEST company UUID.

Example:
  python scripts/benchmark_reporting_sql.py --company-id <uuid>

Use only against a synthetic/test company. EXPLAIN ANALYZE executes each SELECT.
"""
from __future__ import annotations

import argparse
import asyncio
from pathlib import Path
import sys
from uuid import UUID

from sqlalchemy import text

BACKEND_ROOT = Path(__file__).resolve().parents[1]
if str(BACKEND_ROOT) not in sys.path:
    sys.path.insert(0, str(BACKEND_ROOT))

from app.database.session import engine


QUERIES = {
    "company_date": """
        SELECT transaction_date, sum(debit), sum(credit)
        FROM public.gl_transactions
        WHERE company_id = :company_id
          AND validation_status = 'valid'
          AND is_elimination = false
        GROUP BY transaction_date
        ORDER BY transaction_date
    """,
    "company_branch_date": """
        SELECT branch_id, date_trunc('month', transaction_date), sum(debit), sum(credit)
        FROM public.gl_transactions
        WHERE company_id = :company_id
          AND validation_status = 'valid'
          AND is_elimination = false
        GROUP BY branch_id, date_trunc('month', transaction_date)
        ORDER BY branch_id, date_trunc('month', transaction_date)
    """,
    "company_account_date": """
        SELECT source_account_code, sum(debit), sum(credit)
        FROM public.gl_transactions
        WHERE company_id = :company_id
          AND validation_status = 'valid'
          AND is_elimination = false
        GROUP BY source_account_code
        ORDER BY source_account_code
    """,
}


async def main(company_id: UUID) -> int:
    if engine is None:
        print("DATABASE_URL is not configured.")
        return 2

    async with engine.connect() as conn:
        count = (
            await conn.execute(
                text("SELECT count(*) FROM public.gl_transactions WHERE company_id=:company_id"),
                {"company_id": company_id},
            )
        ).scalar_one()

        print(f"Benchmark company: {company_id}")
        print(f"GL rows: {count:,}\n")

        for name, query in QUERIES.items():
            print("=" * 80)
            print(name)
            rows = (
                await conn.execute(
                    text("EXPLAIN (ANALYZE, BUFFERS, FORMAT TEXT) " + query),
                    {"company_id": company_id},
                )
            ).scalars().all()
            print("\n".join(str(row) for row in rows))
            print()

    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--company-id", type=UUID, required=True)
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.company_id)))
