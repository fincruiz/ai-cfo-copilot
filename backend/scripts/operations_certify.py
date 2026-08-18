"""Read-only FinCruiz operational readiness certification.

Usage:
  python -m scripts.operations_certify --company-id <uuid>
"""
from __future__ import annotations

import argparse
import asyncio
from pathlib import Path
import sys
from uuid import UUID

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.database.session import AsyncSessionLocal
from app.services.operations_service import OperationsService


async def main(company_id: UUID) -> int:
    if AsyncSessionLocal is None:
        print("BLOCKED: DATABASE_URL is not configured.")
        return 2

    async with AsyncSessionLocal() as session:
        result = await OperationsService(session).readiness(company_id)

    print("\nFinCruiz Operational Readiness")
    print("=" * 96)
    print(f"Status: {result['status'].upper()} | Score: {result['score']}%")
    print("-" * 96)
    for check in result["checks"]:
        print(f"{check['status'].upper():10} {check['label'][:34]:34} {check['detail']}")
        if check.get("action"):
            print(f"{'':10} {'Action:':34} {check['action']}")
    print("-" * 96)
    print(f"DB latency: {result['database_latency_ms']:.2f} ms")
    print(
        f"Ingestion: {result['ingestion_open_jobs']} open, "
        f"{result['ingestion_stale_jobs']} stale, "
        f"{result['ingestion_recent_failures']} recent failure(s)"
    )

    return 0 if result["status"] == "healthy" else 1 if result["status"] == "degraded" else 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--company-id", type=UUID, required=True)
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.company_id)))
