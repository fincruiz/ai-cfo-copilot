"""FinCruiz Stage 9.4C finance reliability certification.

Usage:
  python -m scripts.finance_reliability_certify --company-id <uuid>

Read-only: this command does not change mappings, uploads, branches or transactions.
"""
from __future__ import annotations

import argparse
import asyncio
from pathlib import Path
import sys
from uuid import UUID

BACKEND_ROOT = Path(__file__).resolve().parents[1]
if str(BACKEND_ROOT) not in sys.path:
    sys.path.insert(0, str(BACKEND_ROOT))

from app.database.session import AsyncSessionLocal
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService
from app.services.finance.reliability_service import FinanceReliabilityService


async def main(company_id: UUID) -> int:
    if AsyncSessionLocal is None:
        print("BLOCKED: DATABASE_URL is not configured.")
        return 2

    async with AsyncSessionLocal() as session:
        reporting = ReportingService(GLTransactionRepository(session))
        result = await FinanceReliabilityService(reporting).certify(company_id)

    print("\nFinCruiz Finance Reliability Certification")
    print("=" * 96)
    print(
        f"Status: {result['status'].upper()} | Score: {result['score']}% | "
        f"Assurance: {result['assurance_grade']} ({result['assurance_score']}%)"
    )
    print("-" * 96)
    for item in result["checks"]:
        flag = "BLOCK" if item["blocking"] and item["status"] == "fail" else item["status"].upper()
        print(f"{flag:8} {item['label'][:38]:38} {item['detail']}")
    print("-" * 96)
    print(
        f"PASS {result['pass_count']} | WARNING {result['warning_count']} | FAIL {result['fail_count']}"
    )
    print(result["caveat"])

    if result["status"] == "blocked":
        return 2
    if result["status"] == "attention":
        return 1
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--company-id", type=UUID, required=True)
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.company_id)))
