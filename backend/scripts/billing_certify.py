"""FinCruiz Stage 9.4B billing certification preflight.

Read-only. It verifies configuration and the current workspace billing event state.

Usage:
  python -m scripts.billing_certify --company-id <uuid>

This script does not create a payment, mutate a subscription, or replay webhooks.
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
from app.services.billing.service import BillingService


async def main(company_id: UUID) -> int:
    if AsyncSessionLocal is None:
        print("FAIL: DATABASE_URL is not configured.")
        return 2

    async with AsyncSessionLocal() as session:
        result = await BillingService(session).readiness(company_id=company_id)

    print("\nFinCruiz Billing Certification")
    print("=" * 78)
    print(f"Provider: {result['provider']} ({result['mode']})")

    for check in result["checks"]:
        print(
            f"{check['status'].upper():9} "
            f"{check['label']:30} "
            f"{check['detail']}"
        )

    print("-" * 78)
    print(f"Verified provider events: {result['recent_verified_events']}")
    print(
        "Last verified event: "
        f"{result['last_verified_event_at'] or 'none'}"
    )

    if result["status"] == "blocked":
        print("NO-GO: billing is blocked.")
        return 2

    if result["status"] == "attention":
        print(
            "CONDITIONAL: finish provider configuration/test lifecycle "
            "before enabling live payments."
        )
        return 1

    print(
        "READY FOR SANDBOX CERTIFICATION. "
        "Live payments remain separately gated."
    )
    return 0


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--company-id",
        type=UUID,
        required=True,
        help="Real FinCruiz company UUID from public.companies.id",
    )
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.company_id)))
