"""FinCruiz Stage 10 paid-launch certification and release gate.

Read-only. It does not enable billing, mutate company data, or replay provider events.

Usage:
  python -m scripts.launch_certify --company-id <uuid>

Optional acknowledgement for a geographically non-production performance test environment:
  --accept-test-region-latency
"""
from __future__ import annotations

import argparse
import asyncio
from dataclasses import dataclass
from pathlib import Path
import sys
from uuid import UUID

from sqlalchemy import text

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.core.config import settings


@dataclass(frozen=True)
class Gate:
    key: str
    label: str
    status: str  # pass | conditional | fail
    detail: str
    live_payment_blocker: bool = False


def decision(gates: list[Gate]) -> str:
    if any(g.status == "fail" for g in gates):
        return "NO-GO"
    if any(g.status == "conditional" for g in gates):
        return "CONDITIONAL GO"
    return "GO"


def exit_code(gates: list[Gate]) -> int:
    result = decision(gates)
    return 2 if result == "NO-GO" else 1 if result == "CONDITIONAL GO" else 0


async def database_baseline(session) -> Gate:
    required = {
        "audit_events", "branches", "company_members", "company_invitations",
        "file_uploads", "gl_transactions", "ingestion_jobs", "billing_events",
        "company_subscriptions",
    }
    found = set((await session.execute(text("""
        SELECT table_name FROM information_schema.tables
        WHERE table_schema='public' AND table_name=ANY(CAST(:tables AS text[]))
    """), {"tables": sorted(required)})).scalars().all())
    missing = sorted(required - found)
    if missing:
        return Gate("database", "Database baseline", "fail", "Missing table(s): " + ", ".join(missing))

    generated = (await session.execute(text("""
        SELECT is_generated,generation_expression FROM information_schema.columns
        WHERE table_schema='public' AND table_name='gl_transactions' AND column_name='net_amount'
    """))).mappings().one_or_none()
    valid_generated = bool(generated and str(generated["is_generated"]).upper() == "ALWAYS"
                           and "debit" in str(generated["generation_expression"] or "").lower()
                           and "credit" in str(generated["generation_expression"] or "").lower())
    if not valid_generated:
        return Gate("database", "Database baseline", "fail", "GL net_amount is not database-generated from debit/credit.")
    return Gate("database", "Database baseline", "pass", "Required launch tables and generated GL net_amount are present.")


async def security_gate(session, company_id: UUID) -> Gate:
    tables = ["companies","company_members","company_invitations","branches","file_uploads","gl_transactions",
              "finance_account_mappings","planning_versions","native_plan_lines","integration_connections",
              "integration_records","audit_events","company_subscriptions","billing_events"]
    rows = (await session.execute(text("""
        SELECT c.relname,c.relrowsecurity FROM pg_class c JOIN pg_namespace n ON n.oid=c.relnamespace
        WHERE n.nspname='public' AND c.relname=ANY(CAST(:tables AS text[]))
    """), {"tables": tables})).all()
    rls = {name: bool(enabled) for name, enabled in rows}
    bad = sorted(name for name in tables if not rls.get(name, False))
    duplicates = int((await session.execute(text("""
        SELECT count(*) FROM (
          SELECT company_id,user_id FROM public.company_members
          GROUP BY company_id,user_id HAVING count(*)>1
        ) x
    """))).scalar_one() or 0)
    premature = int((await session.execute(text("""
        SELECT count(*) FROM public.company_invitations i
        JOIN public.company_members m ON m.company_id=i.company_id AND m.user_id=i.accepted_by
        WHERE i.company_id=:company_id AND i.status='accepted' AND m.is_active=true
    """), {"company_id": company_id})).scalar_one() or 0)
    if bad or duplicates or premature:
        return Gate("security", "Security & tenant isolation", "fail",
                    f"RLS missing/off={bad}; duplicate memberships={duplicates}; premature active invites={premature}.")
    return Gate("security", "Security & tenant isolation", "pass", "RLS, membership uniqueness and profile-before-access gate passed.")


def auth_gate() -> Gate:
    url = (settings.auth_frontend_url or "").strip()
    if settings.is_production and url.startswith("https://") and "localhost" not in url and "127.0.0.1" not in url:
        return Gate("auth", "Authentication configuration", "pass", "Production auth callback uses public HTTPS.")
    if settings.is_production:
        return Gate("auth", "Authentication configuration", "fail", "Production auth callback is not a public HTTPS URL.")
    return Gate("auth", "Authentication configuration", "conditional",
                f"Backend ENVIRONMENT is {settings.environment!r}; production auth configuration is not being certified.")


def frontend_static_gate(frontend_root: Path) -> Gate:
    required = [
        "app/login/page.tsx", "app/auth/callback/page.tsx", "app/dashboard/page.tsx",
        "app/dashboard/subscription/page.tsx", "app/dashboard/support/page.tsx", "app/dashboard/access/page.tsx",
    ]
    missing = [p for p in required if not (frontend_root / p).exists()]
    if missing:
        return Gate("frontend", "Frontend release surface", "fail", "Missing critical route file(s): " + ", ".join(missing))
    package = frontend_root / "package.json"
    if not package.exists():
        return Gate("frontend", "Frontend release surface", "fail", "Frontend package.json not found.")
    return Gate("frontend", "Frontend release surface", "pass",
                "Critical route sources are present. npm run build must still pass in CI/local release verification.")


async def main(company_id: UUID, accept_test_region_latency: bool, frontend_root: Path) -> int:
    from app.database.session import AsyncSessionLocal
    from app.services.billing.service import BillingService
    from app.services.finance.reporting_service import ReportingService
    from app.services.finance.reliability_service import FinanceReliabilityService
    from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
    from app.services.operations_service import OperationsService
    from app.services.paid_launch_certification_service import paid_launch_configuration_checks

    if AsyncSessionLocal is None:
        print("NO-GO: DATABASE_URL is not configured.")
        return 2

    gates: list[Gate] = [auth_gate(), frontend_static_gate(frontend_root)]
    for item in paid_launch_configuration_checks():
        gates.append(Gate(
            f"paid_{item.key}",
            item.label,
            "pass" if item.status == "ready" else "fail",
            item.detail if not item.action else f"{item.detail} Action: {item.action}",
            live_payment_blocker=item.key in {"live_payment_switch", "stripe_live_credentials", "razorpay_live_credentials"} and item.status != "ready",
        ))
    async with AsyncSessionLocal() as session:
        company_exists = bool(await session.scalar(text("SELECT 1 FROM public.companies WHERE id=:id AND is_active=true"), {"id": company_id}))
        if not company_exists:
            gates.append(Gate("company", "Certification company", "fail", "Company UUID does not identify an active company."))
        else:
            gates.append(Gate("company", "Certification company", "pass", "Active certification company found."))

        gates.append(await database_baseline(session))
        gates.append(await security_gate(session, company_id))

        finance = await FinanceReliabilityService(ReportingService(GLTransactionRepository(session))).certify(company_id)
        finance_status = "fail" if finance["status"] == "blocked" else "conditional" if finance["status"] == "attention" else "pass"
        gates.append(Gate("finance", "Finance integrity", finance_status,
                          f"Reliability {finance['score']}%; assurance {finance['assurance_grade']} ({finance['assurance_score']}%)."))

        operations = await OperationsService(session).readiness(company_id)
        operational_blockers = [c for c in operations["checks"] if c["status"] == "unhealthy" and c["key"] != "database_latency"]
        if operational_blockers:
            gates.append(Gate("operations", "Ingestion & recovery", "fail",
                              "; ".join(f"{c['label']}: {c['detail']}" for c in operational_blockers)))
        elif operations["status"] != "healthy":
            gates.append(Gate("operations", "Ingestion & recovery", "conditional",
                              f"Operational score {operations['score']}%; no non-latency launch blocker detected."))
        else:
            gates.append(Gate("operations", "Ingestion & recovery", "pass", f"Operational readiness {operations['score']}%."))

        latency = float(operations["database_latency_ms"])
        if latency < settings.database_degraded_ms:
            gates.append(Gate("performance", "Performance", "pass", f"Database health latency {latency:.2f} ms."))
        elif accept_test_region_latency:
            gates.append(Gate("performance", "Performance", "conditional",
                              f"Database health latency {latency:.2f} ms; accepted only for geographically mismatched test infrastructure. Re-certify before production launch."))
        else:
            gates.append(Gate("performance", "Performance", "conditional",
                              f"Database health latency {latency:.2f} ms exceeds {settings.database_degraded_ms} ms target. Run deployed load certification."))

        billing = await BillingService(session).readiness(company_id=company_id)
        if billing["status"] == "blocked":
            billing_status = "fail"
        elif billing["status"] == "attention":
            billing_status = "conditional"
        elif settings.is_production and billing["mode"] != "live":
            billing_status = "fail"
        else:
            billing_status = "pass"
        gates.append(Gate("billing", "Billing runtime configuration", billing_status,
                          f"Provider={billing['provider']} mode={billing['mode']}; verified historical events={billing['recent_verified_events']}.",
                          live_payment_blocker=billing_status != "pass"))

    result = decision(gates)
    blockers = [g for g in gates if g.status == "fail"]
    conditions = [g for g in gates if g.status == "conditional"]
    payment_conditions = [g for g in gates if g.live_payment_blocker]

    print("\nFinCruiz Final Launch Certification")
    print("=" * 104)
    for gate in gates:
        label = "PASS" if gate.status == "pass" else "CONDITIONAL" if gate.status == "conditional" else "FAIL"
        print(f"{label:12} {gate.label:32} {gate.detail}")
    print("-" * 104)
    print(f"LAUNCH DECISION: {result}")
    print(f"Blocking issues: {len(blockers)} | Conditions: {len(conditions)}")
    print("Team beta / controlled customer testing:", "BLOCKED" if blockers else "APPROVED")
    print("Live paid launch:", "NOT APPROVED" if blockers or payment_conditions or conditions else "APPROVED")
    if conditions:
        print("\nConditions to close:")
        for item in conditions:
            print(f"- {item.label}: {item.detail}")
    return exit_code(gates)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--company-id", type=UUID, required=True)
    parser.add_argument("--accept-test-region-latency", action="store_true",
                        help="Keep high latency conditional (never PASS) when test backend/database regions are intentionally mismatched.")
    parser.add_argument("--frontend-root", type=Path, default=ROOT.parent / "frontend")
    args = parser.parse_args()
    raise SystemExit(asyncio.run(main(args.company_id, args.accept_test_region_latency, args.frontend_root)))
