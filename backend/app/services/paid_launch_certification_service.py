from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from typing import Any

from app.core.config import Settings, settings
from app.services.billing.certification import razorpay_mode, stripe_mode


@dataclass(frozen=True)
class PaidLaunchCheck:
    key: str
    label: str
    status: str  # ready | blocked
    detail: str
    action: str | None = None
    evidence_kind: str = "configuration"


def _present(value: str | None) -> bool:
    return bool((value or "").strip())


def _fresh(value: datetime | None, *, max_age_days: int, now: datetime) -> bool:
    if value is None:
        return False
    stamp = value if value.tzinfo else value.replace(tzinfo=timezone.utc)
    return 0 <= (now - stamp.astimezone(timezone.utc)).total_seconds() <= max_age_days * 86400


def configured_paid_launch_providers(config: Settings = settings) -> list[str]:
    providers = []
    for raw in (config.paid_launch_providers or "").split(","):
        provider = raw.strip().lower()
        if provider and provider not in providers:
            providers.append(provider)
    return providers


def paid_launch_configuration_checks(
    config: Settings = settings,
    *,
    now: datetime | None = None,
) -> list[PaidLaunchCheck]:
    now = now or datetime.now(timezone.utc)
    checks: list[PaidLaunchCheck] = []

    production = config.environment.lower() == "production"
    checks.append(PaidLaunchCheck(
        "environment",
        "Production environment",
        "ready" if production else "blocked",
        "ENVIRONMENT is production." if production else f"ENVIRONMENT is {config.environment!r}; paid launch certification must run against production configuration.",
        None if production else "Set ENVIRONMENT=production in the deployed production service, then rerun certification.",
    ))

    providers = configured_paid_launch_providers(config)
    supported = {"stripe", "razorpay"}
    invalid = [item for item in providers if item not in supported]
    checks.append(PaidLaunchCheck(
        "provider_scope",
        "Paid provider scope",
        "ready" if providers and not invalid else "blocked",
        f"Paid launch providers: {', '.join(providers)}." if providers and not invalid else "Configure PAID_LAUNCH_PROVIDERS using stripe and/or razorpay.",
        None if providers and not invalid else "Declare only the payment providers that will be live at launch.",
    ))

    modes = {"stripe": stripe_mode(config.stripe_secret_key), "razorpay": razorpay_mode(config.razorpay_key_id)}
    required_values = {
        "stripe": {
            "secret_key": config.stripe_secret_key,
            "webhook_secret": config.stripe_webhook_secret,
            "starter_monthly_price": config.stripe_starter_monthly_price_id,
            "starter_annual_price": config.stripe_starter_annual_price_id,
            "growth_monthly_price": config.stripe_growth_monthly_price_id,
            "growth_annual_price": config.stripe_growth_annual_price_id,
        },
        "razorpay": {
            "key_id": config.razorpay_key_id,
            "key_secret": config.razorpay_key_secret,
            "webhook_secret": config.razorpay_webhook_secret,
            "starter_monthly_plan": config.razorpay_starter_monthly_plan_id,
            "starter_annual_plan": config.razorpay_starter_annual_plan_id,
            "growth_monthly_plan": config.razorpay_growth_monthly_plan_id,
            "growth_annual_plan": config.razorpay_growth_annual_plan_id,
        },
    }
    for provider in providers:
        missing_values = [key for key, value in required_values.get(provider, {}).items() if not _present(value)]
        checks.append(PaidLaunchCheck(
            f"{provider}_configuration",
            f"{provider.title()} production configuration",
            "ready" if not missing_values else "blocked",
            "Required provider secrets and plan/price identifiers are configured." if not missing_values else "Missing: " + ", ".join(missing_values),
            None if not missing_values else f"Complete the production {provider.title()} server-side configuration before enabling live billing.",
        ))
        mode = modes.get(provider, "unknown")
        checks.append(PaidLaunchCheck(
            f"{provider}_live_credentials",
            f"{provider.title()} live credentials",
            "ready" if mode == "live" else "blocked",
            f"{provider.title()} credential mode is {mode}." if mode != "live" else f"{provider.title()} live credential mode detected.",
            None if mode == "live" else f"Configure the production {provider.title()} credentials only after sandbox lifecycle certification is complete.",
        ))

        certified_at = config.stripe_sandbox_certified_at if provider == "stripe" else config.razorpay_sandbox_certified_at
        fresh = _fresh(certified_at, max_age_days=config.launch_certification_max_age_days, now=now)
        checks.append(PaidLaunchCheck(
            f"{provider}_sandbox_lifecycle",
            f"{provider.title()} sandbox lifecycle",
            "ready" if fresh else "blocked",
            f"Sandbox lifecycle certified at {certified_at.isoformat()}." if fresh and certified_at else "No fresh sandbox lifecycle certification timestamp is configured.",
            None if fresh else f"Run the full {provider.title()} sandbox checkout, activation/payment, failure and cancellation lifecycle and set the certification timestamp after it passes.",
            "operator_evidence",
        ))

    billing_url = (config.billing_frontend_url or "").strip()
    billing_url_ready = production and billing_url.startswith("https://") and "localhost" not in billing_url and "127.0.0.1" not in billing_url
    checks.append(PaidLaunchCheck(
        "billing_frontend_url",
        "Billing return URL",
        "ready" if billing_url_ready else "blocked",
        f"Production HTTPS billing return URL: {billing_url}." if billing_url_ready else f"Billing return URL is not production HTTPS: {billing_url or 'unset'}.",
        None if billing_url_ready else "Set BILLING_FRONTEND_URL to the deployed public HTTPS frontend.",
    ))

    live_switch = bool(config.billing_allow_live_payments)
    live_modes_ready = bool(providers) and all(modes.get(provider) == "live" for provider in providers)
    live_gate_ready = production and live_modes_ready and live_switch
    checks.append(PaidLaunchCheck(
        "live_payment_switch",
        "Live-payment safety switch",
        "ready" if live_gate_ready else "blocked",
        "BILLING_ALLOW_LIVE_PAYMENTS is deliberately enabled with production live provider credentials." if live_gate_ready else "Live billing is not deliberately enabled under a fully certified production provider configuration.",
        None if live_gate_ready else "Leave the switch false during certification. Enable it deliberately only after every other paid-launch gate is green, then rerun this check.",
    ))

    deployment_region = (config.deployment_region or "").strip().lower()
    database_region = (config.database_region or "").strip().lower()
    regions_ready = bool(deployment_region and database_region and deployment_region == database_region)
    checks.append(PaidLaunchCheck(
        "region_alignment",
        "API/database region alignment",
        "ready" if regions_ready else "blocked",
        f"Deployment and database region are both {deployment_region}." if regions_ready else f"Deployment region={deployment_region or 'unset'}; database region={database_region or 'unset'}.",
        None if regions_ready else "Deploy API and database in the intended production region (or document an approved architecture) and set both region values before performance certification.",
        "operator_evidence",
    ))

    performance_fresh = _fresh(config.production_performance_certified_at, max_age_days=config.launch_certification_max_age_days, now=now)
    checks.append(PaidLaunchCheck(
        "production_performance",
        "Deployed performance certification",
        "ready" if performance_fresh else "blocked",
        f"Production performance certified at {config.production_performance_certified_at.isoformat()}." if performance_fresh and config.production_performance_certified_at else "No fresh deployed production performance certification is recorded.",
        None if performance_fresh else "Run scripts/load_test_finance.py against the deployed production-region API using a synthetic company, then record the passing timestamp.",
        "operator_evidence",
    ))

    persistent = bool(config.import_staging_persistent and _present(config.import_staging_dir))
    checks.append(PaidLaunchCheck(
        "persistent_ingestion",
        "Persistent ingestion storage",
        "ready" if persistent else "blocked",
        f"Persistent ingestion staging is declared at {config.import_staging_dir}." if persistent else "Persistent ingestion staging has not been certified.",
        None if persistent else "Back IMPORT_STAGING_DIR with persistent storage and set IMPORT_STAGING_PERSISTENT=true only after verifying restart/deploy survival.",
    ))

    backup_fresh = _fresh(config.backup_restore_verified_at, max_age_days=config.backup_restore_max_age_days, now=now)
    checks.append(PaidLaunchCheck(
        "backup_restore",
        "Backup and restore drill",
        "ready" if backup_fresh else "blocked",
        f"Backup/restore drill verified at {config.backup_restore_verified_at.isoformat()}." if backup_fresh and config.backup_restore_verified_at else "No fresh backup/restore verification is recorded.",
        None if backup_fresh else "Complete a restore drill into an isolated target, validate finance/report integrity, and record the verification timestamp.",
        "operator_evidence",
    ))

    monitoring_fresh = _fresh(config.error_monitoring_verified_at, max_age_days=config.launch_certification_max_age_days, now=now)
    monitoring_ready = _present(config.error_monitoring_dsn) and monitoring_fresh
    checks.append(PaidLaunchCheck(
        "error_monitoring",
        "Error monitoring",
        "ready" if monitoring_ready else "blocked",
        f"Error monitoring is configured and delivery was verified at {config.error_monitoring_verified_at.isoformat()}." if monitoring_ready and config.error_monitoring_verified_at else "Error monitoring is not both configured and freshly delivery-tested.",
        None if monitoring_ready else "Configure production error monitoring, trigger a safe synthetic exception, confirm receipt/alerting, then record ERROR_MONITORING_VERIFIED_AT.",
        "operator_evidence",
    ))

    support_ready = _present(config.support_contact_email) and _present(config.support_runbook_url)
    checks.append(PaidLaunchCheck(
        "support_process",
        "Customer support process",
        "ready" if support_ready else "blocked",
        "Support contact and runbook reference are configured." if support_ready else "Support contact and/or incident runbook reference is missing.",
        None if support_ready else "Configure a monitored support contact and a production incident/support runbook before paid launch.",
    ))

    return checks


def paid_launch_summary(config: Settings = settings, *, now: datetime | None = None) -> dict[str, Any]:
    checks = paid_launch_configuration_checks(config, now=now)
    ready = sum(1 for item in checks if item.status == "ready")
    total = len(checks)
    return {
        "status": "ready" if ready == total and total else "blocked",
        "live_paid_launch_approved": bool(total and ready == total),
        "score": round(ready / max(total, 1) * 100),
        "checks": [asdict(item) for item in checks],
        "checked_at": (now or datetime.now(timezone.utc)),
    }
