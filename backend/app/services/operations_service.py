from __future__ import annotations

import os
from datetime import datetime, timezone
from pathlib import Path
from time import perf_counter
from typing import Any
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings


def grade_database_latency(milliseconds: float) -> str:
    if milliseconds < settings.database_degraded_ms:
        return "healthy"
    if milliseconds < settings.database_unhealthy_ms:
        return "degraded"
    return "unhealthy"


def overall_status(checks: list[dict[str, Any]]) -> str:
    statuses = {str(check.get("status")) for check in checks}
    if "unhealthy" in statuses:
        return "unhealthy"
    if "degraded" in statuses:
        return "degraded"
    return "healthy"


class OperationsService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def readiness(self, company_id: UUID) -> dict[str, Any]:
        checks: list[dict[str, Any]] = []

        started = perf_counter()
        try:
            await self.session.execute(text("SELECT 1"))
            database_latency_ms = round((perf_counter() - started) * 1000, 2)
            db_status = grade_database_latency(database_latency_ms)
            checks.append({
                "key": "database_latency",
                "label": "Database response",
                "status": db_status,
                "detail": f"Health query completed in {database_latency_ms:.2f} ms.",
                "action": "Review database connection path and slow-query workload." if db_status != "healthy" else None,
            })
        except Exception:
            database_latency_ms = round((perf_counter() - started) * 1000, 2)
            checks.append({
                "key": "database_latency",
                "label": "Database response",
                "status": "unhealthy",
                "detail": "Database health query failed.",
                "action": "Check database availability and connection configuration.",
            })

        jobs = (
            await self.session.execute(
                text("""
                    SELECT
                      COUNT(*) FILTER (WHERE status IN ('queued','processing','retry'))::int AS open_jobs,
                      COUNT(*) FILTER (
                        WHERE status='processing'
                          AND updated_at < now() - (:stale_minutes * interval '1 minute')
                      )::int AS stale_jobs,
                      COUNT(*) FILTER (
                        WHERE status IN ('failed','validation_failed')
                          AND updated_at > now() - (:failure_hours * interval '1 hour')
                      )::int AS recent_failures,
                      MAX(updated_at) AS latest_update_at
                    FROM public.ingestion_jobs
                    WHERE company_id=:company_id
                """),
                {
                    "company_id": company_id,
                    "stale_minutes": settings.ingestion_stale_minutes,
                    "failure_hours": settings.operational_failure_window_hours,
                },
            )
        ).mappings().one()

        open_jobs = int(jobs["open_jobs"] or 0)
        stale_jobs = int(jobs["stale_jobs"] or 0)
        recent_failures = int(jobs["recent_failures"] or 0)

        if stale_jobs:
            ingestion_status = "unhealthy"
            ingestion_action = (
                "Do not retry a still-processing job blindly. Confirm the worker stopped, "
                "then re-upload or mark the stale job failed through controlled support recovery."
            )
        elif recent_failures:
            ingestion_status = "degraded"
            ingestion_action = "Review failed/validation-failed imports and confirm the intended good dataset remains active."
        else:
            ingestion_status = "healthy"
            ingestion_action = None

        checks.append({
            "key": "ingestion_queue",
            "label": "Background ingestion",
            "status": ingestion_status,
            "detail": (
                f"{open_jobs} open job(s), {stale_jobs} stale processing job(s), "
                f"{recent_failures} failed/validation-failed job(s) in the last "
                f"{settings.operational_failure_window_hours} hour(s)."
            ),
            "action": ingestion_action,
        })

        active_gl = int(
            (
                await self.session.execute(
                    text("""
                        SELECT COUNT(*)::int
                        FROM public.file_uploads
                        WHERE company_id=:company_id
                          AND document_type='general_ledger'
                          AND is_active=true
                    """),
                    {"company_id": company_id},
                )
            ).scalar_one()
            or 0
        )

        dataset_status = "healthy" if active_gl == 1 else "degraded" if active_gl == 0 else "unhealthy"
        checks.append({
            "key": "active_dataset",
            "label": "Active finance dataset",
            "status": dataset_status,
            "detail": f"{active_gl} active General Ledger dataset(s).",
            "action": (
                "Load/activate one validated General Ledger."
                if active_gl == 0
                else "Investigate multiple active General Ledger versions."
                if active_gl > 1
                else None
            ),
        })

        root = Path(settings.import_staging_dir)
        absolute_or_parent = root if root.exists() else root.parent
        writable = os.access(absolute_or_parent.resolve(), os.W_OK)
        staging_status = "healthy" if writable else "unhealthy"
        staging_detail = f"Staging path is writable: {root}." if writable else f"Staging path is not writable: {root}."
        staging_action = None if writable else "Configure a writable staging directory before accepting large uploads."

        if settings.is_production and not settings.import_staging_persistent:
            staging_status = "degraded" if writable else "unhealthy"
            staging_detail += " Production persistent-storage confirmation is not enabled."
            staging_action = (
                "Set IMPORT_STAGING_PERSISTENT=true only after the staging path is backed by persistent storage; "
                "ephemeral storage can lose queued imports during restart/deploy."
            )

        checks.append({
            "key": "staging_storage",
            "label": "Import staging storage",
            "status": staging_status,
            "detail": staging_detail,
            "action": staging_action,
        })

        expected_indexes = {
            "ix_gl_company_date_valid",
            "ix_gl_company_branch_date_valid",
            "ix_gl_company_account_date_valid",
            "ix_file_upload_company_active_document",
            "ix_ingestion_jobs_company_status_updated",
            "ix_mapping_company_confirmed_account",
        }
        index_names = set(
            (
                await self.session.execute(
                    text("""
                        SELECT indexname
                        FROM pg_indexes
                        WHERE schemaname='public'
                          AND indexname = ANY(CAST(:names AS text[]))
                    """),
                    {"names": sorted(expected_indexes)},
                )
            ).scalars().all()
        )
        missing = sorted(expected_indexes - index_names)
        checks.append({
            "key": "reporting_indexes",
            "label": "Launch query indexes",
            "status": "healthy" if not missing else "degraded",
            "detail": "All launch indexes are present." if not missing else "Missing index(es): " + ", ".join(missing),
            "action": "Run the Stage 9.4E operational indexes migration during a quiet testing window." if missing else None,
        })

        status = overall_status(checks)
        weights = {"healthy": 1.0, "degraded": 0.5, "unhealthy": 0.0}
        score = round(sum(weights.get(c["status"], 0) for c in checks) / max(len(checks), 1) * 100)

        return {
            "status": status,
            "score": score,
            "checks": checks,
            "database_latency_ms": database_latency_ms,
            "ingestion_open_jobs": open_jobs,
            "ingestion_stale_jobs": stale_jobs,
            "ingestion_recent_failures": recent_failures,
            "active_gl_datasets": active_gl,
            "latest_ingestion_update_at": jobs["latest_update_at"],
            "checked_at": datetime.now(timezone.utc),
        }
