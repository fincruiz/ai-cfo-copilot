from __future__ import annotations

from typing import Any
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.repositories.finance.account_mapping_repository import AccountMappingRepository
from app.services.finance.assurance_service import FinancialAssuranceService
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService
from app.services.intelligence.brain_service import BrainService


def determine_commercial_onboarding_stage(*, transaction_count: int, pending_branches: int, unmapped_accounts: int) -> tuple[str, str, str]:
    """Return stage, next path and customer-facing next action."""
    if transaction_count <= 0:
        return "data_needed", "/dashboard/uploads?welcome=1", "Load your finance data"
    if pending_branches > 0:
        return "branch_review_required", "/dashboard/branches?welcome=1", "Review detected branches"
    if unmapped_accounts > 0:
        return "mapping_required", "/dashboard/mapping?welcome=1", "Review account mappings"
    return "ready", "/dashboard", "Open your management dashboard"


class CommercialOnboardingService:
    """Build a deterministic post-ingestion onboarding summary.

    No LLM is required to determine readiness.  The first briefing is produced
    only after the active GL is structurally ready and is built from the same
    deterministic finance/intelligence services used by the main product.
    """

    def __init__(self, session: AsyncSession) -> None:
        self.session = session

    async def _active_gl_profile(self, company_id: UUID) -> dict[str, Any]:
        row = (
            await self.session.execute(
                text(
                    """
                    SELECT
                        count(*)::int AS transaction_count,
                        count(DISTINCT g.source_account_code)::int AS account_count,
                        min(g.transaction_date) AS period_start,
                        max(g.transaction_date) AS period_end,
                        count(DISTINCT date_trunc('month', g.transaction_date))::int AS months_history
                    FROM public.gl_transactions g
                    JOIN public.file_uploads f ON f.id = g.file_upload_id
                    WHERE g.company_id=:company_id
                      AND g.validation_status='valid'
                      AND g.is_elimination=false
                      AND f.is_active=true
                      AND f.processing_status='validated'
                      AND f.document_type='general_ledger'
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()
        return dict(row)

    async def _branch_profile(self, company_id: UUID) -> dict[str, int]:
        row = (
            await self.session.execute(
                text(
                    """
                    SELECT
                        count(*) FILTER (WHERE is_active=true AND review_status='accepted')::int AS accepted,
                        count(*) FILTER (WHERE is_active=true AND review_status='pending')::int AS pending
                    FROM public.branches
                    WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()
        return {"accepted": int(row["accepted"] or 0), "pending": int(row["pending"] or 0)}

    async def _mapping_profile(self, company_id: UUID) -> dict[str, int]:
        # Only the currently active validated GL belongs in onboarding readiness.
        # Old/deactivated uploads must not keep obsolete accounts in the mapping queue.
        row = (
            await self.session.execute(
                text(
                    """
                    WITH active_accounts AS (
                        SELECT DISTINCT g.source_account_code
                        FROM public.gl_transactions g
                        JOIN public.file_uploads f ON f.id=g.file_upload_id
                        WHERE g.company_id=:company_id
                          AND g.validation_status='valid'
                          AND g.is_elimination=false
                          AND f.is_active=true
                          AND f.processing_status='validated'
                          AND f.document_type='general_ledger'
                    )
                    SELECT
                        count(*)::int AS active_accounts,
                        count(m.id)::int AS mapped_accounts,
                        count(*) FILTER (WHERE m.id IS NULL)::int AS unmapped_accounts
                    FROM active_accounts a
                    LEFT JOIN public.finance_account_mappings m
                      ON m.company_id=:company_id
                     AND m.source_account_code=a.source_account_code
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()
        return {key: int(row[key] or 0) for key in ("active_accounts", "mapped_accounts", "unmapped_accounts")}

    async def _latest_ingestion(self, company_id: UUID) -> dict[str, Any] | None:
        try:
            row = (
                await self.session.execute(
                    text(
                        """
                        SELECT id,status,phase,progress_percent,original_file_name,completed_at
                        FROM public.ingestion_jobs
                        WHERE company_id=:company_id
                        ORDER BY created_at DESC
                        LIMIT 1
                        """
                    ),
                    {"company_id": company_id},
                )
            ).mappings().one_or_none()
            return dict(row) if row else None
        except Exception:
            await self.session.rollback()
            return None

    async def build(self, company_id: UUID) -> dict[str, Any]:
        gl = await self._active_gl_profile(company_id)
        branches = await self._branch_profile(company_id)
        mappings = await self._mapping_profile(company_id)
        latest_ingestion = await self._latest_ingestion(company_id)

        transaction_count = int(gl.get("transaction_count") or 0)
        pending_branches = int(branches["pending"])
        unmapped_accounts = int(mappings["unmapped_accounts"])
        stage, next_path, next_label = determine_commercial_onboarding_stage(
            transaction_count=transaction_count,
            pending_branches=pending_branches,
            unmapped_accounts=unmapped_accounts,
        )

        data_ready = transaction_count > 0
        structure_ready = data_ready and pending_branches == 0
        mapping_ready = structure_ready and mappings["active_accounts"] > 0 and unmapped_accounts == 0

        assurance: dict[str, Any] | None = None
        briefing: dict[str, Any] | None = None
        briefing_error: str | None = None
        if mapping_ready:
            reporting = ReportingService(GLTransactionRepository(self.session))
            try:
                assurance = await FinancialAssuranceService(reporting).assess(company_id)
                overview = await BrainService(self.session).overview(company_id)
                briefing = {
                    "executive_summary": overview.get("executive_summary"),
                    "priorities": (overview.get("priorities") or [])[:3],
                    "financial_snapshot": overview.get("financial_snapshot") or [],
                    "monthly_trends": overview.get("monthly_trends") or [],
                    "suggested_questions": overview.get("suggested_questions") or [],
                }
            except Exception:
                # A briefing problem should not destroy the customer's setup state.
                # The support/error layer can diagnose the underlying endpoint later.
                await self.session.rollback()
                briefing_error = "Your finance structure is ready, but the first briefing could not be generated yet. Refresh once or open the dashboard."

        assurance_complete = assurance is not None
        intelligence_ready = mapping_ready and briefing is not None
        step_flags = [data_ready, structure_ready, mapping_ready, assurance_complete, intelligence_ready]
        completed_steps = sum(1 for value in step_flags if value)
        progress = int(round(completed_steps / len(step_flags) * 100))

        if mapping_ready and not intelligence_ready:
            stage = "briefing_pending"
            next_path = "/dashboard/getting-started"
            next_label = "Retry first briefing"
        elif intelligence_ready:
            stage = "ready"
            next_path = "/dashboard"
            next_label = "Open your management dashboard"

        latest = None
        if latest_ingestion:
            latest = {
                **latest_ingestion,
                "id": str(latest_ingestion.get("id")),
                "completed_at": latest_ingestion.get("completed_at").isoformat() if latest_ingestion.get("completed_at") else None,
            }

        return {
            "stage": stage,
            "ready_for_intelligence": intelligence_ready,
            "progress_percent": progress,
            "completed_steps": completed_steps,
            "total_steps": len(step_flags),
            "steps": [
                {"key": "data", "label": "Finance data", "complete": data_ready},
                {"key": "structure", "label": "Business structure", "complete": structure_ready},
                {"key": "mapping", "label": "Account mapping", "complete": mapping_ready},
                {"key": "assurance", "label": "Financial checks", "complete": assurance_complete},
                {"key": "intelligence", "label": "First intelligence", "complete": intelligence_ready},
            ],
            "transaction_count": transaction_count,
            "account_count": int(gl.get("account_count") or 0),
            "mapping_count": int(mappings["mapped_accounts"]),
            "unmapped_account_count": unmapped_accounts,
            "branch_count": int(branches["accepted"]),
            "pending_branch_count": pending_branches,
            "months_history": int(gl.get("months_history") or 0),
            "period_start": gl.get("period_start").isoformat() if gl.get("period_start") else None,
            "period_end": gl.get("period_end").isoformat() if gl.get("period_end") else None,
            "financial_confidence_score": assurance.get("score") if assurance else None,
            "financial_confidence_grade": assurance.get("grade") if assurance else None,
            "financial_checks": assurance.get("checks", []) if assurance else [],
            "latest_ingestion": latest,
            "briefing": briefing,
            "briefing_error": briefing_error,
            "next_path": next_path,
            "next_label": next_label,
        }
