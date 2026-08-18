from __future__ import annotations

from datetime import date, datetime, timezone
from decimal import Decimal
from typing import Any
from uuid import UUID

from sqlalchemy import text

from app.services.finance.reporting_service import ReportingService


class FinanceReliabilityService:
    """Launch-grade structural reliability certification for the active finance dataset.

    These checks prove consistency of the data FinCruiz has loaded and transformed.
    They do not assert that the source accounting records themselves are economically,
    tax, audit or statutory correct.
    """

    def __init__(self, reporting: ReportingService) -> None:
        self.reporting = reporting

    @staticmethod
    def _month_index(value: date) -> int:
        return value.year * 12 + value.month

    @staticmethod
    def _number(value: Any) -> Decimal:
        try:
            return Decimal(str(value or 0))
        except Exception:
            return Decimal("0")

    async def certify(self, company_id: UUID) -> dict[str, Any]:
        session = self.reporting.repository.session
        health = await self.reporting.data_health(company_id)
        monthly = await self.reporting.monthly_actuals(company_id)
        assurance = await __import__(
            "app.services.finance.assurance_service",
            fromlist=["FinancialAssuranceService"],
        ).FinancialAssuranceService(self.reporting).assess(company_id)

        upload = (
            await session.execute(
                text(
                    """
                    SELECT
                      COUNT(*) FILTER (WHERE is_active=true)::int AS active_count,
                      COUNT(*)::int AS total_count,
                      MAX(created_at) FILTER (WHERE is_active=true) AS active_uploaded_at,
                      MAX(processed_at) FILTER (WHERE is_active=true) AS active_processed_at,
                      MAX(id::text) FILTER (WHERE is_active=true) AS active_upload_id
                    FROM public.file_uploads
                    WHERE company_id=:company_id
                      AND document_type='general_ledger'
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        branch = (
            await session.execute(
                text(
                    """
                    SELECT
                      COUNT(*) FILTER (WHERE is_active=true)::int AS active_branches,
                      COUNT(*) FILTER (
                        WHERE is_active=true AND COALESCE(review_status,'accepted') <> 'accepted'
                      )::int AS pending_branches
                    FROM public.branches
                    WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        active_rows = (
            await session.execute(
                text(
                    """
                    SELECT
                      COUNT(gt.id)::int AS transaction_count,
                      COUNT(*) FILTER (WHERE gt.branch_id IS NULL)::int AS unassigned_branch_rows,
                      COUNT(*) FILTER (WHERE gt.file_upload_id IS NULL)::int AS rows_without_source_upload,
                      COUNT(*) FILTER (WHERE gt.transaction_date > CURRENT_DATE)::int AS future_rows,
                      COUNT(*) FILTER (WHERE COALESCE(gt.exchange_rate,1) <= 0)::int AS invalid_fx_rows,
                      COUNT(*) FILTER (
                        WHERE COALESCE(gt.debit,0) <> 0 AND COALESCE(gt.credit,0) <> 0
                      )::int AS both_sides_rows,
                      COUNT(*) FILTER (
                        WHERE COALESCE(gt.debit,0) = 0 AND COALESCE(gt.credit,0) = 0
                      )::int AS zero_value_rows,
                      COALESCE(SUM(gt.net_amount),0) AS consolidated_net
                    FROM public.gl_transactions gt
                    JOIN public.file_uploads fu
                      ON fu.id=gt.file_upload_id
                     AND fu.company_id=:company_id
                     AND fu.is_active=true
                    WHERE gt.company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        branch_net = (
            await session.execute(
                text(
                    """
                    SELECT COALESCE(SUM(branch_total),0)
                    FROM (
                      SELECT gt.branch_id, COALESCE(SUM(gt.net_amount),0) AS branch_total
                      FROM public.gl_transactions gt
                      JOIN public.file_uploads fu
                        ON fu.id=gt.file_upload_id
                       AND fu.company_id=:company_id
                       AND fu.is_active=true
                      WHERE gt.company_id=:company_id
                      GROUP BY gt.branch_id
                    ) totals
                    """
                ),
                {"company_id": company_id},
            )
        ).scalar_one()

        mapping_duplicates = int(
            (
                await session.execute(
                    text(
                        """
                        SELECT COUNT(*)::int
                        FROM (
                          SELECT source_account_code
                          FROM public.finance_account_mappings
                          WHERE company_id=:company_id AND is_confirmed=true
                          GROUP BY source_account_code
                          HAVING COUNT(*) > 1
                        ) duplicates
                        """
                    ),
                    {"company_id": company_id},
                )
            ).scalar_one()
            or 0
        )

        jobs = (
            await session.execute(
                text(
                    """
                    SELECT
                      COUNT(*) FILTER (WHERE status IN ('queued','processing','retry'))::int AS open_jobs,
                      COUNT(*) FILTER (
                        WHERE status='processing'
                          AND updated_at < now() - interval '30 minutes'
                      )::int AS stale_processing_jobs,
                      COUNT(*) FILTER (
                        WHERE status IN ('failed','validation_failed')
                          AND created_at > now() - interval '7 days'
                      )::int AS recent_failed_jobs
                    FROM public.ingestion_jobs
                    WHERE company_id=:company_id
                    """
                ),
                {"company_id": company_id},
            )
        ).mappings().one()

        checks: list[dict[str, Any]] = []

        def add(
            key: str,
            label: str,
            status: str,
            detail: str,
            action: str | None = None,
            category: str = "finance",
            blocking: bool = False,
        ) -> None:
            checks.append(
                {
                    "key": key,
                    "label": label,
                    "status": status,
                    "detail": detail,
                    "action": action,
                    "category": category,
                    "blocking": blocking,
                }
            )

        has_data = int(health["transaction_count"] or 0) > 0
        add(
            "active_dataset",
            "Exactly one active General Ledger dataset",
            "pass" if int(upload["active_count"] or 0) == 1 else "fail",
            f"{int(upload['active_count'] or 0)} active General Ledger dataset(s); {int(upload['total_count'] or 0)} retained upload version(s).",
            "Activate one validated General Ledger dataset before relying on reports."
            if int(upload["active_count"] or 0) != 1
            else None,
            "ingestion",
            blocking=True,
        )

        add(
            "ledger_present",
            "Active ledger contains transactions",
            "pass" if has_data else "fail",
            f"{int(health['transaction_count'] or 0):,} active transaction(s) available.",
            "Upload and validate a General Ledger before certification." if not has_data else None,
            "ingestion",
            blocking=True,
        )

        add(
            "trial_balance",
            "Trial balance balances",
            "pass" if bool(health["is_trial_balance_balanced"]) else "fail",
            f"Debit/credit difference: {health['trial_balance_difference']}.",
            "Investigate source rows/import treatment until debit equals credit."
            if not health["is_trial_balance_balanced"]
            else None,
            "reconciliation",
            blocking=True,
        )

        add(
            "balance_sheet",
            "Balance sheet reconciles",
            "pass" if bool(health["is_balance_sheet_balanced"]) else "fail",
            f"Assets less liabilities and equity: {health['balance_sheet_difference']}.",
            "Review mappings and current-period/retained earnings treatment."
            if not health["is_balance_sheet_balanced"]
            else None,
            "reconciliation",
            blocking=True,
        )

        add(
            "mapping_complete",
            "All active accounts are confirmed mapped",
            "pass" if bool(health["is_mapping_complete"]) else "fail",
            f"{int(health['mapped_account_count'] or 0)} of {int(health['account_count'] or 0)} account(s) mapped; {int(health['unmapped_account_count'] or 0)} unmapped.",
            "Complete account mapping before management reports are certified."
            if not health["is_mapping_complete"]
            else None,
            "mapping",
            blocking=True,
        )

        add(
            "mapping_uniqueness",
            "Confirmed mappings are unique per account",
            "pass" if mapping_duplicates == 0 else "fail",
            f"{mapping_duplicates} account code(s) have multiple confirmed mappings.",
            "Resolve duplicate confirmed mappings." if mapping_duplicates else None,
            "mapping",
            blocking=True,
        )

        invalid_rows = int(health["invalid_transaction_count"] or 0)
        add(
            "valid_transactions",
            "Active transactions passed validation",
            "pass" if invalid_rows == 0 else "fail",
            f"{invalid_rows} invalid active transaction row(s).",
            "Correct invalid source rows and reload the ledger." if invalid_rows else None,
            "ingestion",
            blocking=True,
        )

        source_missing = int(active_rows["rows_without_source_upload"] or 0)
        add(
            "source_traceability",
            "Every active transaction traces to an upload",
            "pass" if source_missing == 0 else "fail",
            f"{source_missing} active transaction row(s) have no source upload reference.",
            "Repair source lineage before relying on reports." if source_missing else None,
            "traceability",
            blocking=True,
        )

        consolidated = self._number(active_rows["consolidated_net"])
        branch_total = self._number(branch_net)
        branch_difference = consolidated - branch_total
        branch_reconciles = abs(branch_difference) <= Decimal("0.01")
        add(
            "branch_raw_reconciliation",
            "Consolidated ledger reconciles to branch grouping",
            "pass" if branch_reconciles else "fail",
            f"Consolidated net {consolidated}; grouped branch/unassigned net {branch_total}; difference {branch_difference}.",
            "Investigate branch assignment or aggregation logic." if not branch_reconciles else None,
            "branches",
            blocking=True,
        )

        active_branches = int(branch["active_branches"] or 0)
        pending_branches = int(branch["pending_branches"] or 0)
        unassigned = int(active_rows["unassigned_branch_rows"] or 0)
        if active_branches == 0:
            branch_status = "pass"
            branch_detail = "No branch structure is configured; consolidated reporting only."
            branch_action = None
        elif pending_branches:
            branch_status = "warning"
            branch_detail = f"{pending_branches} active branch(es) still require review; {unassigned} transaction row(s) are unassigned."
            branch_action = "Review detected branches before using branch-level reporting."
        elif unassigned:
            branch_status = "warning"
            branch_detail = f"{active_branches} accepted active branch(es); {unassigned} transaction row(s) remain unassigned."
            branch_action = "Confirm whether unassigned rows belong at company level or should be allocated to a branch."
        else:
            branch_status = "pass"
            branch_detail = f"{active_branches} active branch(es), all reviewed; all active transactions have branch assignments."
            branch_action = None
        add("branch_coverage", "Branch structure is review-complete", branch_status, branch_detail, branch_action, "branches")

        duplicates = int(health["duplicate_candidate_count"] or 0)
        add(
            "duplicate_screen",
            "Potential duplicate transactions screened",
            "pass" if duplicates == 0 else "warning",
            f"{duplicates} potential duplicate transaction row(s) detected using date/account/amount/document/description matching.",
            "Review the duplicate candidates; repeated accounting entries can be legitimate, so this is a review flag rather than an automatic deletion."
            if duplicates
            else None,
            "ingestion",
        )

        future_rows = int(active_rows["future_rows"] or 0)
        add(
            "future_dates",
            "Transaction dates are not unexpectedly future-dated",
            "pass" if future_rows == 0 else "warning",
            f"{future_rows} active transaction(s) have a transaction date after today.",
            "Confirm future-dated journals are intentional." if future_rows else None,
            "periods",
        )

        invalid_fx = int(active_rows["invalid_fx_rows"] or 0)
        add(
            "exchange_rates",
            "Exchange rates are positive",
            "pass" if invalid_fx == 0 else "fail",
            f"{invalid_fx} transaction(s) have a zero/negative exchange rate.",
            "Correct exchange-rate data before consolidated analysis." if invalid_fx else None,
            "ingestion",
            blocking=invalid_fx > 0,
        )

        zero_rows = int(active_rows["zero_value_rows"] or 0)
        both_sides = int(active_rows["both_sides_rows"] or 0)
        amount_status = "pass" if zero_rows == 0 and both_sides == 0 else "warning"
        add(
            "journal_shape",
            "Journal amount shape reviewed",
            amount_status,
            f"{zero_rows} zero-value row(s); {both_sides} row(s) contain both debit and credit amounts.",
            "Review unusual rows to confirm they reflect the source system rather than parsing errors."
            if amount_status == "warning"
            else None,
            "ingestion",
        )

        missing_months = 0
        history_months = 0
        if monthly:
            months = sorted({row["month"] for row in monthly})
            if months:
                history_months = self._month_index(months[-1]) - self._month_index(months[0]) + 1
                missing_months = max(history_months - len(months), 0)

        if not has_data:
            period_status = "fail"
            period_detail = "No period history is available."
            period_action = "Load a General Ledger."
        elif missing_months:
            period_status = "warning"
            period_detail = f"{missing_months} month(s) are missing within {history_months} month(s) of ledger coverage."
            period_action = "Confirm whether missing accounting periods are intentional."
        elif history_months < 6:
            period_status = "warning"
            period_detail = f"Only {history_months or 1} month(s) of history are available."
            period_action = "Load more historical periods where available to improve trend and forecast reliability."
        else:
            period_status = "pass"
            period_detail = f"{history_months} continuous month(s) of history are available."
            period_action = None
        add("period_coverage", "Historical period coverage", period_status, period_detail, period_action, "periods", blocking=not has_data)

        last_date = health.get("last_transaction_date")
        age_days = None
        if last_date:
            age_days = (date.today() - last_date).days
        recency_status = "warning" if age_days is not None and age_days > 90 else "pass"
        add(
            "dataset_recency",
            "Active dataset recency",
            recency_status,
            "No latest transaction date available."
            if age_days is None
            else f"Latest transaction date is {last_date} ({max(age_days,0)} day(s) before today).",
            "Confirm the active ledger is current enough for the management decision being made."
            if recency_status == "warning"
            else None,
            "periods",
        )

        stale_jobs = int(jobs["stale_processing_jobs"] or 0)
        recent_failed = int(jobs["recent_failed_jobs"] or 0)
        if stale_jobs:
            job_status = "fail"
            job_action = "Investigate stale background import jobs before loading more files."
        elif recent_failed:
            job_status = "warning"
            job_action = "Review recent failed imports and confirm the active dataset is the intended successful version."
        else:
            job_status = "pass"
            job_action = None
        add(
            "ingestion_jobs",
            "Background ingestion is healthy",
            job_status,
            f"{int(jobs['open_jobs'] or 0)} open job(s); {stale_jobs} stale processing job(s); {recent_failed} failed/validation-failed job(s) in the last 7 days.",
            job_action,
            "ingestion",
            blocking=stale_jobs > 0,
        )

        blocking_failures = [c for c in checks if c["blocking"] and c["status"] == "fail"]
        warnings = [c for c in checks if c["status"] == "warning"]
        failures = [c for c in checks if c["status"] == "fail"]

        if blocking_failures:
            status = "blocked"
        elif failures or warnings:
            status = "attention"
        else:
            status = "ready"

        pass_count = sum(1 for c in checks if c["status"] == "pass")
        warning_count = len(warnings)
        fail_count = len(failures)
        score = round(
            (pass_count + warning_count * Decimal("0.5"))
            / max(len(checks), 1)
            * 100
        )

        return {
            "status": status,
            "score": int(score),
            "pass_count": pass_count,
            "warning_count": warning_count,
            "fail_count": fail_count,
            "checks": checks,
            "active_upload_id": upload["active_upload_id"],
            "first_transaction_date": health.get("first_transaction_date"),
            "last_transaction_date": health.get("last_transaction_date"),
            "assurance_score": int(assurance["score"]),
            "assurance_grade": str(assurance["grade"]),
            "certified_at": datetime.now(timezone.utc),
            "caveat": (
                "FinCruiz Reliability Certification tests structural consistency, reconciliation, "
                "lineage, mapping, branch coverage and ingestion state of the active dataset. "
                "It does not replace audit, source-accounting review, tax advice or statutory assurance."
            ),
        }
