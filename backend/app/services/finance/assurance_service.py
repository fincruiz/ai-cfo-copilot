from __future__ import annotations

from datetime import date
from uuid import UUID

from app.services.finance.reporting_service import ReportingService


class FinancialAssuranceService:
    """Deterministic structural checks for finance data.

    This does not assert that source accounting records are economically correct;
    it verifies that the loaded dataset is internally consistent enough for analysis.
    """

    def __init__(self, reporting: ReportingService) -> None:
        self.reporting = reporting

    @staticmethod
    def _month_index(value: date) -> int:
        return value.year * 12 + value.month

    async def assess(self, company_id: UUID) -> dict:
        health = await self.reporting.data_health(company_id)
        monthly = await self.reporting.monthly_actuals(company_id)
        checks: list[dict] = []

        def add(key: str, label: str, passed: bool, weight: int, detail: str, action: str | None = None, partial: bool = False):
            score = weight if passed else (weight // 2 if partial else 0)
            checks.append({
                "key": key,
                "label": label,
                "status": "pass" if passed else ("warning" if partial else "fail"),
                "score": score,
                "max_score": weight,
                "detail": detail,
                "action": action,
            })

        has_data = health["transaction_count"] > 0
        add(
            "ledger_present", "Ledger loaded", has_data, 10,
            f"{health['transaction_count']:,} active ledger transactions are available." if has_data else "No active General Ledger transactions are loaded.",
            "Upload a General Ledger CSV to begin." if not has_data else None,
        )
        add(
            "trial_balance", "Debits equal credits", bool(health["is_trial_balance_balanced"]), 20,
            f"Trial balance difference is {health['trial_balance_difference']}.",
            "Investigate source rows or import rules until the difference is zero." if not health["is_trial_balance_balanced"] else None,
        )
        add(
            "balance_sheet", "Balance sheet reconciles", bool(health["is_balance_sheet_balanced"]), 20,
            f"Balance sheet difference is {health['balance_sheet_difference']}.",
            "Review mappings and retained/current-period earnings treatment." if not health["is_balance_sheet_balanced"] else None,
        )
        add(
            "mapping", "Accounts mapped", bool(health["is_mapping_complete"]), 20,
            f"{health['mapped_account_count']} of {health['account_count']} accounts are confirmed mapped.",
            "Review unmapped accounts before relying on reports." if not health["is_mapping_complete"] else None,
        )
        valid = health["invalid_transaction_count"] == 0
        add(
            "valid_rows", "Transaction validation", valid, 10,
            f"{health['invalid_transaction_count']} invalid transaction rows detected.",
            "Correct invalid rows and re-upload the source file." if not valid else None,
        )
        no_dupes = health["duplicate_candidate_count"] == 0
        add(
            "duplicates", "Duplicate screening", no_dupes, 10,
            f"{health['duplicate_candidate_count']} potential duplicate rows detected.",
            "Review duplicates before finalising management reports." if not no_dupes else None,
            partial=health["duplicate_candidate_count"] > 0,
        )

        continuity = True
        missing_months = 0
        if len(monthly) >= 2:
            months = sorted(row["month"] for row in monthly)
            expected = self._month_index(months[-1]) - self._month_index(months[0]) + 1
            missing_months = max(expected - len(set(months)), 0)
            continuity = missing_months == 0
        elif has_data:
            continuity = False
        add(
            "period_continuity", "Period continuity", continuity, 10,
            "Monthly history is continuous." if continuity else f"{missing_months or 'One or more'} month(s) appear missing from the active history.",
            "Confirm that all required accounting periods were included in the upload." if not continuity else None,
            partial=has_data and not continuity,
        )

        total = sum(item["score"] for item in checks)
        maximum = sum(item["max_score"] for item in checks) or 100
        score = round(total / maximum * 100)
        for item in checks:
            item.pop("max_score", None)
        grade = "A" if score >= 90 else "B" if score >= 80 else "C" if score >= 65 else "D" if score >= 50 else "E"
        status = "ready" if score >= 90 else "review" if score >= 65 else "not_ready"
        return {
            "score": score,
            "grade": grade,
            "status": status,
            "checks": checks,
            "caveat": "This score checks structural consistency and reconciliation of the data loaded into FinCruiz. It does not replace source-accounting review, audit, tax advice or statutory assurance.",
        }
