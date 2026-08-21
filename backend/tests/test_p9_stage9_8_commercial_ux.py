from datetime import date
from pathlib import Path
from types import SimpleNamespace
from uuid import uuid4

import pytest

from app.services.finance.reporting_service import ReportingService


@pytest.mark.asyncio
async def test_report_context_uses_financial_year_and_latest_valid_actual():
    class Repo:
        session = None

        async def transaction_summary(self, *, company_id, branch_id=None):
            return {
                "transaction_count": 321,
                "first_transaction_date": date(2025, 7, 1),
                "last_transaction_date": date(2026, 8, 20),
            }

        async def latest_transaction_date(self, *, company_id, branch_id=None):
            return date(2026, 8, 20)

    service = ReportingService(Repo())

    async def fy_end(_company_id):
        return 6

    service._financial_year_end_month = fy_end
    result = await service.report_context(uuid4())
    assert result["period_start"] == date(2026, 7, 1)
    assert result["period_end"] == date(2026, 8, 20)
    assert result["data_as_of"] == date(2026, 8, 20)
    assert result["transaction_count"] == 321


def test_single_dashboard_ai_and_contextual_assistant_elsewhere():
    dashboard = Path("../frontend/app/dashboard/page.tsx").read_text(encoding="utf-8")
    floating = Path("../frontend/components/ai-cfo-floating.tsx").read_text(encoding="utf-8")
    contextual = Path("../frontend/components/contextual-ai-bar.tsx").read_text(encoding="utf-8")

    assert dashboard.count("<AskFinCruizDashboard") == 1
    assert 'pathname !== "/dashboard"' in floating
    assert 'if (pathname === "/dashboard") return null;' in contextual
    assert "AI CFO & Platform Guide" not in floating


def test_reporting_period_is_globally_visible_and_truth_backed():
    layout = Path("../frontend/app/dashboard/layout.tsx").read_text(encoding="utf-8")
    indicator = Path("../frontend/components/reporting-period-indicator.tsx").read_text(encoding="utf-8")
    frontend_service = Path("../frontend/services/finance-service.ts").read_text(encoding="utf-8")
    backend_router = Path("app/api/v1/finance/reports/router.py").read_text(encoding="utf-8")

    assert "<ReportingPeriodIndicator/>" in layout
    assert "Data as of" in indicator
    assert "getReportContext" in frontend_service
    assert '@router.get("/context"' in backend_router


def test_ai_evidence_can_drill_into_report_evidence_and_transactions():
    schema = Path("app/schemas/finance/ai_cfo.py").read_text(encoding="utf-8")
    service = Path("app/services/finance/ai_cfo_service.py").read_text(encoding="utf-8")
    reports = Path("../frontend/app/dashboard/reports/page.tsx").read_text(encoding="utf-8")
    frontend_types = Path("../frontend/types/analytics.ts").read_text(encoding="utf-8")

    assert "route: str | None = None" in schema
    assert '"Profit & Loss": "/dashboard/reports?tab=profit-and-loss"' in service
    assert "route?: string | null" in frontend_types
    assert '"transactions"' in reports
    assert "openAccountTransactions" in reports
    assert "Ledger transactions" in reports


def test_management_dashboard_defaults_are_less_dense():
    dashboard = Path("../frontend/app/dashboard/page.tsx").read_text(encoding="utf-8")
    assert 'owner: ["headline", "metrics", "performance", "priorities"]' in dashboard
    assert 'cfo: ["headline", "metrics", "performance", "priorities", "briefing", "confidence"]' in dashboard
