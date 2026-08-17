from app.services.finance.ai_cfo_service import AICFOService


def base_context():
    return {
        "company": {"currency": "AUD"},
        "monthly_actuals": [
            {"month": "2026-06", "revenue": 100000, "gross_profit": 40000, "net_profit": 10000},
            {"month": "2026-07", "revenue": 110000, "gross_profit": 41800, "net_profit": 8000},
        ],
        "branch_comparison": [
            {"branch_name": "Melbourne", "revenue": 80000, "net_profit": 9000},
            {"branch_name": "Sydney", "revenue": 30000, "net_profit": -1000},
        ],
        "ar_summary": {"total_outstanding": 50000, "overdue_amount": 20000, "overdue_percent": 40.0},
        "ap_summary": {"total_outstanding": 20000, "overdue_amount": 2000, "overdue_percent": 10.0},
        "balance_sheet": {"current_assets": 100000, "current_liabilities": 95000},
    }


def test_short_followup_preserves_previous_management_topic():
    resolved = AICFOService._resolve_follow_up(
        "Which branch caused most of it?",
        [{"role": "user", "content": "Why did profit fall?"}, {"role": "assistant", "content": "Profit fell."}],
    )
    assert "Why did profit fall?" in resolved
    assert "Which branch caused most of it?" in resolved


def test_branch_followup_identifies_weakest_branch_from_loaded_data():
    result = AICFOService._executive_finance_answer(
        "Why did profit fall? Follow-up: Which branch caused most of it?",
        base_context(),
    )
    assert result is not None
    answer, action = result
    assert "Sydney" in answer
    assert "-1,000.00" in answer
    assert action["route"] == "/dashboard/visual-bi"


def test_cash_pressure_answer_is_explicitly_indicator_based():
    result = AICFOService._executive_finance_answer("What is putting pressure on cash?", base_context())
    assert result is not None
    answer, action = result
    assert "overdue receivables" in answer
    assert "40.0% of AR" in answer
    assert "indicators" in answer
    assert action["route"] == "/dashboard/working-capital"


def test_followup_without_history_stays_unchanged():
    assert AICFOService._resolve_follow_up("Show me the trend", []) == "Show me the trend"
