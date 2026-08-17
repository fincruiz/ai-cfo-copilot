from app.services.finance.ai_cfo_service import AICFOService


def context():
    return {
        "company": {"currency": "AUD"},
        "monthly_actuals": [
            {"month": "2026-06", "revenue": 100000, "gross_profit": 40000, "net_profit": 10000},
            {"month": "2026-07", "revenue": 110000, "gross_profit": 41800, "net_profit": 8000},
        ],
        "ar_summary": {"overdue_percent": 28.0},
    }


def test_profit_question_explains_movement_not_just_totals():
    result = AICFOService._executive_finance_answer("Why did profit change this month?", context())
    assert result is not None
    answer, action = result
    assert "down 20.0%" in answer
    assert "Revenue grew 10.0%" in answer
    assert "lagging" in answer
    assert action["route"] == "/dashboard/analytics"


def test_management_priority_uses_real_signals():
    result = AICFOService._executive_finance_answer("What should management focus on today?", context())
    assert result is not None
    answer, _ = result
    assert "revenue is up 10.0%" in answer
    assert "net profit is down 20.0%" in answer
    assert "28.0% of receivables are overdue" in answer


def test_no_monthly_data_does_not_invent_answer():
    assert AICFOService._executive_finance_answer(
        "Why did profit change?", {"company": {"currency": "AUD"}, "monthly_actuals": []}
    ) is None
