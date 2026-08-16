from app.services.finance.ai_cfo_service import AICFOService
from app.schemas.finance.ai_cfo import AICFOAnswerResponse


def test_decision_handoff_extracts_headcount_without_inventing_financial_assumptions():
    handoff = AICFOService._decision_handoff("Can we afford to hire 3 people?")
    assert handoff is not None
    assert handoff["title"] == "Hiring / payroll decision"
    assert handoff["assumptions"]["headcount_change"] == 3
    assert "from_ai=1" in handoff["route"]
    assert "payroll" not in handoff["assumptions"]


def test_finance_evidence_is_deterministic_and_capped():
    context = {
        "company": {"currency": "INR"},
        "pnl": {"revenue": 1000, "net_profit": 125},
        "balance_sheet": {"total_assets": 2000, "total_liabilities": 800, "equity": 1200},
        "monthly_actuals": [{"month": "2026-07", "revenue": 900, "net_profit": 100}],
    }
    evidence = AICFOService._evidence(context, "Why did profit change this month and what is our cash balance?")
    assert any(x["label"] == "Revenue" and x["value"] == "INR 1,000.00" for x in evidence)
    assert any(x["source"] == "Monthly actuals" for x in evidence)
    assert len(evidence) <= 6


def test_response_schema_accepts_evidence_confidence_and_handoff():
    response = AICFOAnswerResponse(
        answer="Grounded answer", mode="grounded_finance_assistant", suggested_questions=[], sources=[],
        external_context_used=False, evidence=[{"label":"Revenue","value":"INR 1,000.00","source":"Profit & Loss"}],
        confidence="high", confidence_reason="Prepared finance data passed assurance checks.",
        decision_handoff={"scenario_type":"hiring_payroll_decision","title":"Hiring / payroll decision","assumptions":{"headcount_change":3},"route":"/dashboard/three-way-forecast?from_ai=1"},
    )
    assert response.confidence == "high"
    assert response.evidence[0].source == "Profit & Loss"
    assert response.decision_handoff.assumptions["headcount_change"] == 3
