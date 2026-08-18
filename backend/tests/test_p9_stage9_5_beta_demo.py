from app.services.demo_ai_service import deterministic_demo_answer


def test_demo_ai_never_requires_customer_context():
    result=deterministic_demo_answer("Why is cash tight?")
    assert "₹1.18M" in result["answer"]
    assert "synthetic" in result["confidence_reason"].lower()
    assert result["evidence"]


def test_demo_ai_refuses_to_invent_out_of_scope_evidence():
    result=deterministic_demo_answer("How many employees are in the warehouse?")
    assert "won't invent" in result["answer"]


def test_feedback_migration_keeps_screenshot_private():
    from pathlib import Path
    text=Path("migrations/20260818_p9_stage9_5_beta_feedback.sql").read_text()
    assert "attachment_bytes bytea" in text
    assert "ENABLE ROW LEVEL SECURITY" in text
    assert "uploads/" not in text
