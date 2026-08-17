from app.services.commercial_onboarding_service import determine_commercial_onboarding_stage

def test_onboarding_requires_data_first():
    assert determine_commercial_onboarding_stage(transaction_count=0, pending_branches=0, unmapped_accounts=0) == ("data_needed", "/dashboard/uploads?welcome=1", "Load your finance data")

def test_onboarding_requires_branch_review_before_mapping():
    stage, path, _ = determine_commercial_onboarding_stage(transaction_count=100, pending_branches=2, unmapped_accounts=4)
    assert stage == "branch_review_required"
    assert path == "/dashboard/branches?welcome=1"

def test_onboarding_requires_mapping_after_structure_review():
    stage, path, _ = determine_commercial_onboarding_stage(transaction_count=100, pending_branches=0, unmapped_accounts=4)
    assert stage == "mapping_required"
    assert path == "/dashboard/mapping?welcome=1"

def test_onboarding_reaches_ready_gate_when_structure_and_mapping_are_complete():
    stage, path, _ = determine_commercial_onboarding_stage(transaction_count=100, pending_branches=0, unmapped_accounts=0)
    assert stage == "ready"
    assert path == "/dashboard"
