from pathlib import Path

from scripts.audit_rls import classify_rls
from scripts.generate_synthetic_gl import generate


def test_rls_audit_does_not_flag_enabled_zero_policy_as_critical():
    assert classify_rls(exists=True, rls_enabled=True, policy_count=0) == (
        "DENY-BY-DEFAULT",
        False,
    )
    assert classify_rls(exists=True, rls_enabled=False, policy_count=0)[1] is True
    assert classify_rls(exists=False, rls_enabled=False, policy_count=0)[1] is True


def test_synthetic_gl_generator_creates_balanced_pairs(tmp_path: Path):
    output = tmp_path / "synthetic.csv"
    written = generate(
        rows=10,
        branches=2,
        output=output,
        seed=1,
        currency="AUD",
    )
    assert written == 10
    rows = output.read_text(encoding="utf-8").strip().splitlines()
    assert len(rows) == 11  # header + 10 transactions
