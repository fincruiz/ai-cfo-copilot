from scripts.launch_certify import Gate, decision, exit_code


def test_launch_decision_go_only_when_every_gate_passes():
    gates = [Gate("a", "A", "pass", "ok"), Gate("b", "B", "pass", "ok")]
    assert decision(gates) == "GO"
    assert exit_code(gates) == 0


def test_launch_decision_conditional_go_without_blocker():
    gates = [Gate("a", "A", "pass", "ok"), Gate("b", "B", "conditional", "pending")]
    assert decision(gates) == "CONDITIONAL GO"
    assert exit_code(gates) == 1


def test_launch_decision_no_go_when_any_gate_fails():
    gates = [Gate("a", "A", "conditional", "pending"), Gate("b", "B", "fail", "blocked")]
    assert decision(gates) == "NO-GO"
    assert exit_code(gates) == 2


def test_payment_condition_is_explicit():
    gate = Gate("billing", "Billing", "conditional", "sandbox pending", live_payment_blocker=True)
    assert gate.live_payment_blocker is True
