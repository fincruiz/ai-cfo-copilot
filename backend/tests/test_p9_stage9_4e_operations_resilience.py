from app.services.operations_service import grade_database_latency, overall_status
from scripts.load_test_finance import certification_status, percentile


def test_database_latency_thresholds(monkeypatch):
    from app.services import operations_service
    monkeypatch.setattr(operations_service.settings, "database_degraded_ms", 500)
    monkeypatch.setattr(operations_service.settings, "database_unhealthy_ms", 1500)

    assert grade_database_latency(100) == "healthy"
    assert grade_database_latency(700) == "degraded"
    assert grade_database_latency(2000) == "unhealthy"


def test_overall_operational_status_uses_worst_check():
    assert overall_status([{"status": "healthy"}, {"status": "healthy"}]) == "healthy"
    assert overall_status([{"status": "healthy"}, {"status": "degraded"}]) == "degraded"
    assert overall_status([{"status": "healthy"}, {"status": "unhealthy"}]) == "unhealthy"


def test_performance_certification_thresholds():
    assert certification_status(success_percent=100, p95_ms=800, success_target=99, p95_target_ms=1500) == "ready"
    assert certification_status(success_percent=98, p95_ms=800, success_target=99, p95_target_ms=1500) == "attention"
    assert certification_status(success_percent=100, p95_ms=1800, success_target=99, p95_target_ms=1500) == "attention"
    assert certification_status(success_percent=94, p95_ms=800, success_target=99, p95_target_ms=1500) == "blocked"


def test_percentile_is_deterministic():
    assert percentile([10, 20, 30, 40, 50], 0.95) == 40
