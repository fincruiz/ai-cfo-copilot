from app.api.v1.health import _grade_latency

def test_health_latency_grading():
    assert _grade_latency(20) == "healthy"
    assert _grade_latency(249.9) == "healthy"
    assert _grade_latency(250) == "degraded"
    assert _grade_latency(999.9) == "degraded"
    assert _grade_latency(1000) == "unhealthy"
