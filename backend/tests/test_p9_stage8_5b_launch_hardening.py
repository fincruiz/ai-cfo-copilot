from types import SimpleNamespace
from datetime import datetime, timedelta, timezone

from app.services.integrations.health import integration_health
from app.services.launch_readiness_service import build_launch_readiness


def company(**overrides):
    base=dict(legal_name="Acme",country_code="AU",currency_code="AUD",industry="Services",business_model="Service")
    base.update(overrides); return SimpleNamespace(**base)


def test_launch_readiness_guides_empty_workspace_to_data():
    result=build_launch_readiness(company=company(), workspace={"has_financial_data":False,"upload_count":0,"transaction_count":0,"mapping_count":0}, connections=[])
    assert result["score"] == 25
    assert result["next_label"] == "Business data"
    assert result["next_path"] == "/dashboard/integrations"


def test_launch_readiness_reaches_management_use():
    result=build_launch_readiness(company=company(), workspace={"has_financial_data":True,"upload_count":1,"transaction_count":200,"mapping_count":12}, connections=[])
    assert result["score"] == 100
    assert result["ready_for_management_use"] is True


def test_integration_health_flags_failures_and_stale_sources():
    failed=integration_health({"status":"connected","configured":True,"last_sync_status":"failed","last_sync_message":"token expired"})
    assert failed["health_status"] == "failed"
    stale=integration_health({"status":"connected","configured":True,"last_sync_status":"success","last_synced_at":datetime.now(timezone.utc)-timedelta(hours=80)})
    assert stale["health_status"] == "stale"
    healthy=integration_health({"status":"connected","configured":True,"last_sync_status":"success","last_synced_at":datetime.now(timezone.utc)})
    assert healthy["health_status"] == "healthy"
