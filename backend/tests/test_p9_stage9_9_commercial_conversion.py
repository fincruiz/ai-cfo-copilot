from pathlib import Path

from app.schemas.marketing import DemoLeadRequest
from app.services.marketing_event_service import ALLOWED_EVENTS, DENIED_PROPERTY_KEYS


BACKEND = Path(__file__).resolve().parents[1]
FRONTEND = BACKEND.parent / "frontend"


def test_demo_lead_schema_requires_business_identity_and_has_honeypot():
    payload = DemoLeadRequest(
        name="Finance Lead",
        work_email="finance@example.com",
        company_name="Example Co",
        persona="finance",
        website="",
    )
    assert payload.work_email == "finance@example.com"
    assert payload.website == ""


def test_demo_lead_endpoint_is_rate_limited_and_server_side():
    router = (BACKEND / "app/api/v1/marketing/router.py").read_text(encoding="utf-8")
    assert '@router.post("/demo-leads"' in router
    assert "_lead_allowed" in router
    assert "SalesLeadService" in router
    assert "payload.website" in router


def test_sales_leads_are_rls_protected_without_public_policy():
    migration = (BACKEND / "migrations/20260821_p9_stage9_9_sales_leads.sql").read_text(encoding="utf-8")
    assert "CREATE TABLE IF NOT EXISTS public.sales_leads" in migration
    assert "ENABLE ROW LEVEL SECURITY" in migration
    assert "CREATE POLICY" not in migration.upper()


def test_public_conversion_pages_and_persona_close_exist():
    for relative in ["app/book-demo/page.tsx", "app/security/page.tsx", "app/privacy/page.tsx", "app/trust/page.tsx"]:
        assert (FRONTEND / relative).exists(), relative
    demo = (FRONTEND / "app/demo/page.tsx").read_text(encoding="utf-8")
    assert "closeHeadline" in demo
    assert "/book-demo?persona=${audience}" in demo
    assert 'marketingService.track("demo_book_demo_clicked"' in demo


def test_customer_proof_defaults_empty_and_requires_approval_reference():
    proof = (FRONTEND / "lib/customer-proof.ts").read_text(encoding="utf-8")
    homepage = (FRONTEND / "app/page.tsx").read_text(encoding="utf-8")
    assert "permission_reference" in proof
    assert "approvedCustomerProof: ApprovedCustomerProof[] = [];" in proof
    assert "approvedCustomerProof.length > 0" in homepage


def test_marketing_telemetry_does_not_store_contact_free_text():
    assert {"homepage_book_demo_clicked", "demo_book_demo_clicked", "demo_lead_submitted"} <= ALLOWED_EVENTS
    assert {"email", "name", "message"} <= DENIED_PROPERTY_KEYS
