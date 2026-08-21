from __future__ import annotations

from typing import Annotated

from fastapi import APIRouter, Depends, Header, HTTPException
from fastapi.responses import RedirectResponse
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.config import settings
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import (
    get_current_company,
    require_company_admin,
    require_finance_write,
)
from app.schemas.auth import CurrentUser
from app.schemas.integrations import TallyPushRequest, TenantSelection
from app.schemas.responses import APIResponse
from app.services.audit_service import AuditService
from app.services.integrations.base import IntegrationStore
from app.services.integrations.finance_truth import CanonicalIntegrationGLService
from app.services.integrations.health import integration_health
from app.services.integrations.tally import normalize_tally_record
from app.services.integrations.xero import XeroConnector
from app.services.integrations.zoho import ZohoConnector

router = APIRouter(prefix="/integrations", tags=["Integrations"])


def _front(status: str, provider: str):
    return (
        f"{settings.integration_frontend_url.rstrip('/')}/dashboard/integrations"
        f"?provider={provider}&status={status}"
    )


@router.get("")
async def list_integrations(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    existing = {
        item["provider"]: item
        for item in await IntegrationStore(session).list_connections(company.id)
    }
    data = []
    for provider, configured in [
        ("xero", bool(settings.xero_client_id and settings.xero_client_secret and settings.xero_redirect_uri)),
        ("zoho", bool(settings.zoho_client_id and settings.zoho_client_secret and settings.zoho_redirect_uri)),
        ("tally", True),
    ]:
        item = existing.get(provider) or {
            "provider": provider,
            "status": "disconnected",
            "configured": configured,
            "metadata": {},
        }
        item["configured"] = configured
        item.update(integration_health(item))
        data.append(item)
    return APIResponse(message="Integrations retrieved.", data=data)


@router.post("/xero/start")
async def xero_start(
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    connector = XeroConnector(IntegrationStore(session))
    if not connector.configured():
        raise HTTPException(503, "Xero credentials are not configured on the server.")
    return APIResponse(
        message="Xero authorization ready.",
        data={"authorization_url": await connector.authorization_url(company.id, user.id)},
    )


@router.get("/xero/callback")
async def xero_callback(
    code: str,
    state: str,
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    try:
        await XeroConnector(IntegrationStore(session)).callback(code, state)
        return RedirectResponse(_front("connected", "xero"))
    except Exception:
        return RedirectResponse(_front("error", "xero"))


@router.post("/xero/select-tenant")
async def xero_select(
    payload: TenantSelection,
    company: Annotated[Company, Depends(get_current_company)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    await XeroConnector(IntegrationStore(session)).select_tenant(
        company.id, payload.tenant_id
    )
    return APIResponse(message="Xero organisation selected.", data=True)


@router.post("/xero/sync")
async def xero_sync(
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_finance_write)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    source_counts = await XeroConnector(IntegrationStore(session)).sync(company.id)
    finance_truth = await CanonicalIntegrationGLService(session).activate(
        company=company,
        provider="xero",
        activated_by=user.id,
    )
    audit_metadata = {
        "source_records": source_counts,
        "finance_truth": finance_truth,
    }
    await AuditService(session).record(
        company_id=company.id,
        user_id=user.id,
        action="sync",
        module="integrations",
        summary="Synced Xero data and evaluated canonical GL activation.",
        metadata=audit_metadata,
        commit=True,
    )
    return APIResponse(
        message="Xero sync complete.",
        data=audit_metadata,
    )


@router.post("/zoho/start")
async def zoho_start(
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    connector = ZohoConnector(IntegrationStore(session))
    if not connector.configured():
        raise HTTPException(503, "Zoho credentials are not configured on the server.")
    return APIResponse(
        message="Zoho authorization ready.",
        data={"authorization_url": await connector.authorization_url(company.id, user.id)},
    )


@router.get("/zoho/callback")
async def zoho_callback(
    code: str,
    state: str,
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    try:
        await ZohoConnector(IntegrationStore(session)).callback(code, state)
        return RedirectResponse(_front("connected", "zoho"))
    except Exception:
        return RedirectResponse(_front("error", "zoho"))


@router.post("/zoho/select-tenant")
async def zoho_select(
    payload: TenantSelection,
    company: Annotated[Company, Depends(get_current_company)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    await ZohoConnector(IntegrationStore(session)).select_tenant(
        company.id, payload.tenant_id
    )
    return APIResponse(message="Zoho Books organisation selected.", data=True)


@router.post("/zoho/sync")
async def zoho_sync(
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_finance_write)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    source_counts = await ZohoConnector(IntegrationStore(session)).sync(company.id)
    finance_truth = await CanonicalIntegrationGLService(session).activate(
        company=company,
        provider="zoho",
        activated_by=user.id,
    )
    audit_metadata = {
        "source_records": source_counts,
        "finance_truth": finance_truth,
    }
    await AuditService(session).record(
        company_id=company.id,
        user_id=user.id,
        action="sync",
        module="integrations",
        summary="Synced Zoho Books data and evaluated canonical GL activation.",
        metadata=audit_metadata,
        commit=True,
    )
    return APIResponse(
        message="Zoho Books sync complete.",
        data=audit_metadata,
    )


@router.post("/tally/bridge-token")
async def tally_token(
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    token = await IntegrationStore(session).tally_bridge_token(company.id, user.id)
    return APIResponse(
        message="Tally bridge token created. Copy it now; it is shown once.",
        data={"bridge_token": token, "ingest_url": "/api/v1/integrations/tally/push"},
    )


@router.post("/tally/push")
async def tally_push(
    payload: TallyPushRequest,
    authorization: Annotated[str | None, Header()] = None,
    session: Annotated[AsyncSession, Depends(get_db_session)] = None,
):
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "Tally bridge token required.")

    store = IntegrationStore(session)
    company_id = await store.company_for_bridge(authorization.split(" ", 1)[1])
    if not company_id:
        raise HTTPException(401, "Invalid Tally bridge token.")

    if (payload.snapshot_start or payload.snapshot_complete) and not payload.snapshot_id:
        raise HTTPException(422, "Tally snapshot_id is required when starting or completing a ledger snapshot.")

    connection = await store.get(company_id, "tally") or {}
    current_snapshot = (connection.get("metadata") or {}).get("tally_snapshot") or {}
    current_snapshot_id = current_snapshot.get("snapshot_id")

    if payload.snapshot_start:
        # A new full-ledger snapshot must not inherit stale lines from the prior snapshot.
        await store.clear_records(company_id, "tally", "gl_line", commit=False)
        await store.merge_metadata(
            company_id,
            "tally",
            {
                "tally_snapshot": {
                    "snapshot_id": payload.snapshot_id,
                    "status": "collecting",
                    "snapshot_start": True,
                    "snapshot_complete": False,
                }
            },
            commit=False,
        )
    elif payload.snapshot_id:
        if not current_snapshot_id or str(current_snapshot_id) != str(payload.snapshot_id):
            raise HTTPException(409, "Tally snapshot does not match the active bridge snapshot. Start a new snapshot before sending this chunk.")

    normalized = [normalize_tally_record(item.model_dump()) for item in payload.records]

    grouped: dict[str, list[dict]] = {}
    for row in normalized:
        grouped.setdefault(row["entity_type"], []).append(row)
    for entity_type, rows in grouped.items():
        await store.replace_records(company_id, "tally", entity_type, rows)

    await store.upsert_connection(
        company_id=company_id,
        provider="tally",
        status="connected",
    )

    finance_truth = {
        "status": "collecting",
        "provider": "tally",
        "message": "Tally snapshot chunk accepted; active GL unchanged until the bridge marks the snapshot complete.",
        "active_ledger_changed": False,
        "canonical_rows": 0,
    }
    if payload.snapshot_complete:
        company = await session.get(Company, company_id)
        if company is None:
            raise HTTPException(404, "Company not found for Tally bridge.")
        finance_truth = await CanonicalIntegrationGLService(session).activate(
            company=company,
            provider="tally",
            activated_by=None,
        )

    await store.merge_metadata(
        company_id,
        "tally",
        {
            "tally_snapshot": {
                "snapshot_id": payload.snapshot_id,
                "status": "complete" if payload.snapshot_complete else "collecting",
                "snapshot_start": payload.snapshot_start,
                "snapshot_complete": payload.snapshot_complete,
                "records_received_last_chunk": len(payload.records),
            },
            "finance_truth": finance_truth,
        },
    )
    await store.mark_sync(
        company_id,
        "tally",
        "success",
        (
            f"Received {len(payload.records)} Tally records. "
            f"Finance truth status: {finance_truth['status']}."
        ),
    )
    return APIResponse(
        message="Tally data accepted.",
        data={
            "records": len(payload.records),
            "snapshot_complete": payload.snapshot_complete,
            "finance_truth": finance_truth,
        },
    )


@router.delete("/{provider}")
async def disconnect(
    provider: str,
    company: Annotated[Company, Depends(get_current_company)],
    user: Annotated[CurrentUser, Depends(get_current_user)],
    _: Annotated[object, Depends(require_company_admin)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    if provider not in {"xero", "zoho", "tally"}:
        raise HTTPException(404, "Unknown provider.")
    purge = await CanonicalIntegrationGLService(session).purge_provider(
        company_id=company.id, provider=provider
    )
    await IntegrationStore(session).disconnect(company.id, provider, True)
    await AuditService(session).record(
        company_id=company.id,
        user_id=user.id,
        action="disconnect",
        module="integrations",
        summary=f"Disconnected {provider} and deleted its synchronized FinCruiz copy.",
        metadata=purge,
        commit=True,
    )
    return APIResponse(
        message=f"{provider.title()} disconnected and synchronized FinCruiz data deleted.",
        data=purge,
    )
