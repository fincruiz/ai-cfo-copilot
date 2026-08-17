from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company, require_company_admin
from app.schemas.responses import APIResponse
from app.schemas.subscription import BetaReadinessCheck, BetaReadinessOut, SubscriptionStatusOut
from app.services.core.workspace_lifecycle_service import WorkspaceLifecycleService
from app.services.integrations.base import IntegrationStore
from app.services.subscription_service import SubscriptionService

router = APIRouter(prefix="/subscription", tags=["Subscription & Beta"])


@router.get("/status", response_model=APIResponse[SubscriptionStatusOut])
async def subscription_status(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    data = await SubscriptionService(session).status(company_id=company.id)
    return APIResponse(message="Subscription status retrieved.", data=SubscriptionStatusOut(**data))


@router.get("/beta-readiness", response_model=APIResponse[BetaReadinessOut])
async def beta_readiness(
    company: Annotated[Company, Depends(get_current_company)],
    session: Annotated[AsyncSession, Depends(get_db_session)],
    _admin=Depends(require_company_admin),
):
    workspace = await WorkspaceLifecycleService(session).status(company_id=company.id)
    subscription = await SubscriptionService(session).status(company_id=company.id)
    connections = await IntegrationStore(session).list_connections(company.id)

    connected = [x for x in connections if x.get("status") == "connected"]
    successful_sync = [x for x in connected if x.get("last_sync_status") in (None, "success")]

    checks: list[BetaReadinessCheck] = []
    checks.append(BetaReadinessCheck(
        key="subscription",
        label="Plan access",
        status="ready" if subscription["is_access_active"] else "blocked",
        detail=f"{subscription['plan'].title()} plan · {subscription['status'].replace('_', ' ')}",
    ))
    checks.append(BetaReadinessCheck(
        key="financial_data",
        label="Financial data",
        status="ready" if workspace.get("has_financial_data") else "attention",
        detail=(
            f"{workspace.get('transaction_count', 0):,} ledger transactions available."
            if workspace.get("has_financial_data")
            else "Load company data or a synthetic demo workspace before a customer walkthrough."
        ),
    ))
    checks.append(BetaReadinessCheck(
        key="mapping",
        label="Account mapping",
        status="ready" if workspace.get("mapping_count", 0) > 0 else "attention",
        detail=f"{workspace.get('mapping_count', 0)} confirmed mapping(s).",
    ))
    checks.append(BetaReadinessCheck(
        key="integration",
        label="Connected source",
        status="ready" if successful_sync else ("attention" if connected else "attention"),
        detail=(
            f"{len(successful_sync)} healthy connected source(s)."
            if successful_sync
            else "Uploads can be used for beta; connect Xero when live sync is required."
        ),
    ))

    weights = {"ready": 25, "attention": 12, "blocked": 0}
    score = min(100, sum(weights[item.status] for item in checks))
    overall = "ready" if score >= 85 else "attention" if score >= 50 else "blocked"
    return APIResponse(
        message="Beta readiness retrieved.",
        data=BetaReadinessOut(score=score, status=overall, checks=checks),
    )
