from collections import defaultdict, deque
from time import monotonic
from typing import Annotated
from urllib.parse import urlparse

from fastapi import APIRouter, Depends, Query, Request
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.session import get_db_session
from app.dependencies.company import require_company_admin
from app.schemas.responses import APIResponse
from app.services.marketing_event_service import ALLOWED_EVENTS, MarketingEventService

router=APIRouter(prefix="/marketing",tags=["Marketing telemetry"])
_hits: dict[str, deque[float]] = defaultdict(deque)

def _allowed(ip: str) -> bool:
    now=monotonic(); bucket=_hits[ip]
    while bucket and now-bucket[0] > 60: bucket.popleft()
    if len(bucket) >= 60: return False
    bucket.append(now); return True

@router.post("/events")
async def marketing_event(request: Request, session:Annotated[AsyncSession,Depends(get_db_session)]):
    ip=request.client.host if request.client else "unknown"
    if not _allowed(ip): return APIResponse(message="Event ignored.",data={"recorded":False})
    payload=await request.json()
    event_name=str(payload.get("event_name") or "")
    if event_name not in ALLOWED_EVENTS: return APIResponse(message="Event ignored.",data={"recorded":False})
    referrer=str(payload.get("referrer") or "")
    host=urlparse(referrer).hostname if referrer else None
    await MarketingEventService(session).record(
      event_name=event_name, session_id=str(payload.get("session_id") or "anonymous"),
      path=str(payload.get("path") or "/"), referrer_host=host, properties=payload.get("properties") or {}
    )
    return APIResponse(message="Event accepted.",data={"recorded":True})

@router.get("/funnel")
async def marketing_funnel(session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin),days:int=Query(30,ge=1,le=365)):
    return APIResponse(message="Homepage conversion funnel retrieved.",data=await MarketingEventService(session).funnel(days=days))
