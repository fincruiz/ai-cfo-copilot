from typing import Annotated
from uuid import UUID

from fastapi import APIRouter, Depends, File, Form, Query, UploadFile
from fastapi.responses import Response
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.auth import get_current_user
from app.dependencies.company import get_current_company, require_company_admin
from app.schemas.auth import CurrentUser
from app.schemas.beta import BetaFeedbackUpdate
from app.schemas.responses import APIResponse
from app.services.beta_feedback_service import BetaFeedbackService

router=APIRouter(prefix="/beta-feedback",tags=["Beta Feedback"])
ALLOWED_CATEGORIES={"bug","incorrect_number","ai_answer","confusing_ux","feature_request","other"}
ALLOWED_SEVERITIES={"p0","p1","p2"}
ALLOWED_IMAGE_TYPES={"image/png","image/jpeg","image/webp"}
MAX_ATTACHMENT_BYTES=2*1024*1024

@router.post("")
async def create_feedback(
    current_company:Annotated[Company,Depends(get_current_company)],
    current_user:Annotated[CurrentUser,Depends(get_current_user)],
    session:Annotated[AsyncSession,Depends(get_db_session)],
    category:Annotated[str,Form()],severity:Annotated[str,Form()],title:Annotated[str,Form()],description:Annotated[str,Form()],
    path:Annotated[str,Form()]="",app_version:Annotated[str|None,Form()]=None,browser:Annotated[str|None,Form()]=None,
    viewport:Annotated[str|None,Form()]=None,request_id:Annotated[str|None,Form()]=None,
    screenshot:UploadFile|None=File(default=None),
):
    if category not in ALLOWED_CATEGORIES: raise ApplicationError(message="Unsupported feedback category.",error_code="INVALID_FEEDBACK_CATEGORY",status_code=422)
    if severity not in ALLOWED_SEVERITIES: raise ApplicationError(message="Unsupported feedback severity.",error_code="INVALID_FEEDBACK_SEVERITY",status_code=422)
    if len(title.strip())<3 or len(description.strip())<3: raise ApplicationError(message="Add a short title and description.",error_code="FEEDBACK_DETAIL_REQUIRED",status_code=422)
    mime=None; raw=None
    if screenshot:
        mime=(screenshot.content_type or "").lower()
        if mime not in ALLOWED_IMAGE_TYPES: raise ApplicationError(message="Screenshot must be PNG, JPEG or WebP.",error_code="INVALID_FEEDBACK_ATTACHMENT",status_code=415)
        raw=await screenshot.read(MAX_ATTACHMENT_BYTES+1)
        if len(raw)>MAX_ATTACHMENT_BYTES: raise ApplicationError(message="Screenshot must be 2 MB or smaller.",error_code="FEEDBACK_ATTACHMENT_TOO_LARGE",status_code=413)
    data=await BetaFeedbackService(session).create(company_id=current_company.id,user_id=current_user.id,category=category,severity=severity,title=title.strip(),description=description.strip(),path=path,app_version=app_version,browser=browser,viewport=viewport,request_id=request_id,attachment_mime=mime,attachment_bytes=raw)
    return APIResponse(message="Feedback captured for beta review.",data=data)

@router.get("")
async def list_feedback(current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin),status: str|None=Query(default=None),severity:str|None=Query(default=None)):
    return APIResponse(message="Beta feedback retrieved.",data=await BetaFeedbackService(session).list(company_id=current_company.id,status=status,severity=severity))

@router.get("/summary")
async def feedback_summary(current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    return APIResponse(message="Beta feedback summary retrieved.",data=await BetaFeedbackService(session).summary(company_id=current_company.id))

@router.patch("/{feedback_id}")
async def update_feedback(feedback_id:UUID,payload:BetaFeedbackUpdate,current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    data=await BetaFeedbackService(session).update(company_id=current_company.id,feedback_id=feedback_id,status=payload.status,resolution_notes=payload.resolution_notes)
    if not data: raise ApplicationError(message="Feedback item not found.",error_code="FEEDBACK_NOT_FOUND",status_code=404)
    return APIResponse(message="Feedback status updated.",data=data)

@router.get("/{feedback_id}/attachment")
async def feedback_attachment(feedback_id:UUID,current_company:Annotated[Company,Depends(get_current_company)],session:Annotated[AsyncSession,Depends(get_db_session)],_admin=Depends(require_company_admin)):
    row=await BetaFeedbackService(session).attachment(company_id=current_company.id,feedback_id=feedback_id)
    if not row: raise ApplicationError(message="Feedback attachment not found.",error_code="FEEDBACK_ATTACHMENT_NOT_FOUND",status_code=404)
    return Response(content=bytes(row["attachment_bytes"]),media_type=row["attachment_mime"],headers={"Cache-Control":"private, no-store"})
