from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.schemas.finance.ai_cfo import AICFOAnswerResponse, AICFOQuestionRequest
from app.schemas.responses import APIResponse
from app.services.finance.ai_cfo_service import AICFOService

router = APIRouter(prefix="/ai-cfo", tags=["AI CFO"])


def get_service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
) -> AICFOService:
    return AICFOService(session)


@router.post("/ask", response_model=APIResponse[AICFOAnswerResponse])
async def ask_ai_cfo(
    request: AICFOQuestionRequest,
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[AICFOService, Depends(get_service)],
):
    result = await service.answer(current_company.id, request.question)
    return APIResponse(
        message="AI CFO response generated.",
        data=AICFOAnswerResponse(**result),
    )
