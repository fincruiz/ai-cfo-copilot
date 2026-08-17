from typing import Annotated

from fastapi import APIRouter, Depends
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.dependencies.subscription import require_entitlement
from app.schemas.finance.ai_cfo import AICFOAnswerResponse, AICFOQuestionRequest, AICFOSignalsResponse
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
    result = await service.answer(current_company.id, request.question, request.include_external_context)
    return APIResponse(
        message="AI CFO response generated.",
        data=AICFOAnswerResponse(**result),
    )


@router.get("/executive-brief", response_model=APIResponse[AICFOAnswerResponse])
async def executive_brief(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[AICFOService, Depends(get_service)],
):
    question = "Give me a proactive executive CFO briefing. Identify the 3 most material internal movements or risks in the loaded company data, connect them to relevant current economic or industry conditions when useful, and recommend the 3 highest-priority management actions for the next 30-90 days."
    result = await service.answer(current_company.id, question, True)
    return APIResponse(message="Executive briefing generated.", data=AICFOAnswerResponse(**result))

@router.get("/industry-benchmark", response_model=APIResponse[AICFOAnswerResponse])
async def industry_benchmark(
    current_company: Annotated[Company, Depends(get_current_company)],
    _entitlement: Annotated[object, Depends(require_entitlement('benchmarking'))],
    service: Annotated[AICFOService, Depends(get_service)],
):
    question = "Benchmark our profitability, growth, working capital and balance-sheet position against credible current information for our industry and country. Clearly label which comparisons are supported by external sources, avoid inventing unavailable benchmarks, and give management implications."
    result = await service.answer(current_company.id, question, True)
    return APIResponse(message="Industry benchmark generated.", data=AICFOAnswerResponse(**result))


@router.get("/signals", response_model=APIResponse[AICFOSignalsResponse])
async def proactive_signals(
    current_company: Annotated[Company, Depends(get_current_company)],
    service: Annotated[AICFOService, Depends(get_service)],
):
    result = await service.proactive_signals(current_company.id)
    return APIResponse(message="Proactive finance signals generated.", data=AICFOSignalsResponse(**result))
