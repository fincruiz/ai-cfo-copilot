from collections import defaultdict, deque
from time import time
from fastapi import APIRouter, Request

from app.core.exceptions import ApplicationError
from app.schemas.beta import DemoAnswer, DemoQuestionRequest
from app.schemas.responses import APIResponse
from app.services.demo_ai_service import answer_demo_question

router=APIRouter(prefix="/demo",tags=["Public Demo"])
_hits:dict[str,deque[float]]=defaultdict(deque)
WINDOW_SECONDS=300
MAX_QUESTIONS=20

def _allow(key:str)->bool:
 now=time();q=_hits[key]
 while q and q[0] < now-WINDOW_SECONDS:q.popleft()
 if len(q)>=MAX_QUESTIONS:return False
 q.append(now);return True

@router.post("/ask",response_model=APIResponse[DemoAnswer])
async def ask_demo(payload:DemoQuestionRequest,request:Request):
 key=(request.client.host if request.client else "unknown")[:100]
 if not _allow(key):raise ApplicationError(message="Demo question limit reached. Try again in a few minutes.",error_code="DEMO_RATE_LIMIT",status_code=429)
 result=await answer_demo_question(payload.question.strip())
 return APIResponse(message="Synthetic demo answer generated.",data=DemoAnswer(**result))
