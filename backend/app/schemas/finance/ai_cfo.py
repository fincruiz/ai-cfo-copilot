from pydantic import BaseModel, Field


class AICFOQuestionRequest(BaseModel):
    question: str = Field(min_length=2, max_length=1000)


class AICFOAnswerResponse(BaseModel):
    answer: str
    mode: str
    suggested_questions: list[str]
