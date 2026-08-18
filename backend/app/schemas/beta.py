from datetime import datetime
from pydantic import BaseModel, Field


class BetaFeedbackUpdate(BaseModel):
    status: str = Field(pattern="^(open|reviewing|fixed|closed)$")
    resolution_notes: str | None = Field(default=None, max_length=4000)


class DemoQuestionRequest(BaseModel):
    question: str = Field(min_length=2, max_length=500)


class DemoAnswer(BaseModel):
    answer: str
    mode: str = "synthetic_demo"
    evidence: list[dict] = Field(default_factory=list)
    confidence: str = "high"
    confidence_reason: str
    suggested_questions: list[str] = Field(default_factory=list)
    visualization: dict | None = None
    action: dict | None = None
