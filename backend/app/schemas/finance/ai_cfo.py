from typing import Literal
from pydantic import BaseModel, Field


class AICFOConversationTurn(BaseModel):
    role: Literal["user", "assistant"]
    content: str = Field(min_length=1, max_length=2000)


class AICFOQuestionRequest(BaseModel):
    question: str = Field(min_length=2, max_length=1000)
    include_external_context: bool = True
    conversation: list[AICFOConversationTurn] = Field(default_factory=list, max_length=8)


class AICFOSource(BaseModel):
    title: str
    url: str


class AICFOAction(BaseModel):
    label: str
    route: str


class AICFOEvidence(BaseModel):
    label: str
    value: str
    source: str
    period: str | None = None


class AICFODecisionHandoff(BaseModel):
    scenario_type: str
    title: str
    assumptions: dict[str, float | int | str] = {}
    route: str


class AIChartSeries(BaseModel):
    name: str
    data: list[float]


class AIVisualization(BaseModel):
    type: Literal["line", "area", "bar", "stacked_bar", "donut", "waterfall"]
    title: str
    subtitle: str | None = None
    labels: list[str]
    series: list[AIChartSeries]
    value_format: Literal["number", "currency", "percent"] = "number"
    currency: str | None = None


class AICFOAnswerResponse(BaseModel):
    answer: str
    mode: str
    suggested_questions: list[str]
    sources: list[AICFOSource] = []
    action: AICFOAction | None = None
    external_context_used: bool = False
    visualization: AIVisualization | None = None
    evidence: list[AICFOEvidence] = []
    confidence: Literal["high", "medium", "low"] = "medium"
    confidence_reason: str | None = None
    decision_handoff: AICFODecisionHandoff | None = None
    interpreted_question: str | None = None


class AICFOSignal(BaseModel):
    severity: str
    title: str
    evidence: str
    action: str


class AICFOSignalsResponse(BaseModel):
    signals: list[AICFOSignal] = []
    generated_from_months: int = 0
