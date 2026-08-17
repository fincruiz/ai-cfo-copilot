from typing import Annotated

from fastapi import APIRouter, Depends, Query

from app.database.models.core.company import Company
from app.dependencies.company import get_current_company
from app.schemas.market import MarketProfileOut
from app.schemas.responses import APIResponse
from app.services.market_service import MARKETS, market_payload

router = APIRouter(prefix="/markets", tags=["Markets & Localisation"])


@router.get("/catalog", response_model=APIResponse[list[MarketProfileOut]])
async def market_catalog() -> APIResponse[list[MarketProfileOut]]:
    data = [MarketProfileOut(**market_payload(code)) for code in ["IN", "AU", "AE", "GB", "US"]]
    return APIResponse(message="Regional market catalog retrieved.", data=data)


@router.get("/resolve", response_model=APIResponse[MarketProfileOut])
async def resolve_market(country_code: Annotated[str, Query(min_length=2, max_length=2)] = "AU") -> APIResponse[MarketProfileOut]:
    return APIResponse(message="Market profile resolved.", data=MarketProfileOut(**market_payload(country_code)))


@router.get("/current", response_model=APIResponse[MarketProfileOut])
async def current_market(company: Annotated[Company, Depends(get_current_company)]) -> APIResponse[MarketProfileOut]:
    return APIResponse(message="Company market profile retrieved.", data=MarketProfileOut(**market_payload(company.country_code)))
