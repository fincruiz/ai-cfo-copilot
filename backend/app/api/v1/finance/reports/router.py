from datetime import date
from typing import Annotated
from uuid import UUID

from fastapi import APIRouter, Depends, Query
from sqlalchemy.ext.asyncio import AsyncSession

from app.database.models.core.company import Company
from app.database.session import get_db_session
from app.dependencies.company import get_current_company
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.schemas.finance.reports import (
    BalanceSheetResponse,
    DataHealthResponse,
    BranchComparisonResponse,
    MonthlyActualResponse,
    ProfitAndLossResponse,
    RatioResponse,
    ReportLineResponse,
    TrialBalanceResponse,
)
from app.schemas.finance.assurance import FinancialAssuranceResponse
from app.schemas.finance.reliability import FinanceReliabilityResponse
from app.services.finance.assurance_service import FinancialAssuranceService
from app.services.finance.reliability_service import FinanceReliabilityService
from app.schemas.responses import APIResponse
from app.services.finance.reporting_service import ReportingService

router = APIRouter(prefix="/reports", tags=["Finance Reports"])


def service(
    session: Annotated[AsyncSession, Depends(get_db_session)],
):
    return ReportingService(GLTransactionRepository(session))


def lines(items):
    return [
        ReportLineResponse(
            code=item.code,
            label=item.label,
            amount=item.amount,
            order=item.order,
            is_total=item.is_total,
        )
        for item in items
    ]


@router.get("/trial-balance", response_model=APIResponse[TrialBalanceResponse])
async def trial_balance(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    start_date: date | None = None,
    end_date: date | None = None,
    branch_id: UUID | None = None,
):
    report = await svc.trial_balance(
        current_company.id,
        start_date,
        end_date,
        branch_id,
    )
    return APIResponse(
        message="Trial balance retrieved.",
        data=TrialBalanceResponse(
            total_debit=report.total_debit,
            total_credit=report.total_credit,
            difference=report.difference,
            lines=[
                ReportLineResponse(
                    code=item.account_code,
                    label=item.account_name or item.account_code,
                    amount=item.signed_amount,
                    order=0,
                )
                for item in report.accounts
            ],
        ),
    )


@router.get("/assurance", response_model=APIResponse[FinancialAssuranceResponse])
async def assurance(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
):
    result = await FinancialAssuranceService(svc).assess(current_company.id)
    return APIResponse(
        message="Financial assurance checks completed.",
        data=FinancialAssuranceResponse(**result),
    )


@router.get("/reliability", response_model=APIResponse[FinanceReliabilityResponse])
async def reliability(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
):
    result = await FinanceReliabilityService(svc).certify(current_company.id)
    return APIResponse(
        message="Finance reliability certification completed.",
        data=FinanceReliabilityResponse(**result),
    )


@router.get("/data-health", response_model=APIResponse[DataHealthResponse])
async def data_health(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
):
    result = await svc.data_health(current_company.id)
    return APIResponse(
        message="Finance data health retrieved.",
        data=DataHealthResponse(**result),
    )


@router.get("/profit-and-loss", response_model=APIResponse[ProfitAndLossResponse])
async def pnl(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    start_date: date | None = None,
    end_date: date | None = None,
    branch_id: UUID | None = None,
):
    report = await svc.pnl(
        current_company.id,
        start_date,
        end_date,
        branch_id,
    )
    return APIResponse(
        message="Profit and loss retrieved.",
        data=ProfitAndLossResponse(
            **{
                key: getattr(report, key)
                for key in ProfitAndLossResponse.model_fields
                if key != "lines"
            },
            lines=lines(report.lines),
        ),
    )


@router.get("/balance-sheet", response_model=APIResponse[BalanceSheetResponse])
async def balance_sheet(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    end_date: date | None = None,
    branch_id: UUID | None = None,
):
    report = await svc.balance_sheet(
        current_company.id,
        end_date,
        branch_id,
    )
    return APIResponse(
        message="Balance sheet retrieved.",
        data=BalanceSheetResponse(
            **{
                key: getattr(report, key)
                for key in BalanceSheetResponse.model_fields
                if key != "lines"
            },
            lines=lines(report.lines),
        ),
    )


@router.get("/kpis", response_model=APIResponse[list[RatioResponse]])
async def kpis(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    start_date: date | None = None,
    end_date: date | None = None,
    branch_id: UUID | None = None,
    period_days: Annotated[int, Query(ge=1, le=366)] = 365,
):
    ratios = await svc.kpis(
        current_company.id,
        start_date,
        end_date,
        period_days,
        branch_id=branch_id,
    )
    return APIResponse(
        message="KPIs retrieved.",
        data=[RatioResponse(**ratio.__dict__) for ratio in ratios],
    )


@router.get("/monthly-actuals", response_model=APIResponse[list[MonthlyActualResponse]])
async def monthly_actuals(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    branch_id: UUID | None = None,
    start_date: date | None = None,
    end_date: date | None = None,
):
    rows = await svc.monthly_actuals(
        current_company.id,
        branch_id=branch_id,
        start_date=start_date,
        end_date=end_date,
    )
    return APIResponse(
        message="Monthly actuals retrieved.",
        data=[MonthlyActualResponse(**row) for row in rows],
    )


@router.get("/branch-comparison", response_model=APIResponse[list[BranchComparisonResponse]])
async def branch_comparison(
    current_company: Annotated[Company, Depends(get_current_company)],
    svc: Annotated[ReportingService, Depends(service)],
    start_date: date | None = None,
    end_date: date | None = None,
):
    rows = await svc.branch_comparison(
        current_company.id,
        start_date,
        end_date,
    )
    return APIResponse(
        message="Branch comparison retrieved.",
        data=[BranchComparisonResponse(**row) for row in rows],
    )
