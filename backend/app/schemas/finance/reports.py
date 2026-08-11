from datetime import date
from decimal import Decimal
from uuid import UUID

from pydantic import BaseModel


class ReportLineResponse(BaseModel):
    code: str
    label: str
    amount: Decimal
    order: int
    is_total: bool = False


class TrialBalanceResponse(BaseModel):
    total_debit: Decimal
    total_credit: Decimal
    difference: Decimal
    lines: list[ReportLineResponse]


class ProfitAndLossResponse(BaseModel):
    revenue: Decimal
    cost_of_sales: Decimal
    gross_profit: Decimal
    operating_expenses: Decimal
    operating_profit: Decimal
    depreciation: Decimal
    ebit: Decimal
    other_income: Decimal
    other_expenses: Decimal
    finance_costs: Decimal
    profit_before_tax: Decimal
    tax: Decimal
    net_profit: Decimal
    lines: list[ReportLineResponse]


class BalanceSheetResponse(BaseModel):
    current_assets: Decimal
    non_current_assets: Decimal
    total_assets: Decimal
    current_liabilities: Decimal
    non_current_liabilities: Decimal
    total_liabilities: Decimal
    contributed_equity: Decimal
    current_period_earnings: Decimal
    equity: Decimal
    total_liabilities_and_equity: Decimal
    balance_difference: Decimal
    lines: list[ReportLineResponse]


class RatioResponse(BaseModel):
    name: str
    category: str
    value: Decimal | None
    unit: str
    status: str
    tone: str
    interpretation: str


class MonthlyActualResponse(BaseModel):
    month: date
    revenue: Decimal
    cost_of_sales: Decimal
    gross_profit: Decimal
    operating_expenses: Decimal
    depreciation: Decimal
    ebit: Decimal
    finance_costs: Decimal
    tax: Decimal
    net_profit: Decimal


class BranchComparisonResponse(BaseModel):
    branch_id: UUID
    branch_code: str
    branch_name: str
    revenue: Decimal
    gross_profit: Decimal
    operating_expenses: Decimal
    ebit: Decimal
    net_profit: Decimal
    gross_margin_percent: Decimal | None
    net_margin_percent: Decimal | None
