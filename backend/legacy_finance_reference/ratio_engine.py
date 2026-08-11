from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from typing import Any

import pandas as pd


def _num(value: Any) -> float:
    try:
        return float(pd.to_numeric(value, errors="coerce")) if pd.notna(value) else 0.0
    except Exception:
        return 0.0


def _period_days(profile: dict) -> int:
    try:
        start = pd.to_datetime(profile.get("Period Start Date"))
        end = pd.to_datetime(profile.get("Period End Date"))
        days = int((end - start).days) + 1
        return max(1, days)
    except Exception:
        label = str(profile.get("Report Period", "")).lower()
        if "quarter" in label:
            return 91
        if "month" in label:
            return 30
        if "half" in label or "six" in label:
            return 182
        return 365


def _frame_lookup(df: pd.DataFrame | None, label_cols: list[str], value_col: str, keywords: list[str], *, absolute: bool = True) -> float:
    if df is None or df.empty or value_col not in df.columns:
        return 0.0
    labels = pd.Series("", index=df.index, dtype="object")
    for col in label_cols:
        if col in df.columns:
            labels = labels + " " + df[col].fillna("").astype(str)
    mask = pd.Series(False, index=df.index)
    for keyword in keywords:
        mask = mask | labels.str.contains(keyword, case=False, regex=False)
    values = pd.to_numeric(df.loc[mask, value_col], errors="coerce").fillna(0)
    result = float(values.sum())
    return abs(result) if absolute else result


def _pnl_line(pnl: pd.DataFrame | None, keywords: list[str]) -> float:
    return _frame_lookup(pnl, ["Reporting Group", "Reporting Subgroup"], "Report Value", keywords)


def _bs_line(bs: pd.DataFrame | None, keywords: list[str]) -> float:
    return _frame_lookup(bs, ["Reporting Group", "Reporting Subgroup"], "Balance", keywords)


def _summary_total(summary: Any) -> float:
    if summary is None:
        return 0.0
    if isinstance(summary, dict):
        for key in ["Total", "Total AR", "Total AP", "Outstanding", "Balance", "Value"]:
            if key in summary:
                return abs(_num(summary[key]))
        for value in summary.values():
            if isinstance(value, (int, float)):
                return abs(float(value))
    if isinstance(summary, pd.DataFrame) and not summary.empty:
        for col in ["Amount", "Balance", "Outstanding", "Value"]:
            if col in summary.columns:
                return float(pd.to_numeric(summary[col], errors="coerce").fillna(0).abs().sum())
    return 0.0


def _safe_div(num: float, den: float, multiplier: float = 1.0) -> float | None:
    return (num / den * multiplier) if den not in (0, 0.0, None) else None


def _status(name: str, value: float | None) -> tuple[str, str]:
    if value is None:
        return "Not available", "neutral"
    rules = {
        "Current Ratio": (1.5, 1.0, True),
        "Quick Ratio": (1.0, 0.7, True),
        "Cash Ratio": (0.5, 0.2, True),
        "Debt to Equity": (1.0, 2.0, False),
        "Debt to Assets": (0.4, 0.65, False),
        "DSO": (45, 70, False),
        "DPO": (45, 75, False),
        "DIO": (60, 100, False),
        "Cash Conversion Cycle": (60, 100, False),
        "Gross Margin": (35, 20, True),
        "Net Profit Margin": (10, 3, True),
        "Return on Assets": (8, 3, True),
        "Return on Equity": (15, 5, True),
        "Asset Turnover": (1.5, 0.8, True),
    }
    good, warning, higher_is_better = rules.get(name, (None, None, True))
    if good is None:
        return "Monitor", "neutral"
    if higher_is_better:
        if value >= good:
            return "Healthy", "good"
        if value >= warning:
            return "Watch", "warning"
        return "Action required", "bad"
    if value <= good:
        return "Healthy", "good"
    if value <= warning:
        return "Watch", "warning"
    return "Action required", "bad"


def calculate_management_ratios(state: dict, people_data: dict | None = None) -> pd.DataFrame:
    pnl = state.get("consolidated_pnl")
    bs = state.get("consolidated_bs")
    profile = state.get("company_profile", {}) or {}
    days = _period_days(profile)

    revenue = _pnl_line(pnl, ["total revenue", "revenue", "sales"])
    cogs = _pnl_line(pnl, ["total cost of sales", "cost of sales", "cogs"])
    gross_profit = _pnl_line(pnl, ["gross profit"])
    net_profit = _pnl_line(pnl, ["net profit"])
    if not gross_profit and revenue and cogs:
        gross_profit = revenue - cogs

    current_assets = _bs_line(bs, ["current asset"])
    current_liabilities = _bs_line(bs, ["current liabil"])
    cash = _bs_line(bs, ["cash", "bank"])
    receivables = _bs_line(bs, ["receivable", "trade debtor", "accounts receivable"])
    inventory = _bs_line(bs, ["inventory", "stock"])
    payables = _bs_line(bs, ["payable", "trade creditor", "accounts payable"])
    debt = _bs_line(bs, ["loan", "borrow", "debt", "finance lease"])
    equity = _bs_line(bs, ["equity", "capital", "retained earning"])
    total_assets = _bs_line(bs, ["total asset"])
    if not total_assets:
        total_assets = _bs_line(bs, ["asset"])

    ar_total = _summary_total(state.get("ar_summary")) or receivables
    ap_total = _summary_total(state.get("ap_summary")) or payables

    people_data = people_data or state.get("board_people_data", {}) or {}
    employees = _num(people_data.get("Total employees"))

    metrics: list[tuple[str, str, float | None, str, str]] = [
        ("Current Ratio", "Liquidity", _safe_div(current_assets, current_liabilities), "x", "Ability to cover short-term obligations with current assets."),
        ("Quick Ratio", "Liquidity", _safe_div(current_assets - inventory, current_liabilities), "x", "Liquidity excluding inventory."),
        ("Cash Ratio", "Liquidity", _safe_div(cash, current_liabilities), "x", "Immediate liquidity from cash and bank balances."),
        ("DSO", "Working Capital", _safe_div(ar_total, revenue, days), "days", "Average collection period for receivables."),
        ("DPO", "Working Capital", _safe_div(ap_total, cogs, days), "days", "Average payment period for suppliers."),
        ("DIO", "Working Capital", _safe_div(inventory, cogs, days), "days", "Average days inventory remains on hand."),
        ("Cash Conversion Cycle", "Working Capital", None, "days", "DSO plus DIO less DPO."),
        ("Debt to Equity", "Leverage", _safe_div(debt, equity), "x", "Debt funding relative to shareholder equity."),
        ("Debt to Assets", "Leverage", _safe_div(debt, total_assets), "x", "Share of assets financed by debt."),
        ("Gross Margin", "Profitability", _safe_div(gross_profit, revenue, 100), "%", "Gross profit earned from each unit of revenue."),
        ("Net Profit Margin", "Profitability", _safe_div(net_profit, revenue, 100), "%", "Net profit retained from each unit of revenue."),
        ("Return on Assets", "Returns", _safe_div(net_profit, total_assets, 100), "%", "Profit generated from the asset base."),
        ("Return on Equity", "Returns", _safe_div(net_profit, equity, 100), "%", "Return generated on shareholder capital."),
        ("Asset Turnover", "Efficiency", _safe_div(revenue, total_assets), "x", "Revenue generated per unit of assets."),
        ("Revenue per Employee", "People Productivity", _safe_div(revenue, employees), "currency", "Revenue generated per employee for the reporting period."),
        ("Gross Profit per Employee", "People Productivity", _safe_div(gross_profit, employees), "currency", "Gross profit generated per employee."),
        ("Net Profit per Employee", "People Productivity", _safe_div(net_profit, employees), "currency", "Net profit generated per employee."),
    ]
    dso = next(v for n, _, v, _, _ in metrics if n == "DSO")
    dpo = next(v for n, _, v, _, _ in metrics if n == "DPO")
    dio = next(v for n, _, v, _, _ in metrics if n == "DIO")
    ccc = (dso + dio - dpo) if all(v is not None for v in [dso, dpo, dio]) else None
    metrics = [(n, c, ccc if n == "Cash Conversion Cycle" else v, u, d) for n, c, v, u, d in metrics]

    rows = []
    for name, category, value, unit, description in metrics:
        status, tone = _status(name, value)
        rows.append({"Ratio": name, "Category": category, "Value": value, "Unit": unit, "Status": status, "Tone": tone, "Interpretation": description})
    return pd.DataFrame(rows)


def format_ratio(value: float | None, unit: str, currency: str = "AUD") -> str:
    if value is None or pd.isna(value):
        return "Not available"
    if unit == "%":
        return f"{value:,.1f}%"
    if unit == "days":
        return f"{value:,.1f} days"
    if unit == "x":
        return f"{value:,.2f}x"
    if unit == "currency":
        return f"{currency} {value:,.0f}"
    return f"{value:,.2f}"
