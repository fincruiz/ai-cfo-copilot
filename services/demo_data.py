from __future__ import annotations

from datetime import date
from typing import Any

import pandas as pd
import streamlit as st

from modules.reporting import (
    build_actuals_by_branch_reporting_group,
    build_ageing_summary,
    build_balance_sheet_detail,
    build_balance_sheet_from_gl,
    build_benchmark_comparison,
    build_coa_mapping_review,
    build_executive_summary,
    build_financial_logic_review,
    build_monthly_actuals,
    build_monthly_branch_actuals,
    build_pnl,
    build_pnl_detail,
    compare_plan_vs_actual,
    compare_pnl_to_forecast,
    compare_pnl_to_previous_year,
    detect_anomalies,
    summarize_plan_vs_actual,
)
from core.pipeline import prepare_data


DEMO_DATA_KEYS = [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped",
    "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary",
    "detected_branches", "validation_passed", "bs_disclaimer", "ai_commentary", "prior_pnl", "prior_bs",
    "prior_kpis", "anomaly_flags", "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df",
    "budget_compare", "budget_summary", "benchmark_df", "py_compare", "benchmark_compare",
    "monthly_actuals", "monthly_branch_actuals", "executive_summary_df", "forecast_pnl", "forecast_bs",
    "previous_year_pnl", "forecast_pnl_compare", "previous_year_pnl_compare", "fx_rate_info",
    "country_indicators", "external_benchmark_df", "consolidated_pnl_detail", "consolidated_bs_detail",
    "coa_duplicate_rows", "coa_mapping_review", "financial_logic_review", "last_validation_report",
    "reporting_structure",
]


TOUR_STEPS = [
    {
        "page": "dashboard",
        "title": "Executive Dashboard",
        "copy": "Start with revenue, profit, working-capital signals, readiness and the AI CFO brief.",
    },
    {
        "page": "reports",
        "title": "Financial Statements",
        "copy": "Review the sample P&L, balance sheet, KPI pack, trends and budget/forecast comparisons.",
    },
    {
        "page": "working_capital",
        "title": "Working Capital Centre",
        "copy": "Explore sample receivables, payables, ageing buckets and overdue exposure.",
    },
    {
        "page": "insights",
        "title": "AI & Management Insights",
        "copy": "See sample anomalies and use the floating AI CFO assistant to ask data-specific questions.",
    },
    {
        "page": "downloads",
        "title": "Download Centre",
        "copy": "Preview the management pack outputs available from the preloaded demonstration workspace.",
    },
]


def _excel_buffer(df: pd.DataFrame):
    """Create an in-memory Excel file compatible with prepare_data()."""
    from io import BytesIO

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    output.seek(0)
    return output


def _build_demo_source_data() -> dict[str, pd.DataFrame]:
    months = pd.period_range("2026-01", "2026-06", freq="M")
    branches = ["Melbourne", "Sydney", "Brisbane"]
    branch_weights = {"Melbourne": 0.44, "Sydney": 0.36, "Brisbane": 0.20}

    account_names = {
        "410000": "Product Revenue",
        "411000": "Service Revenue",
        "510000": "Materials",
        "511000": "Direct Labour",
        "512000": "Freight and Delivery",
        "610000": "Salaries and Wages",
        "611000": "Rent and Occupancy",
        "612000": "Marketing",
        "613000": "Software and IT",
        "614000": "Depreciation",
        "710000": "Other Income",
        "810000": "Interest Expense",
        "910000": "Income Tax Expense",
    }

    rows: list[dict[str, Any]] = []
    base_revenue = [1_720_000, 1_790_000, 1_865_000, 1_940_000, 2_020_000, 2_150_000]
    for month_idx, month in enumerate(months):
        month_end = month.end_time.date()
        total_revenue = base_revenue[month_idx]
        for branch in branches:
            weight = branch_weights[branch]
            product_revenue = total_revenue * weight * 0.72
            service_revenue = total_revenue * weight * 0.28
            materials = product_revenue * (0.39 - month_idx * 0.002)
            direct_labour = total_revenue * weight * 0.115
            freight = total_revenue * weight * (0.044 + month_idx * 0.0015)
            salaries = total_revenue * weight * 0.105
            rent = {"Melbourne": 63_000, "Sydney": 55_000, "Brisbane": 31_000}[branch]
            marketing = total_revenue * weight * (0.026 + month_idx * 0.001)
            software = total_revenue * weight * 0.018
            depreciation = total_revenue * weight * 0.012
            other_income = total_revenue * weight * 0.004
            interest = total_revenue * weight * 0.009
            tax = total_revenue * weight * 0.034

            values = {
                "410000": -product_revenue,
                "411000": -service_revenue,
                "510000": materials,
                "511000": direct_labour,
                "512000": freight,
                "610000": salaries,
                "611000": rent,
                "612000": marketing,
                "613000": software,
                "614000": depreciation,
                "710000": -other_income,
                "810000": interest,
                "910000": tax,
            }
            for account_code, net in values.items():
                rows.append({
                    "Account code": account_code,
                    "Debit": max(net, 0),
                    "Credit": abs(min(net, 0)),
                    "Net": net,
                    "Branch": branch,
                    "Date": month_end,
                    "Period": month.strftime("%b %Y"),
                    "Description": f"Demo {account_names[account_code]} - {branch}",
                    "Account Name": account_names[account_code],
                })

    # Latest balance-sheet snapshot at 30 June 2026.
    bs_values = {
        "110000": ("Cash at Bank", 1_430_000),
        "120000": ("Trade Receivables", 2_280_000),
        "130000": ("Inventory", 1_860_000),
        "150000": ("Property Plant and Equipment", 3_750_000),
        "210000": ("Trade Payables", -1_460_000),
        "220000": ("Accrued Expenses", -620_000),
        "250000": ("Bank Loans", -1_900_000),
        "310000": ("Share Capital", -2_000_000),
        "320000": ("Retained Earnings", -3_340_000),
    }
    for code, (name, net) in bs_values.items():
        rows.append({
            "Account code": code,
            "Debit": max(net, 0),
            "Credit": abs(min(net, 0)),
            "Net": net,
            "Branch": "Consolidated",
            "Date": date(2026, 6, 30),
            "Period": "Jun 2026",
            "Description": f"Demo closing balance - {name}",
            "Account Name": name,
        })

    gl = pd.DataFrame(rows)

    coa_rows = [
        ("410000", "Product Revenue", "Revenue", "Product Revenue", "Income Statement", 1),
        ("411000", "Service Revenue", "Revenue", "Service Revenue", "Income Statement", 1),
        ("510000", "Materials", "Cost of Sales", "Materials", "Income Statement", 2),
        ("511000", "Direct Labour", "Cost of Sales", "Direct Labour", "Income Statement", 2),
        ("512000", "Freight and Delivery", "Cost of Sales", "Freight", "Income Statement", 2),
        ("610000", "Salaries and Wages", "Operating Expense", "Payroll", "Income Statement", 4),
        ("611000", "Rent and Occupancy", "Operating Expense", "Rent", "Income Statement", 4),
        ("612000", "Marketing", "Operating Expense", "Marketing", "Income Statement", 4),
        ("613000", "Software and IT", "Operating Expense", "Technology", "Income Statement", 4),
        ("614000", "Depreciation", "Operating Expense", "Depreciation", "Income Statement", 4),
        ("710000", "Other Income", "Other Income", "Other Income", "Income Statement", 6),
        ("810000", "Interest Expense", "Interest", "Finance Costs", "Income Statement", 8),
        ("910000", "Income Tax Expense", "Tax", "Income Tax", "Income Statement", 9),
        ("110000", "Cash at Bank", "Current Assets", "Cash", "Balance Sheet", 21),
        ("120000", "Trade Receivables", "Current Assets", "Trade Receivables", "Balance Sheet", 21),
        ("130000", "Inventory", "Current Assets", "Inventory", "Balance Sheet", 21),
        ("150000", "Property Plant and Equipment", "Non Current Assets", "PPE", "Balance Sheet", 22),
        ("210000", "Trade Payables", "Current Liabilities", "Trade Payables", "Balance Sheet", 31),
        ("220000", "Accrued Expenses", "Current Liabilities", "Accruals", "Balance Sheet", 31),
        ("250000", "Bank Loans", "Non Current Liabilities", "Borrowings", "Balance Sheet", 32),
        ("310000", "Share Capital", "Equity", "Share Capital", "Balance Sheet", 40),
        ("320000", "Retained Earnings", "Equity", "Retained Earnings", "Balance Sheet", 40),
    ]
    coa = pd.DataFrame(
        [
            {
                "Account code": code,
                "Account Name": name,
                "Reporting Group": group,
                "Reporting Subgroup": subgroup,
                "Statement": statement,
                "Sign Convention": "positive",
                "Display Order": order,
            }
            for code, name, group, subgroup, statement, order in coa_rows
        ]
    )

    budget_rows = []
    for month_idx, month in enumerate(months):
        actual_revenue = base_revenue[month_idx]
        for branch in branches:
            weight = branch_weights[branch]
            budget_rows.extend([
                {"Month": str(month), "Branch": branch, "Reporting Group": "Revenue", "Amount": actual_revenue * weight * 0.985},
                {"Month": str(month), "Branch": branch, "Reporting Group": "Cost of Sales", "Amount": actual_revenue * weight * 0.535},
                {"Month": str(month), "Branch": branch, "Reporting Group": "Operating Expense", "Amount": actual_revenue * weight * 0.205},
            ])
    budget = pd.DataFrame(budget_rows)

    ar = pd.DataFrame([
        {"Party Name": "Atlas Retail Group", "Outstanding Amount": 540_000, "Document Number": "INV-1042", "Document Date": "2026-03-18", "Due Date": "2026-04-17", "Branch": "Melbourne", "Age Bucket": "61-90"},
        {"Party Name": "Brightline Distribution", "Outstanding Amount": 420_000, "Document Number": "INV-1108", "Document Date": "2026-04-12", "Due Date": "2026-05-12", "Branch": "Sydney", "Age Bucket": "31-60"},
        {"Party Name": "Coastal Projects", "Outstanding Amount": 315_000, "Document Number": "INV-1163", "Document Date": "2026-05-02", "Due Date": "2026-06-01", "Branch": "Brisbane", "Age Bucket": "1-30"},
        {"Party Name": "Delta Industrial", "Outstanding Amount": 615_000, "Document Number": "INV-1189", "Document Date": "2026-05-26", "Due Date": "2026-06-25", "Branch": "Melbourne", "Age Bucket": "Current"},
        {"Party Name": "Evergreen Services", "Outstanding Amount": 390_000, "Document Number": "INV-0991", "Document Date": "2026-02-08", "Due Date": "2026-03-10", "Branch": "Sydney", "Age Bucket": "90+"},
    ])
    ar["Document Date"] = pd.to_datetime(ar["Document Date"])
    ar["Due Date"] = pd.to_datetime(ar["Due Date"])

    ap = pd.DataFrame([
        {"Party Name": "National Materials", "Outstanding Amount": 510_000, "Document Number": "BILL-887", "Document Date": "2026-05-14", "Due Date": "2026-06-13", "Branch": "Melbourne", "Age Bucket": "1-30"},
        {"Party Name": "FreightLink Australia", "Outstanding Amount": 265_000, "Document Number": "BILL-914", "Document Date": "2026-04-18", "Due Date": "2026-05-18", "Branch": "Sydney", "Age Bucket": "31-60"},
        {"Party Name": "Cloud Systems", "Outstanding Amount": 165_000, "Document Number": "BILL-952", "Document Date": "2026-06-05", "Due Date": "2026-07-05", "Branch": "Consolidated", "Age Bucket": "Current"},
        {"Party Name": "Metro Property", "Outstanding Amount": 290_000, "Document Number": "BILL-821", "Document Date": "2026-03-22", "Due Date": "2026-04-21", "Branch": "Brisbane", "Age Bucket": "61-90"},
        {"Party Name": "PeopleWorks", "Outstanding Amount": 230_000, "Document Number": "BILL-963", "Document Date": "2026-06-18", "Due Date": "2026-07-18", "Branch": "Consolidated", "Age Bucket": "Current"},
    ])
    ap["Document Date"] = pd.to_datetime(ap["Document Date"])
    ap["Due Date"] = pd.to_datetime(ap["Due Date"])

    benchmark = pd.DataFrame([
        {"Metric": "Gross Margin %", "Benchmark Value": 38.0},
        {"Metric": "Operating Margin %", "Benchmark Value": 12.5},
        {"Metric": "Opex as % of Revenue", "Benchmark Value": 22.0},
        {"Metric": "AR Overdue %", "Benchmark Value": 25.0},
        {"Metric": "AP Overdue %", "Benchmark Value": 25.0},
    ])

    return {"gl": gl, "coa": coa, "budget": budget, "ar": ar, "ap": ap, "benchmark": benchmark}


def _build_kpis_from_pnl(pnl: pd.DataFrame, bs: pd.DataFrame, ar_summary: dict, ap_summary: dict) -> pd.DataFrame:
    values = {}
    if pnl is not None and not pnl.empty:
        for _, row in pnl.iterrows():
            values[str(row.get("Reporting Group", ""))] = float(row.get("Report Value", 0) or 0)
    bs_values = {}
    if bs is not None and not bs.empty:
        bs_values = bs.groupby("Reporting Subgroup")["Balance"].sum().to_dict()

    revenue = values.get("Total Revenue", 0.0)
    gross_profit = values.get("Gross Profit", 0.0)
    total_opex = values.get("Total Overheads", 0.0)
    net_profit = values.get("Net Profit", 0.0)
    cash = float(bs_values.get("Cash", 0.0))
    ar_total = float(ar_summary.get("total", 0.0))
    ap_total = float(ap_summary.get("total", 0.0))
    dso = ar_total / revenue * 181 if revenue else 0.0
    dpo = ap_total / max(values.get("Total COGS", 0.0), 1) * 181

    rows = [
        ("Revenue", revenue, "value"),
        ("Gross Profit", gross_profit, "value"),
        ("Gross Margin %", gross_profit / revenue * 100 if revenue else 0.0, "percent"),
        ("Operating Expenses", total_opex, "value"),
        ("Operating Margin %", (gross_profit - total_opex) / revenue * 100 if revenue else 0.0, "percent"),
        ("Net Profit", net_profit, "value"),
        ("Net Margin %", net_profit / revenue * 100 if revenue else 0.0, "percent"),
        ("Cash", cash, "value"),
        ("DSO", dso, "number"),
        ("DPO", dpo, "number"),
    ]
    return pd.DataFrame([
        {
            "KPI": name,
            "Value": round(value, 2),
            "Output Type": output_type,
            "Display Value": f"{value:.2f}%" if output_type == "percent" else round(value, 2),
        }
        for name, value, output_type in rows
    ])


def load_demo_workspace(force: bool = False) -> None:
    """Populate a complete, read-only demonstration workspace with sample data."""
    if st.session_state.get("demo_data_loaded") and not force:
        return

    source = _build_demo_source_data()
    gl_buffer = _excel_buffer(source["gl"])
    coa_buffer = _excel_buffer(source["coa"])
    gl, coa, _, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
        gl_buffer,
        coa_buffer,
        reporting_structure="Branch / Business Unit Reporting",
    )

    consolidated_pnl = build_pnl(pnl_mapped)
    consolidated_pnl_detail = build_pnl_detail(pnl_mapped)
    consolidated_bs = build_balance_sheet_from_gl(bs_mapped)
    consolidated_bs_detail = build_balance_sheet_detail(bs_mapped)
    ar_summary = build_ageing_summary(source["ar"], "AR")
    ap_summary = build_ageing_summary(source["ap"], "AP")
    consolidated_kpis = _build_kpis_from_pnl(consolidated_pnl, consolidated_bs, ar_summary, ap_summary)

    branches = sorted([b for b in pnl_mapped["Branch"].dropna().unique().tolist() if b != "Consolidated"])
    branch_outputs: dict[str, dict[str, pd.DataFrame]] = {}
    branch_summary_rows = []
    for branch in branches:
        branch_df = pnl_mapped[pnl_mapped["Branch"] == branch].copy()
        branch_pnl = build_pnl(branch_df)
        branch_kpis = _build_kpis_from_pnl(branch_pnl, consolidated_bs, ar_summary, ap_summary)
        branch_outputs[branch] = {
            "pnl": branch_pnl,
            "pnl_detail": build_pnl_detail(branch_df),
            "kpis": branch_kpis,
        }
        row = {"Branch": branch}
        for _, item in branch_kpis.iterrows():
            row[item["KPI"]] = item["Display Value"]
        branch_summary_rows.append(row)
    branch_summary = pd.DataFrame(branch_summary_rows)

    actuals = build_actuals_by_branch_reporting_group(pnl_mapped)
    budget_compare = compare_plan_vs_actual(actuals, source["budget"], "Budget")
    budget_summary = summarize_plan_vs_actual(budget_compare, "Budget")

    # Management-level forecast and prior-year views.
    current_base = consolidated_pnl[~consolidated_pnl["Line Type"].isin(["Total", "Subtotal", "Final Profit"])].copy()
    forecast_pnl = current_base[["Reporting Group", "Reporting Subgroup", "Report Value"]].copy()
    forecast_pnl["Report Value"] = forecast_pnl["Report Value"] * 1.045
    previous_year_pnl = current_base[["Reporting Group", "Reporting Subgroup", "Report Value"]].copy()
    previous_year_pnl["Report Value"] = previous_year_pnl["Report Value"] * 0.91
    forecast_pnl_compare = compare_pnl_to_forecast(current_base, forecast_pnl)
    previous_year_pnl_compare = compare_pnl_to_previous_year(current_base, previous_year_pnl)

    benchmark_compare = build_benchmark_comparison(consolidated_kpis, source["benchmark"], ar_summary, ap_summary)
    monthly_actuals = build_monthly_actuals(pnl_mapped)
    monthly_branch_actuals = build_monthly_branch_actuals(pnl_mapped)
    executive_summary_df = build_executive_summary(
        consolidated_kpis,
        ar_summary=ar_summary,
        ap_summary=ap_summary,
        budget_summary=budget_summary,
        benchmark_compare=benchmark_compare,
        forecast_pnl_compare=forecast_pnl_compare,
        previous_year_pnl_compare=previous_year_pnl_compare,
    )
    anomaly_flags = detect_anomalies(
        consolidated_kpis,
        ar_summary=ar_summary,
        ap_summary=ap_summary,
        budget_summary=budget_summary,
        forecast_pnl_compare=forecast_pnl_compare,
    )

    profile = {
        "Company Name": "Northstar Manufacturing — Demo",
        "Industry": "Manufacturing",
        "Country": "Australia",
        "State / Region": "Victoria",
        "Currency": "AUD",
        "Financial Year": "FY2026",
        "Report Period": "YTD June 2026",
        "Period Start Date": "2026-01-01",
        "Period End Date": "2026-06-30",
        "Reporting Period": "Monthly",
        "Reporting Structure": "Branch / Business Unit Reporting",
        "Tax Identifier": "DEMO-ABN-00000000000",
        "Business Notes": "Sample workspace supplied with AI CFO Copilot for product demonstration only.",
    }

    mapping_review = build_coa_mapping_review(coa)
    financial_logic_review = build_financial_logic_review(consolidated_pnl)
    last_validation_report = {
        "critical": [],
        "warnings": [
            {
                "Area": "Working Capital",
                "Issue": "Receivables over 60 days exceed the demonstration target.",
                "Recommendation": "Review Atlas Retail Group and Evergreen Services first.",
            }
        ],
        "recommendations": [
            {
                "Area": "Gross Margin",
                "Issue": "Freight cost has risen over the last three months.",
                "Recommendation": "Review carrier pricing and customer freight recovery.",
            },
            {
                "Area": "Forecast",
                "Issue": "Revenue momentum is ahead of the sample budget.",
                "Recommendation": "Refresh the rolling forecast and working-capital assumptions.",
            },
        ],
        "info": [{"Area": "Demo", "Issue": "This workspace uses preloaded sample data.", "Recommendation": "Explore freely or exit demo to create a real workspace."}],
        "score": 94,
    }

    values = {
        "company_profile": profile,
        "reporting_structure": profile["Reporting Structure"],
        "gl": gl,
        "coa": coa,
        "kpi_master": None,
        "latest_bs": latest_bs,
        "mapped": mapped,
        "pnl_mapped": pnl_mapped,
        "bs_mapped": bs_mapped,
        "unmapped": unmapped,
        "consolidated_pnl": consolidated_pnl,
        "consolidated_pnl_detail": consolidated_pnl_detail,
        "consolidated_bs": consolidated_bs,
        "consolidated_bs_detail": consolidated_bs_detail,
        "consolidated_kpis": consolidated_kpis,
        "branch_outputs": branch_outputs,
        "branch_summary": branch_summary,
        "detected_branches": branches,
        "validation_passed": True,
        "bs_disclaimer": None,
        "ai_commentary": (
            "Northstar Manufacturing is ahead of budget on revenue, but freight and overdue receivables require attention. "
            "Cash remains healthy, and the sample forecast indicates continued profit growth if collections improve."
        ),
        "ar_df": source["ar"],
        "ap_df": source["ap"],
        "ar_summary": ar_summary,
        "ap_summary": ap_summary,
        "budget_df": source["budget"],
        "budget_compare": budget_compare,
        "budget_summary": budget_summary,
        "benchmark_df": source["benchmark"],
        "benchmark_compare": benchmark_compare,
        "py_compare": None,
        "monthly_actuals": monthly_actuals,
        "monthly_branch_actuals": monthly_branch_actuals,
        "executive_summary_df": executive_summary_df,
        "forecast_pnl": forecast_pnl,
        "forecast_bs": consolidated_bs.copy(),
        "previous_year_pnl": previous_year_pnl,
        "forecast_pnl_compare": forecast_pnl_compare,
        "previous_year_pnl_compare": previous_year_pnl_compare,
        "anomaly_flags": anomaly_flags,
        "coa_duplicate_rows": pd.DataFrame(),
        "coa_mapping_review": mapping_review,
        "financial_logic_review": financial_logic_review,
        "last_validation_report": last_validation_report,
        "external_benchmark_df": source["benchmark"],
        "fx_rate_info": {"Base": "AUD", "Target": "USD", "Rate": 0.66, "Date": "2026-06-30", "Source": "Demo data"},
        "country_indicators": pd.DataFrame([
            {"Indicator": "GDP growth %", "Year": "2025", "Value": 2.1, "Source": "Demo data"},
            {"Indicator": "Inflation %", "Year": "2026", "Value": 2.8, "Source": "Demo data"},
        ]),
        "save_run_preference": False,
    }
    for key, value in values.items():
        st.session_state[key] = value

    st.session_state["auth_mode"] = "demo"
    st.session_state["app_logged_in"] = True
    st.session_state["workspace_modules"] = [
        "Financial Statements", "KPI Dashboard", "AI CFO", "Forecasting", "Working Capital", "Benchmarking", "Board Pack"
    ]
    st.session_state["demo_data_loaded"] = True
    st.session_state["demo_tour_step"] = 0
    st.session_state["demo_tour_active"] = False
    st.session_state["demo_welcome_pending"] = True
    st.session_state["demo_read_only"] = True
    st.session_state["ai_cfo_chat_messages"] = []
    st.query_params["page"] = "dashboard"


def clear_demo_workspace(return_to_login: bool = True) -> None:
    """Clear preloaded demo state and optionally return to the login screen."""
    for key in DEMO_DATA_KEYS:
        if key in st.session_state:
            st.session_state[key] = {} if key == "company_profile" else None
    for key in ["demo_data_loaded", "demo_tour_step", "demo_tour_active", "demo_welcome_pending", "demo_read_only"]:
        st.session_state.pop(key, None)
    st.session_state["ai_cfo_chat_messages"] = []
    if return_to_login:
        st.session_state["app_logged_in"] = False
        st.session_state["auth_mode"] = None
        st.session_state["auth_view"] = "login"
        st.session_state["onboarding_step"] = 1
        st.query_params.clear()


def _demo_welcome_dialog() -> None:
    @st.dialog("Welcome to the guided demo")
    def _dialog() -> None:
        st.markdown("### A complete sample finance workspace is ready")
        st.write(
            "Northstar Manufacturing has been preloaded with six months of GL data, COA mapping, P&L, balance sheet, "
            "KPIs, budget, forecast, branch results, AR/AP ageing, benchmarks and AI observations."
        )
        st.info("Nothing in demo mode is saved to a real company workspace.")
        c1, c2 = st.columns(2)
        if c1.button("Start guided tour", type="primary", use_container_width=True, key="demo_start_tour"):
            st.session_state["demo_welcome_pending"] = False
            st.session_state["demo_tour_active"] = True
            st.session_state["demo_tour_step"] = 0
            st.query_params["page"] = TOUR_STEPS[0]["page"]
            st.rerun()
        if c2.button("Explore freely", use_container_width=True, key="demo_explore_free"):
            st.session_state["demo_welcome_pending"] = False
            st.session_state["demo_tour_active"] = False
            st.rerun()

    _dialog()


def render_demo_experience() -> None:
    """Render a prominent, working guided-demo controller and exit route."""
    if st.session_state.get("auth_mode") != "demo":
        return
    if not st.session_state.get("demo_data_loaded"):
        load_demo_workspace()

    st.markdown(
        """
        <style>
        .demo-shell {
          position: sticky; top: .35rem; z-index: 99990;
          border-radius: 22px; padding: .95rem 1rem; margin: .2rem 0 .9rem 0;
          background: linear-gradient(135deg, rgba(15,23,42,.98), rgba(30,64,175,.94), rgba(88,28,135,.92));
          border: 1px solid rgba(147,197,253,.48); box-shadow: 0 20px 52px rgba(2,6,23,.38);
          color: #eff6ff; animation: demoDrop .34s ease-out both;
          backdrop-filter: blur(16px);
        }
        .demo-topline{display:flex;align-items:center;justify-content:space-between;gap:1rem;flex-wrap:wrap}
        .demo-title{font-weight:950;color:#fff;font-size:1.02rem;letter-spacing:-.02em}
        .demo-sub{color:#dbeafe;font-size:.84rem;margin-top:.18rem}
        .demo-progress{height:7px;border-radius:999px;background:rgba(255,255,255,.15);overflow:hidden;margin-top:.7rem}
        .demo-progress > span{display:block;height:100%;border-radius:999px;background:linear-gradient(90deg,#22d3ee,#60a5fa,#c084fc);box-shadow:0 0 18px rgba(96,165,250,.55);transition:width .35s ease}
        .demo-chip{display:inline-flex;padding:.3rem .58rem;border-radius:999px;background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.20);font-weight:850;font-size:.74rem;color:#fff}
        .st-key-demo_exit_signin button, .st-key-demo_exit_quick button {
          background: linear-gradient(135deg,#dc2626,#ef4444) !important;
          color:#fff !important;border:1px solid rgba(254,202,202,.6) !important;
          box-shadow:0 12px 28px rgba(220,38,38,.32) !important;font-weight:900 !important;
        }
        .st-key-demo_start_or_resume button {background:linear-gradient(135deg,#0284c7,#4f46e5)!important;color:#fff!important;font-weight:900!important}
        @keyframes demoDrop{from{opacity:0;transform:translateY(-10px)}to{opacity:1;transform:none}}
        </style>
        """,
        unsafe_allow_html=True,
    )

    step = int(st.session_state.get("demo_tour_step", 0) or 0)
    step = max(0, min(step, len(TOUR_STEPS) - 1))
    active = bool(st.session_state.get("demo_tour_active", False))
    item = TOUR_STEPS[step]
    progress = int(((step + 1) / len(TOUR_STEPS)) * 100)
    mode_copy = f"Guided tour {step + 1}/{len(TOUR_STEPS)}" if active else "Explore freely or start the guided tour"

    st.markdown(
        f"""
        <div class="demo-shell">
          <div class="demo-topline">
            <div>
              <div class="demo-title">Demo Workspace · Northstar Manufacturing <span class="demo-chip">Sample data</span></div>
              <div class="demo-sub">{mode_copy} · <b>{item['title']}</b> — {item['copy']}</div>
            </div>
            <div class="demo-chip">Nothing is saved</div>
          </div>
          <div class="demo-progress"><span style="width:{progress}%"></span></div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if active:
        c1, c2, c3, c4 = st.columns([.9, 1.05, 1, 1.35])
        if c1.button("← Previous", use_container_width=True, disabled=step == 0, key="demo_tour_prev"):
            step -= 1
            st.session_state["demo_tour_step"] = step
            st.query_params["page"] = TOUR_STEPS[step]["page"]
            st.rerun()
        if c2.button("Next step →", type="primary", use_container_width=True, disabled=step >= len(TOUR_STEPS) - 1, key="demo_tour_next"):
            step += 1
            st.session_state["demo_tour_step"] = step
            st.query_params["page"] = TOUR_STEPS[step]["page"]
            st.rerun()
        if c3.button("Stop tour", use_container_width=True, key="demo_stop_tour"):
            st.session_state["demo_tour_active"] = False
            st.rerun()
        if c4.button("Exit Demo & Sign In", use_container_width=True, key="demo_exit_signin"):
            clear_demo_workspace(return_to_login=True)
            st.rerun()
    else:
        c1, c2, c3 = st.columns([1.2, 1, 1.3])
        if c1.button("Start guided tour", type="primary", use_container_width=True, key="demo_start_or_resume"):
            st.session_state["demo_tour_active"] = True
            st.session_state["demo_tour_step"] = 0
            st.query_params["page"] = TOUR_STEPS[0]["page"]
            st.rerun()
        if c2.button("Restart sample workspace", use_container_width=True, key="demo_reset_workspace"):
            load_demo_workspace(force=True)
            st.session_state["demo_tour_active"] = False
            st.rerun()
        if c3.button("Exit Demo & Sign In", use_container_width=True, key="demo_exit_signin"):
            clear_demo_workspace(return_to_login=True)
            st.rerun()

    # Always-visible quick exit at the top of the content flow.
    if st.session_state.get("demo_welcome_pending") and hasattr(st, "dialog"):
        _demo_welcome_dialog()

