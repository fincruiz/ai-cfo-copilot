import os
import re
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openai import OpenAI
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="AI CFO Copilot", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"]  {
    font-family: Arial, sans-serif;
}
h1 {
    font-family: Arial, sans-serif !important;
    font-size: 32px !important;
    font-weight: 700 !important;
}
h2 {
    font-family: Arial, sans-serif !important;
    font-size: 24px !important;
    font-weight: 700 !important;
}
h3 {
    font-family: Arial, sans-serif !important;
    font-size: 20px !important;
    font-weight: 700 !important;
}
div[data-testid="stDataFrame"] * {
    font-family: Arial, sans-serif !important;
    font-size: 13px !important;
}
div[data-testid="stMetric"] * {
    font-family: Arial, sans-serif !important;
}
button {
    font-family: Arial, sans-serif !important;
    font-size: 14px !important;
    font-weight: 600 !important;
}
</style>
""", unsafe_allow_html=True)

HISTORY_ROOT = Path("history")
HISTORY_ROOT.mkdir(exist_ok=True)


# ----------------------------
# Generic helpers
# ----------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def slugify_company_name(name: str) -> str:
    name = str(name).strip().lower()
    name = re.sub(r"[^a-z0-9]+", "_", name)
    name = re.sub(r"_+", "_", name).strip("_")
    return name or "unknown_company"


def style_dataframe(df: pd.DataFrame):
    return df.style.set_properties(**{
        "font-family": "Arial",
        "font-size": "13px",
        "text-align": "left"
    })


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], file_label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(
            f"{file_label} → Missing columns: {missing} | Found: {list(df.columns)}"
        )


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def show_required_columns(title, required_cols, optional_cols=None):
    st.markdown(f"**{title}**")
    req_df = pd.DataFrame({
        "Column": required_cols,
        "Required": ["Yes"] * len(required_cols),
    })

    if optional_cols:
        opt_df = pd.DataFrame({
            "Column": optional_cols,
            "Required": ["Optional"] * len(optional_cols),
        })
        display_df = pd.concat([req_df, opt_df], ignore_index=True)
    else:
        display_df = req_df

    st.dataframe(display_df, use_container_width=True, hide_index=True)


# ----------------------------
# Excel / template helpers
# ----------------------------
def format_excel_sheet(ws):
    header_fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
    header_font = Font(name="Arial", size=11, bold=True)
    body_font = Font(name="Arial", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font = body_font
            cell.alignment = Alignment(horizontal="left", vertical="center")
            cell.border = thin_border

    for col_cells in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            try:
                max_length = max(max_length, len(str(cell.value)) if cell.value is not None else 0)
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_length + 3, 35)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22


def dataframe_to_excel_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            safe_sheet = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            format_excel_sheet(writer.book[safe_sheet])
    return output.getvalue()


def make_sample_template_bytes(df: pd.DataFrame) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Template", index=False)
        ws = writer.book["Template"]
        format_excel_sheet(ws)
    return output.getvalue()


def get_sample_templates():
    templates = {}

    templates["Current GL Report"] = pd.DataFrame([
        {"Account code": "4000", "Debit": 0, "Credit": 25000, "Branch": "Sydney", "Net": -25000, "Date": "2026-03-01", "Description": "Sales invoice"},
        {"Account code": "5000", "Debit": 8000, "Credit": 0, "Branch": "Sydney", "Net": 8000, "Date": "2026-03-02", "Description": "Cost of sales"},
        {"Account code": "6100", "Debit": 3000, "Credit": 0, "Branch": "Melbourne", "Net": 3000, "Date": "2026-03-03", "Description": "Rent expense"},
    ])

    templates["COA Mapping"] = pd.DataFrame([
        {"Account code": "4000", "Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Statement": "Income Statement", "Sign Convention": "positive"},
        {"Account code": "5000", "Reporting Group": "Gross Profit", "Reporting Subgroup": "Cost of Sales", "Statement": "Income Statement", "Sign Convention": "negative"},
        {"Account code": "6100", "Reporting Group": "Operating Expenses", "Reporting Subgroup": "Rent", "Statement": "Income Statement", "Sign Convention": "positive"},
    ])

    templates["KPI Master"] = pd.DataFrame([
        {"KPI Name": "Revenue", "Formula Type": "direct", "Numerator Group": "Revenue", "Denominator Group": "", "Output Type": "value", "Display Order": 1},
        {"KPI Name": "Gross Margin %", "Formula Type": "ratio", "Numerator Group": "Gross Profit", "Denominator Group": "Revenue", "Output Type": "percent", "Display Order": 2},
        {"KPI Name": "Operating Margin %", "Formula Type": "ratio", "Numerator Group": "Operating Profit", "Denominator Group": "Revenue", "Output Type": "percent", "Display Order": 3},
    ])

    templates["Latest Previous Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 50000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 22000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 28000},
    ])

    templates["Budget Data"] = pd.DataFrame([
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Revenue", "Amount": 100000},
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Gross Profit", "Amount": 42000},
        {"Month": "2026-01", "Branch": "Melbourne", "Reporting Group": "Revenue", "Amount": 85000},
    ])

    templates["Forecast P&L"] = pd.DataFrame([
        {"Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 120000},
        {"Reporting Group": "Gross Profit", "Reporting Subgroup": "Gross Profit", "Report Value": 48000},
        {"Reporting Group": "Operating Profit", "Reporting Subgroup": "EBIT", "Report Value": 18000},
    ])

    templates["Forecast Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 65000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 28000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 37000},
    ])

    templates["AR Ageing"] = pd.DataFrame([
        {"Party Name": "Customer A", "Outstanding Amount": 12000, "Document Number": "INV001", "Document Date": "2026-02-01", "Due Date": "2026-03-01", "Branch": "Sydney", "Age Bucket": "1-30"},
        {"Party Name": "Customer B", "Outstanding Amount": 8000, "Document Number": "INV002", "Document Date": "2026-01-15", "Due Date": "2026-02-15", "Branch": "Melbourne", "Age Bucket": "31-60"},
        {"Party Name": "Customer C", "Outstanding Amount": 5000, "Document Number": "INV003", "Document Date": "2026-03-05", "Due Date": "2026-04-05", "Branch": "Sydney", "Age Bucket": "Current"},
    ])

    templates["AP Ageing"] = pd.DataFrame([
        {"Party Name": "Supplier A", "Outstanding Amount": 9000, "Document Number": "BILL001", "Document Date": "2026-02-01", "Due Date": "2026-03-01", "Branch": "Sydney", "Age Bucket": "1-30"},
        {"Party Name": "Supplier B", "Outstanding Amount": 14000, "Document Number": "BILL002", "Document Date": "2026-01-10", "Due Date": "2026-02-10", "Branch": "Melbourne", "Age Bucket": "31-60"},
        {"Party Name": "Supplier C", "Outstanding Amount": 6000, "Document Number": "BILL003", "Document Date": "2026-03-04", "Due Date": "2026-04-04", "Branch": "Sydney", "Age Bucket": "Current"},
    ])

    templates["Industry Benchmark File"] = pd.DataFrame([
        {"Metric": "Gross Margin %", "Benchmark Value": 35},
        {"Metric": "Operating Margin %", "Benchmark Value": 12},
        {"Metric": "Opex as % of Revenue", "Benchmark Value": 20},
    ])

    templates["Prior Period GL Report"] = pd.DataFrame([
        {"Account code": "4000", "Debit": 0, "Credit": 22000, "Branch": "Sydney", "Net": -22000, "Date": "2025-03-01", "Description": "Prior sales"},
        {"Account code": "5000", "Debit": 7000, "Credit": 0, "Branch": "Sydney", "Net": 7000, "Date": "2025-03-02", "Description": "Prior COS"},
        {"Account code": "6100", "Debit": 2500, "Credit": 0, "Branch": "Melbourne", "Net": 2500, "Date": "2025-03-03", "Description": "Prior rent"},
    ])

    templates["Prior Period P&L"] = pd.DataFrame([
        {"Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 22000},
        {"Reporting Group": "Gross Profit", "Reporting Subgroup": "Gross Profit", "Report Value": 15000},
        {"Reporting Group": "Operating Profit", "Reporting Subgroup": "EBIT", "Report Value": 7000},
    ])

    templates["Prior Period Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 42000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 18000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 24000},
    ])

    templates["Previous Year P&L"] = pd.DataFrame([
        {"Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 98000},
        {"Reporting Group": "Gross Profit", "Reporting Subgroup": "Gross Profit", "Report Value": 39000},
        {"Reporting Group": "Operating Profit", "Reporting Subgroup": "EBIT", "Report Value": 14000},
    ])

    templates["Prior Period KPI Pack"] = pd.DataFrame([
        {"KPI": "Revenue", "Value": 22000, "Display Value": 22000, "Output Type": "value"},
        {"KPI": "Gross Margin %", "Value": 68.18, "Display Value": "68.18%", "Output Type": "percent"},
        {"KPI": "Operating Margin %", "Value": 31.82, "Display Value": "31.82%", "Output Type": "percent"},
    ])

    return templates


# ----------------------------
# Normalizers for direct P&L / BS uploads
# ----------------------------
def normalize_uploaded_pnl(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={
        "Reporting group": "Reporting Group",
        "Reporting subgroup": "Reporting Subgroup",
        "Report value": "Report Value",
    }, inplace=True)

    validate_required_columns(df, ["Reporting Group", "Reporting Subgroup", "Report Value"], label)
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Reporting Subgroup"] = df["Reporting Subgroup"].astype(str).str.strip()
    df["Report Value"] = pd.to_numeric(df["Report Value"], errors="coerce").fillna(0)
    return df


def normalize_uploaded_bs(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={
        "Reporting group": "Reporting Group",
        "Reporting subgroup": "Reporting Subgroup",
        "Balance ": "Balance",
    }, inplace=True)

    validate_required_columns(df, ["Reporting Group", "Reporting Subgroup", "Balance"], label)
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Reporting Subgroup"] = df["Reporting Subgroup"].astype(str).str.strip()
    df["Balance"] = pd.to_numeric(df["Balance"], errors="coerce").fillna(0)
    return df


def compare_pnl_to_forecast(actual_pnl: pd.DataFrame, forecast_pnl: pd.DataFrame) -> pd.DataFrame:
    if actual_pnl is None or actual_pnl.empty or forecast_pnl is None or forecast_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Actual", "Forecast", "Variance", "Variance %"])

    actual = actual_pnl.copy().rename(columns={"Report Value": "Actual"})
    forecast = forecast_pnl.copy().rename(columns={"Report Value": "Forecast"})

    merged = actual.merge(
        forecast,
        on=["Reporting Group", "Reporting Subgroup"],
        how="outer"
    ).fillna(0)

    merged["Variance"] = merged["Actual"] - merged["Forecast"]
    merged["Variance %"] = merged.apply(
        lambda r: (r["Variance"] / r["Forecast"] * 100) if r["Forecast"] != 0 else 0.0,
        axis=1,
    )
    return merged.sort_values(["Reporting Group", "Reporting Subgroup"]).reset_index(drop=True)


def compare_pnl_to_previous_year(actual_pnl: pd.DataFrame, previous_pnl: pd.DataFrame) -> pd.DataFrame:
    if actual_pnl is None or actual_pnl.empty or previous_pnl is None or previous_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Actual", "Previous Year", "Variance", "Variance %"])

    actual = actual_pnl.copy().rename(columns={"Report Value": "Actual"})
    previous = previous_pnl.copy().rename(columns={"Report Value": "Previous Year"})

    merged = actual.merge(
        previous,
        on=["Reporting Group", "Reporting Subgroup"],
        how="outer"
    ).fillna(0)

    merged["Variance"] = merged["Actual"] - merged["Previous Year"]
    merged["Variance %"] = merged.apply(
        lambda r: (r["Variance"] / r["Previous Year"] * 100) if r["Previous Year"] != 0 else 0.0,
        axis=1,
    )
    return merged.sort_values(["Reporting Group", "Reporting Subgroup"]).reset_index(drop=True)


# ----------------------------
# History helpers
# ----------------------------
def save_run_to_history(company_profile, consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary):
    company_name = company_profile.get("Company Name", "").strip()
    if not company_name:
        return

    company_slug = slugify_company_name(company_name)
    financial_year = company_profile.get("Financial Year", "unknown_year").strip().replace(" ", "_")
    reporting_period = company_profile.get("Reporting Period", "unknown_period").strip().replace(" ", "_")

    company_folder = HISTORY_ROOT / company_slug
    run_folder = company_folder / f"{financial_year}_{reporting_period}"
    run_folder.mkdir(parents=True, exist_ok=True)

    consolidated_pnl.to_excel(run_folder / "consolidated_pnl.xlsx", index=False)

    if consolidated_bs is not None and not consolidated_bs.empty:
        consolidated_bs.to_excel(run_folder / "consolidated_bs.xlsx", index=False)

    if consolidated_kpis is not None:
        consolidated_kpis.to_excel(run_folder / "consolidated_kpis.xlsx", index=False)

    if branch_summary is not None and not branch_summary.empty:
        branch_summary.to_excel(run_folder / "branch_summary.xlsx", index=False)


def list_saved_company_runs(company_name: str):
    company_slug = slugify_company_name(company_name)
    company_folder = HISTORY_ROOT / company_slug
    if not company_folder.exists():
        return []
    return sorted([item.name for item in company_folder.iterdir() if item.is_dir()], reverse=True)


def restore_run_from_history(company_name: str, run_name: str):
    company_slug = slugify_company_name(company_name)
    run_folder = HISTORY_ROOT / company_slug / run_name
    restored = {}

    if (run_folder / "consolidated_pnl.xlsx").exists():
        restored["prior_pnl"] = pd.read_excel(run_folder / "consolidated_pnl.xlsx")
    if (run_folder / "consolidated_bs.xlsx").exists():
        restored["prior_bs"] = pd.read_excel(run_folder / "consolidated_bs.xlsx")
    if (run_folder / "consolidated_kpis.xlsx").exists():
        restored["prior_kpis"] = pd.read_excel(run_folder / "consolidated_kpis.xlsx")
    if (run_folder / "branch_summary.xlsx").exists():
        restored["prior_branch_summary"] = pd.read_excel(run_folder / "branch_summary.xlsx")

    return restored


# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped",
    "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary",
    "detected_branches", "validation_passed", "company_profile", "bs_disclaimer", "ai_commentary",
    "prior_pnl", "prior_bs", "prior_kpis", "save_run_preference", "anomaly_flags",
    "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df",
    "budget_compare", "budget_summary",
    "benchmark_df", "py_compare", "benchmark_compare", "monthly_actuals", "monthly_branch_actuals",
    "executive_summary_df", "forecast_pnl", "forecast_bs", "previous_year_pnl",
    "forecast_pnl_compare", "previous_year_pnl_compare"
]:
    if key not in st.session_state:
        st.session_state[key] = None

if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}
if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False


# ----------------------------
# Processing helpers
# ----------------------------
def build_executive_summary(current_kpis, ar_summary=None, ap_summary=None, budget_summary=None,
                            benchmark_compare=None, forecast_pnl_compare=None, previous_year_pnl_compare=None) -> pd.DataFrame:
    rows = []
    current_kpi_map = kpi_map_from_df(current_kpis)

    key_metrics = [
        "Revenue",
        "Gross Margin %",
        "Operating Margin %",
        "Opex as % of Revenue",
    ]

    for metric in key_metrics:
        current_value = safe_float(current_kpi_map.get(metric, 0))
        benchmark_value = ""
        if benchmark_compare is not None and not benchmark_compare.empty:
            match = benchmark_compare[benchmark_compare["Metric"] == metric]
            if not match.empty:
                benchmark_value = safe_float(match.iloc[0]["Benchmark Value"])

        rows.append({
            "Metric": metric,
            "Current Value": current_value,
            "Benchmark Value": benchmark_value,
            "Status": rag_status(metric, current_value, benchmark_value),
        })

    if ar_summary is not None:
        rows.append({
            "Metric": "AR Overdue %",
            "Current Value": safe_float(ar_summary["overdue_pct"]),
            "Benchmark Value": "",
            "Status": rag_status("AR Overdue %", safe_float(ar_summary["overdue_pct"])),
        })

    if ap_summary is not None:
        rows.append({
            "Metric": "AP Overdue %",
            "Current Value": safe_float(ap_summary["overdue_pct"]),
            "Benchmark Value": "",
            "Status": rag_status("AP Overdue %", safe_float(ap_summary["overdue_pct"])),
        })

    if budget_summary is not None and not budget_summary.empty and budget_summary["Budget"].sum() != 0:
        total_variance_pct = budget_summary["Variance"].sum() / budget_summary["Budget"].sum() * 100
        rows.append({
            "Metric": "Budget Variance %",
            "Current Value": total_variance_pct,
            "Benchmark Value": "",
            "Status": "Green" if total_variance_pct >= 0 else ("Amber" if total_variance_pct >= -10 else "Red"),
        })

    if forecast_pnl_compare is not None and not forecast_pnl_compare.empty:
        forecast_total = forecast_pnl_compare["Forecast"].sum()
        variance_total = forecast_pnl_compare["Variance"].sum()
        forecast_var_pct = (variance_total / forecast_total * 100) if forecast_total != 0 else 0
        rows.append({
            "Metric": "Forecast Variance %",
            "Current Value": forecast_var_pct,
            "Benchmark Value": "",
            "Status": "Green" if forecast_var_pct >= 0 else ("Amber" if forecast_var_pct >= -10 else "Red"),
        })

    if previous_year_pnl_compare is not None and not previous_year_pnl_compare.empty:
        py_total = previous_year_pnl_compare["Previous Year"].sum()
        variance_total = previous_year_pnl_compare["Variance"].sum()
        py_var_pct = (variance_total / py_total * 100) if py_total != 0 else 0
        rows.append({
            "Metric": "Previous Year Variance %",
            "Current Value": py_var_pct,
            "Benchmark Value": "",
            "Status": "Green" if py_var_pct >= 0 else ("Amber" if py_var_pct >= -10 else "Red"),
        })

    return pd.DataFrame(rows)


def detect_anomalies(consolidated_kpis, prior_kpis=None, ar_summary=None, ap_summary=None,
                     budget_summary=None, forecast_pnl_compare=None):
    flags = []
    current_kpi_map = kpi_map_from_df(consolidated_kpis)

    revenue = current_kpi_map.get("Revenue", 0)
    gross_margin = current_kpi_map.get("Gross Margin %", 0)
    operating_margin = current_kpi_map.get("Operating Margin %", 0)
    opex_ratio = current_kpi_map.get("Opex as % of Revenue", 0)

    if revenue <= 0:
        flags.append("Revenue is zero or negative.")
    if gross_margin < 20:
        flags.append(f"Gross margin is low at {gross_margin:.2f}%.")
    if operating_margin < 5:
        flags.append(f"Operating margin is weak at {operating_margin:.2f}%.")
    if opex_ratio > 40:
        flags.append(f"Operating expenses are high at {opex_ratio:.2f}% of revenue.")

    if prior_kpis is not None and not prior_kpis.empty and "KPI" in prior_kpis.columns and "Value" in prior_kpis.columns:
        prior_kpi_map = {row["KPI"]: row["Value"] for _, row in prior_kpis.iterrows()}
        prior_revenue = prior_kpi_map.get("Revenue", None)
        if prior_revenue not in (None, 0):
            revenue_change_pct = ((revenue - prior_revenue) / prior_revenue) * 100
            if revenue_change_pct < -10:
                flags.append(f"Revenue declined {revenue_change_pct:.2f}% versus prior period.")

    if ar_summary is not None and ar_summary["overdue_pct"] > 40:
        flags.append(f"AR overdue is high at {ar_summary['overdue_pct']:.2f}% of total receivables.")
    if ap_summary is not None and ap_summary["overdue_pct"] > 40:
        flags.append(f"AP overdue is high at {ap_summary['overdue_pct']:.2f}% of total payables.")

    if budget_summary is not None and not budget_summary.empty and budget_summary["Budget"].sum() != 0:
        total_variance_pct = budget_summary["Variance"].sum() / budget_summary["Budget"].sum() * 100
        if total_variance_pct < -10:
            flags.append(f"Actual performance is {total_variance_pct:.2f}% below budget.")

    if forecast_pnl_compare is not None and not forecast_pnl_compare.empty:
        forecast_total = forecast_pnl_compare["Forecast"].sum()
        variance_total = forecast_pnl_compare["Variance"].sum()
        if forecast_total != 0:
            total_variance_pct = variance_total / forecast_total * 100
            if total_variance_pct < -10:
                flags.append(f"Actual performance is {total_variance_pct:.2f}% below forecast.")

    return flags


def generate_ai_commentary(
    pnl_df,
    kpi_df,
    bs_df,
    profile,
    anomaly_flags=None,
    ar_summary=None,
    ap_summary=None,
    budget_summary=None,
    forecast_pnl_compare=None,
):
    try:
        client = OpenAI()
        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

        pnl_summary = pnl_df.to_string(index=False)[:3000] if pnl_df is not None and not pnl_df.empty else "No P&L data available."
        kpi_summary = (
            kpi_df[["KPI", "Display Value"]].to_string(index=False)[:2000]
            if kpi_df is not None and not kpi_df.empty else "No KPI data available."
        )
        bs_summary = (
            bs_df.to_string(index=False)[:2500]
            if bs_df is not None and not bs_df.empty else "No Balance Sheet data available."
        )
        anomaly_text = "\n".join(anomaly_flags) if anomaly_flags else "No anomaly flags detected."

        ar_text = "No AR ageing data available."
        if ar_summary is not None:
            ar_text = f"Total AR: {ar_summary['total']:.2f}, Overdue AR %: {ar_summary['overdue_pct']:.2f}%"

        ap_text = "No AP ageing data available."
        if ap_summary is not None:
            ap_text = f"Total AP: {ap_summary['total']:.2f}, Overdue AP %: {ap_summary['overdue_pct']:.2f}%"

        budget_text = "No budget data available."
        if budget_summary is not None and not budget_summary.empty:
            budget_text = budget_summary.to_string(index=False)[:1500]

        forecast_text = "No forecast P&L data available."
        if forecast_pnl_compare is not None and not forecast_pnl_compare.empty:
            forecast_text = forecast_pnl_compare.to_string(index=False)[:1500]

        company_name = profile.get("Company Name", "Unknown Company")
        industry = profile.get("Industry", "Unknown Industry")
        country = profile.get("Country", "Unknown Country")
        currency = profile.get("Currency", "")
        reporting_period = profile.get("Reporting Period", "")
        financial_year = profile.get("Financial Year", "")
        notes = profile.get("Business Notes", "")

        prompt = f"""
You are an experienced CFO advisor preparing concise management commentary.

Company: {company_name}
Industry: {industry}
Country: {country}
Currency: {currency}
Reporting Period: {reporting_period}
Financial Year: {financial_year}
Business Notes: {notes}

Use only the data below. Do not invent numbers.

Detected anomaly flags:
{anomaly_text}

Consolidated P&L:
{pnl_summary}

KPIs:
{kpi_summary}

Consolidated Balance Sheet:
{bs_summary}

AR Summary:
{ar_text}

AP Summary:
{ap_text}

Budget vs Actual:
{budget_text}

Forecast vs Actual:
{forecast_text}

Write in this format:
1. Executive Summary
2. Key Insights (5 bullets)
3. Risks / Watchouts (3 bullets)
4. Opportunities (3 bullets)
5. Recommended Actions (3 bullets)

Keep it practical, management-ready, and concise.
"""

        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "developer", "content": "You are a sharp CFO advisor. Be concise, practical, and numerically grounded."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.3,
        )
        return response.choices[0].message.content

    except Exception as e:
        return f"AI Commentary failed: {str(e)}"


# ----------------------------
# Main data prep
# ----------------------------
def prepare_data(gl_file, mapping_file, kpi_file=None, latest_bs_file=None):
    gl = pd.read_excel(gl_file)
    coa = pd.read_excel(mapping_file)
    kpi_master = pd.read_excel(kpi_file) if kpi_file is not None else None
    latest_bs = pd.read_excel(latest_bs_file) if latest_bs_file is not None else None

    gl, coa, kpi_master, latest_bs = standardize_key_columns(gl, coa, kpi_master, latest_bs)

    validate_required_columns(gl, ["Account code", "Debit", "Credit", "Branch"], "GL report")
    validate_required_columns(coa, ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "COA mapping")

    if kpi_master is not None:
        validate_required_columns(
            kpi_master,
            ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"],
            "KPI master",
        )

    if latest_bs is not None:
        validate_required_columns(latest_bs, ["Reporting Group", "Reporting Subgroup", "Balance"], "Latest Balance Sheet")

    gl["Account code"] = gl["Account code"].astype(str).str.strip()
    coa["Account code"] = coa["Account code"].astype(str).str.strip()
    gl["Branch"] = gl["Branch"].astype(str).str.strip()

    gl["Debit"] = pd.to_numeric(gl["Debit"], errors="coerce").fillna(0)
    gl["Credit"] = pd.to_numeric(gl["Credit"], errors="coerce").fillna(0)

    if "Net" not in gl.columns:
        gl["Net"] = gl["Debit"] - gl["Credit"]
    else:
        gl["Net"] = pd.to_numeric(gl["Net"], errors="coerce")
        gl["Net"] = gl["Net"].fillna(gl["Debit"] - gl["Credit"])

    if "Date" in gl.columns:
        gl["Date"] = pd.to_datetime(gl["Date"], errors="coerce")

    data = gl.merge(coa, on="Account code", how="left")
    unmapped = data[data["Reporting Group"].isna()].copy()
    mapped = data[data["Reporting Group"].notna()].copy()

    if "Sign Convention" not in mapped.columns:
        mapped["Sign Convention"] = "positive"

    mapped["Report Value"] = mapped.apply(apply_sign_convention_to_gl, axis=1)

    pnl_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()
    bs_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "balance sheet"].copy()

    if latest_bs is not None:
        latest_bs["Balance"] = pd.to_numeric(latest_bs["Balance"], errors="coerce").fillna(0)

    return gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped


# ----------------------------
# Header / top-level tabs
# ----------------------------
st.title("AI CFO Copilot")
st.caption("Automated branch-wise P&L, KPI packs, dashboarding, working capital, budget/forecast comparison, and AI insights")

main_tabs = st.tabs([
    "Setup",
    "Dashboard",
    "Financials",
    "Working Capital",
    "Insights",
    "Downloads",
])

tab_setup, tab_dashboard, tab_financials, tab_working_capital, tab_insights, tab_downloads = main_tabs


# ----------------------------
# SETUP TAB
# ----------------------------
with tab_setup:
    st.subheader("Setup")

    with st.expander("Company Profile", expanded=True):
        c1, c2 = st.columns(2)

        with c1:
            company_name = st.text_input("Company Name *")
            industry = st.selectbox("Industry", ["Select Industry", "Manufacturing", "Wholesale / Distribution", "Retail", "Professional Services", "Construction", "Logistics", "Hospitality", "Healthcare", "Technology", "Other"])
            country = st.selectbox("Country", ["Select Country", "Australia", "India", "United States", "United Kingdom", "Canada", "New Zealand", "Other"])
            state_region = st.text_input("State / Region")
            financial_year = st.text_input("Financial Year", placeholder="Example: FY2025 or 2024-25")

        with c2:
            currency = st.selectbox("Currency", ["Select Currency", "AUD", "INR", "USD", "GBP", "CAD", "NZD", "Other"])
            tax_identifier = st.text_input("Tax Identifier / ABN / GSTIN (Optional)")
            reporting_period = st.selectbox("Reporting Period", ["Monthly", "Quarterly", "Annual"])
            benchmark_group = st.text_input("Benchmark Group (Optional)")

        business_notes = st.text_area("Business Notes (Optional)")
        save_run_preference = st.checkbox("Save this run for future comparison", value=st.session_state["save_run_preference"])

        if st.button("Save Company Profile", use_container_width=True):
            if not company_name.strip():
                st.error("Company Name is mandatory.")
            elif industry == "Select Industry" or country == "Select Country":
                st.error("Please select at least Industry and Country.")
            else:
                st.session_state["company_profile"] = {
                    "Company Name": company_name.strip(),
                    "Industry": industry,
                    "Country": country,
                    "State / Region": state_region,
                    "Financial Year": financial_year,
                    "Currency": currency if currency != "Select Currency" else "",
                    "Tax Identifier": tax_identifier,
                    "Reporting Period": reporting_period,
                    "Benchmark Group": benchmark_group,
                    "Business Notes": business_notes,
                }
                st.session_state["save_run_preference"] = save_run_preference
                st.success("Company profile saved successfully.")

        if st.session_state["company_profile"]:
            profile_df = pd.DataFrame(st.session_state["company_profile"].items(), columns=["Field", "Value"])
            st.dataframe(style_dataframe(profile_df), use_container_width=True)

    with st.expander("Current Period Uploads", expanded=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            gl_file = st.file_uploader("Current GL Report", type=["xlsx"])
            mapping_file = st.file_uploader("COA Mapping", type=["xlsx"])
            budget_file = st.file_uploader("Budget Data (Optional)", type=["xlsx"])
        with c2:
            kpi_file = st.file_uploader("KPI Master (Optional)", type=["xlsx"])
            latest_bs_file = st.file_uploader("Latest Previous Balance Sheet (Optional)", type=["xlsx"])
            forecast_pnl_file = st.file_uploader("Forecast P&L (Optional)", type=["xlsx"])
        with c3:
            forecast_bs_file = st.file_uploader("Forecast Balance Sheet (Optional)", type=["xlsx"])
            ar_file = st.file_uploader("AR Ageing (Optional)", type=["xlsx"])
            ap_file = st.file_uploader("AP Ageing (Optional)", type=["xlsx"])
            benchmark_file = st.file_uploader("Industry Benchmark File (Optional)", type=["xlsx"])

        previous_year_pnl_file = st.file_uploader("Previous Year P&L (Optional)", type=["xlsx"])

if st.button("Validate & Load Current Files", use_container_width=True):

    st.markdown("## 🔍 Validation Summary")

    validation_errors = []
    validation_success = []

    def log_error(file, msg, df=None):
        validation_errors.append({
            "File": file,
            "Issue": str(msg),
            "Columns Found": ", ".join(df.columns) if df is not None else "Unreadable"
        })

    def log_success(file):
        validation_success.append({
            "File": file,
            "Status": "✅ Valid"
        })

    def preview(file_name, df):
        with st.expander(f"Preview: {file_name}"):
            st.dataframe(df.head(5), use_container_width=True)

    # ----------------------------
    # Profile Check
    # ----------------------------
    profile = st.session_state.get("company_profile", {})
    if not profile or not profile.get("Company Name", "").strip():
        st.error("Please save Company Profile first.")
        st.stop()

    if gl_file is None or mapping_file is None:
        st.error("GL and COA Mapping are mandatory.")
        st.stop()

    # ----------------------------
    # GL
    # ----------------------------
    try:
        df = pd.read_excel(gl_file)
        df = clean_columns(df)
        validate_required_columns(df, ["Account code", "Debit", "Credit", "Branch"], "GL Report")
        loaded_files["gl"] = df
        add_success("GL Report")
        preview("GL Report", df)
    except Exception as e:
        add_error("GL Report", e, df if "df" in locals() else None)

    # ----------------------------
    # COA
    # ----------------------------
    try:
        df = pd.read_excel(mapping_file)
        df = clean_columns(df)
        validate_required_columns(df, ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "COA Mapping")
        loaded_files["coa"] = df
        add_success("COA Mapping")
        preview("COA Mapping", df)
    except Exception as e:
        add_error("COA Mapping", e, df if "df" in locals() else None)

    # ----------------------------
    # KPI
    # ----------------------------
    if kpi_file is not None:
        try:
            df = pd.read_excel(kpi_file)
            df = clean_columns(df)
            validate_required_columns(df,
                ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"],
                "KPI Master"
            )
            loaded_files["kpi"] = df
            add_success("KPI Master")
            preview("KPI Master", df)
        except Exception as e:
            add_error("KPI Master", e, df if "df" in locals() else None)

    # ----------------------------
    # Forecast P&L
    # ----------------------------
    if forecast_pnl_file is not None:
        try:
            df = normalize_uploaded_pnl(pd.read_excel(forecast_pnl_file), "Forecast P&L")
            loaded_files["forecast_pnl"] = df
            add_success("Forecast P&L")
            preview("Forecast P&L", df)
        except Exception as e:
            add_error("Forecast P&L", e)

    # ----------------------------
    # Previous Year
    # ----------------------------
    if previous_year_pnl_file is not None:
        try:
            df = normalize_uploaded_pnl(pd.read_excel(previous_year_pnl_file), "Previous Year P&L")
            loaded_files["py"] = df
            add_success("Previous Year P&L")
            preview("Previous Year P&L", df)
        except Exception as e:
            add_error("Previous Year P&L", e)

    # ----------------------------
    # RESULT DISPLAY
    # ----------------------------
    if validation_success:
        st.success("Validated Files")
        st.dataframe(pd.DataFrame(validation_success), use_container_width=True)

    if validation_errors:
        st.error("Validation Errors Found")
        st.dataframe(pd.DataFrame(validation_errors), use_container_width=True)
        st.stop()

    # ----------------------------
    # IF ALL GOOD → PROCESS
    # ----------------------------
    try:
        gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
            gl_file, mapping_file, kpi_file, latest_bs_file
        )

        st.success("All files loaded successfully.")

        # (keep your existing session_state logic below as is)

        st.session_state["gl"] = gl
        st.session_state["coa"] = coa
        st.session_state["mapped"] = mapped
        st.session_state["pnl_mapped"] = pnl_mapped
        st.session_state["bs_mapped"] = bs_mapped
        st.session_state["unmapped"] = unmapped

    except Exception as e:
        st.error("Processing failed after validation")
        st.exception(e)

    with st.expander("Prior Period / Restore"):
        company_name_for_history = st.session_state["company_profile"].get("Company Name", "").strip()

        if not company_name_for_history:
            st.warning("Please save Company Profile first.")
        else:
            saved_runs = list_saved_company_runs(company_name_for_history)

            if saved_runs:
                selected_run = st.selectbox("Select Saved Run", saved_runs)
                if st.button("Restore Selected Run", use_container_width=True):
                    restored = restore_run_from_history(company_name_for_history, selected_run)
                    st.session_state["prior_pnl"] = restored.get("prior_pnl")
                    st.session_state["prior_bs"] = restored.get("prior_bs")
                    st.session_state["prior_kpis"] = restored.get("prior_kpis")
                    st.success(f"Restored: {selected_run}")
            else:
                st.info("No saved history found for this company.")

        c1, c2 = st.columns(2)
        with c1:
            prior_gl_file = st.file_uploader("Prior Period GL Report (Optional)", type=["xlsx"])
            prior_pnl_file = st.file_uploader("Prior Period P&L (Optional)", type=["xlsx"])
        with c2:
            prior_bs_file = st.file_uploader("Prior Period Balance Sheet (Optional)", type=["xlsx"])
            prior_kpi_file = st.file_uploader("Prior Period KPI Pack (Optional)", type=["xlsx"])

        if st.button("Load Prior Period Inputs", use_container_width=True):
            try:
                loaded_any = False

                if prior_gl_file is not None:
                    st.error("Prior Period GL parsing has been simplified out in this version. Use Prior Period P&L / BS / KPI Pack direct uploads.")
                else:
                    if prior_pnl_file is not None:
                        st.session_state["prior_pnl"] = normalize_uploaded_pnl(pd.read_excel(prior_pnl_file), "Prior Period P&L")
                        loaded_any = True
                    if prior_bs_file is not None:
                        st.session_state["prior_bs"] = normalize_uploaded_bs(pd.read_excel(prior_bs_file), "Prior Period Balance Sheet")
                        loaded_any = True
                    if prior_kpi_file is not None:
                        pk = clean_columns(pd.read_excel(prior_kpi_file))
                        pk.rename(columns={"Kpi": "KPI", "Display value": "Display Value"}, inplace=True)
                        st.session_state["prior_kpis"] = pk
                        loaded_any = True

                if loaded_any:
                    st.success("Prior period data loaded successfully.")
                else:
                    st.info("No prior period file uploaded.")
            except Exception as e:
                st.error(f"Error loading prior period data: {e}")

    with st.expander("Validation Summary"):
        if st.session_state["gl"] is None:
            st.info("No validated files loaded yet.")
        else:
            gl = st.session_state["gl"]
            mapped = st.session_state["mapped"]
            unmapped = st.session_state["unmapped"]
            detected_branches = st.session_state["detected_branches"]

            m1, m2, m3, m4 = st.columns(4)
            m1.metric("GL Rows", len(gl))
            m2.metric("Mapped Rows", len(mapped))
            m3.metric("Unmapped Rows", len(unmapped))
            m4.metric("Branches Found", len(detected_branches))

            if not unmapped.empty:
                cols_to_show = [c for c in ["Account code", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
                st.dataframe(style_dataframe(unmapped[cols_to_show]), use_container_width=True)

    with st.expander("Required Columns Guide"):
        g1, g2 = st.columns(2)

        with g1:
            show_required_columns("Current GL Report", ["Account code", "Debit", "Credit", "Branch"], ["Net", "Date", "Description"])
            show_required_columns("COA Mapping", ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], ["Sign Convention"])
            show_required_columns("KPI Master", ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], [])
            show_required_columns("Latest Previous Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Budget Data", ["Month", "Branch", "Reporting Group", "Amount"], [])
            show_required_columns("Forecast P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], [])

        with g2:
            show_required_columns("Forecast Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Previous Year P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], [])
            show_required_columns("AR Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("AP Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("Industry Benchmark File", ["Metric", "Benchmark Value"], [])
            show_required_columns("Prior Period KPI Pack", ["KPI", "Value"], ["Display Value", "Output Type"])

    with st.expander("Download Sample Templates"):
        st.info("Download a template, replace the sample rows with your own data, and upload the same file back into the app.")

        templates = get_sample_templates()
        for name, df in templates.items():
            template_bytes = make_sample_template_bytes(df)
            st.download_button(
                label=f"Download {name} Template",
                data=template_bytes,
                file_name=f"{name.lower().replace(' ', '_')}_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key=f"tpl_{name}"
            )


# ----------------------------
# DASHBOARD TAB
# ----------------------------
with tab_dashboard:
    st.subheader("Dashboard")

    if st.session_state["mapped"] is None:
        st.warning("Please complete setup and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped GL rows before using dashboard.")
    else:
        exec_df = st.session_state["executive_summary_df"]

        if exec_df is not None and not exec_df.empty:
            c1, c2, c3 = st.columns(3)
            c1.metric("Green", int((exec_df["Status"] == "Green").sum()))
            c2.metric("Amber", int((exec_df["Status"] == "Amber").sum()))
            c3.metric("Red", int((exec_df["Status"] == "Red").sum()))

        current_kpi_map = kpi_map_from_df(st.session_state["consolidated_kpis"])
        ar_summary = st.session_state.get("ar_summary")
        ap_summary = st.session_state.get("ap_summary")

        revenue = current_kpi_map.get("Revenue", 0)
        gp = current_kpi_map.get("Gross Profit", 0)
        gm = current_kpi_map.get("Gross Margin %", 0)
        op = current_kpi_map.get("Operating Profit", 0)
        opm = current_kpi_map.get("Operating Margin %", 0)
        opex_pct = current_kpi_map.get("Opex as % of Revenue", 0)

        st.markdown("### Core KPI Snapshot")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Revenue", f"{revenue:,.2f}")
        k2.metric("Gross Profit", f"{gp:,.2f}")
        k3.metric("Gross Margin %", f"{gm:.2f}%")
        k4.metric("Operating Profit", f"{op:,.2f}")
        k5.metric("Operating Margin %", f"{opm:.2f}%")

        k6, k7, k8, k9, k10 = st.columns(5)
        k6.metric("Opex %", f"{opex_pct:.2f}%")
        k7.metric("Total AR", f"{ar_summary['total']:,.2f}" if ar_summary else "0.00")
        k8.metric("AR Overdue %", f"{ar_summary['overdue_pct']:.2f}%" if ar_summary else "0.00%")
        k9.metric("Total AP", f"{ap_summary['total']:,.2f}" if ap_summary else "0.00")
        k10.metric("AP Overdue %", f"{ap_summary['overdue_pct']:.2f}%" if ap_summary else "0.00%")

        st.markdown("### Key Charts")

        if st.session_state["budget_summary"] is not None and not st.session_state["budget_summary"].empty:
            st.markdown("**Budget vs Actual**")
            st.bar_chart(st.session_state["budget_summary"].set_index("Reporting Group")[["Actual", "Budget"]])

        if st.session_state["forecast_pnl_compare"] is not None and not st.session_state["forecast_pnl_compare"].empty:
            st.markdown("**Actual vs Forecast P&L**")
            fc_chart = (
                st.session_state["forecast_pnl_compare"]
                .groupby("Reporting Group")[["Actual", "Forecast"]]
                .sum()
            )
            st.bar_chart(fc_chart)

        if st.session_state["previous_year_pnl_compare"] is not None and not st.session_state["previous_year_pnl_compare"].empty:
            st.markdown("**Actual vs Previous Year P&L**")
            py_chart = (
                st.session_state["previous_year_pnl_compare"]
                .groupby("Reporting Group")[["Actual", "Previous Year"]]
                .sum()
            )
            st.bar_chart(py_chart)

        if st.session_state["benchmark_compare"] is not None and not st.session_state["benchmark_compare"].empty:
            st.markdown("**Industry Benchmark Comparison**")
            bench_chart = st.session_state["benchmark_compare"].copy().set_index("Metric")[["Current Value", "Benchmark Value"]]
            st.bar_chart(bench_chart)

        branch_rows = []
        for branch, reports in st.session_state["branch_outputs"].items():
            branch_kpi_map = kpi_map_from_df(reports["kpis"])
            branch_rows.append({
                "Branch": branch,
                "Revenue": branch_kpi_map.get("Revenue", 0),
                "Gross Margin %": branch_kpi_map.get("Gross Margin %", 0),
                "Operating Margin %": branch_kpi_map.get("Operating Margin %", 0),
            })
        branch_df = pd.DataFrame(branch_rows)

        if not branch_df.empty:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Revenue by Branch**")
                st.bar_chart(branch_df.set_index("Branch")[["Revenue"]])
            with c2:
                st.markdown("**Operating Margin % by Branch**")
                st.bar_chart(branch_df.set_index("Branch")[["Operating Margin %"]])


# ----------------------------
# FINANCIALS TAB
# ----------------------------
with tab_financials:
    st.subheader("Financials")

    sub_pnl, sub_bs, sub_kpi, sub_trends, sub_variance = st.tabs(
        ["P&L", "Balance Sheet", "KPIs", "Trends", "Variance"]
    )

    with sub_pnl:
        if st.session_state["consolidated_pnl"] is None:
            st.info("No P&L available yet.")
        else:
            st.markdown("### Consolidated P&L")
            st.dataframe(style_dataframe(st.session_state["consolidated_pnl"]), use_container_width=True)

            if st.session_state["branch_outputs"]:
                st.markdown("### Branch P&L")
                for branch, reports in st.session_state["branch_outputs"].items():
                    with st.expander(branch):
                        st.dataframe(style_dataframe(reports["pnl"]), use_container_width=True)

            if st.session_state["forecast_pnl"] is not None:
                st.markdown("### Forecast P&L")
                st.dataframe(style_dataframe(st.session_state["forecast_pnl"]), use_container_width=True)

            if st.session_state["previous_year_pnl"] is not None:
                st.markdown("### Previous Year P&L")
                st.dataframe(style_dataframe(st.session_state["previous_year_pnl"]), use_container_width=True)

    with sub_bs:
        if st.session_state["consolidated_bs"] is None or st.session_state["consolidated_bs"].empty:
            st.info("No Balance Sheet available yet.")
        else:
            if st.session_state["bs_disclaimer"]:
                st.warning(st.session_state["bs_disclaimer"])
            st.dataframe(style_dataframe(st.session_state["consolidated_bs"]), use_container_width=True)

        if st.session_state["forecast_bs"] is not None:
            st.markdown("### Forecast Balance Sheet")
            st.dataframe(style_dataframe(st.session_state["forecast_bs"]), use_container_width=True)

    with sub_kpi:
        if st.session_state["consolidated_kpis"] is None:
            st.info("No KPI master uploaded.")
        else:
            st.markdown("### Consolidated KPIs")
            st.dataframe(style_dataframe(st.session_state["consolidated_kpis"][["KPI", "Display Value"]]), use_container_width=True)

            if st.session_state["branch_summary"] is not None and not st.session_state["branch_summary"].empty:
                st.markdown("### Branch KPI Summary")
                st.dataframe(style_dataframe(st.session_state["branch_summary"]), use_container_width=True)

    with sub_trends:
        monthly_actuals = st.session_state.get("monthly_actuals")
        monthly_branch_actuals = st.session_state.get("monthly_branch_actuals")

        if monthly_actuals is None or monthly_actuals.empty:
            st.info("No monthly trend data available. Upload GL with a valid Date column.")
        else:
            revenue_monthly = monthly_actuals[monthly_actuals["Reporting Group"].astype(str).str.strip().str.lower() == "revenue"].copy()
            gp_monthly = monthly_actuals[monthly_actuals["Reporting Group"].astype(str).str.strip().str.lower() == "gross profit"].copy()
            op_monthly = monthly_actuals[monthly_actuals["Reporting Group"].astype(str).str.strip().str.lower() == "operating profit"].copy()

            if not revenue_monthly.empty:
                st.markdown("### Revenue Trend")
                st.line_chart(revenue_monthly.set_index("Month")[["Amount"]])

            if not gp_monthly.empty:
                st.markdown("### Gross Profit Trend")
                st.line_chart(gp_monthly.set_index("Month")[["Amount"]])

            if not op_monthly.empty:
                st.markdown("### Operating Profit Trend")
                st.line_chart(op_monthly.set_index("Month")[["Amount"]])

            if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
                st.markdown("### Branch Revenue Trend")
                pivot_branch = monthly_branch_actuals.pivot(index="Month", columns="Branch", values="Amount").fillna(0)
                st.line_chart(pivot_branch)

            st.markdown("### Monthly Trend Data")
            st.dataframe(style_dataframe(monthly_actuals), use_container_width=True)

    with sub_variance:
        if st.session_state["budget_compare"] is not None and not st.session_state["budget_compare"].empty:
            st.markdown("### Budget vs Actual")
            st.dataframe(style_dataframe(st.session_state["budget_summary"]), use_container_width=True)
            st.dataframe(style_dataframe(st.session_state["budget_compare"]), use_container_width=True)
        else:
            st.info("No budget data uploaded.")

        if st.session_state["forecast_pnl_compare"] is not None and not st.session_state["forecast_pnl_compare"].empty:
            st.markdown("### Actual vs Forecast P&L")
            st.dataframe(style_dataframe(st.session_state["forecast_pnl_compare"]), use_container_width=True)
        else:
            st.info("No forecast P&L uploaded.")

        if st.session_state["previous_year_pnl_compare"] is not None and not st.session_state["previous_year_pnl_compare"].empty:
            st.markdown("### Actual vs Previous Year P&L")
            st.dataframe(style_dataframe(st.session_state["previous_year_pnl_compare"]), use_container_width=True)
        else:
            st.info("No previous year P&L uploaded.")

        if st.session_state["benchmark_compare"] is not None and not st.session_state["benchmark_compare"].empty:
            st.markdown("### Benchmark Comparison")
            st.dataframe(style_dataframe(st.session_state["benchmark_compare"]), use_container_width=True)


# ----------------------------
# WORKING CAPITAL TAB
# ----------------------------
with tab_working_capital:
    st.subheader("Working Capital")

    wc_ar, wc_ap = st.tabs(["AR", "AP"])

    with wc_ar:
        if st.session_state["ar_summary"] is None:
            st.info("Upload AR file to view AR ageing.")
        else:
            ar = st.session_state["ar_summary"]
            x1, x2, x3 = st.columns(3)
            x1.metric("Total AR", f"{ar['total']:,.2f}")
            x2.metric("Overdue AR", f"{ar['overdue']:,.2f}")
            x3.metric("Overdue AR %", f"{ar['overdue_pct']:.2f}%")
            st.bar_chart(ar["by_bucket"].set_index("Age Bucket")[["Outstanding Amount"]])
            st.dataframe(style_dataframe(ar["by_bucket"]), use_container_width=True)
            st.dataframe(style_dataframe(ar["by_branch"]), use_container_width=True)
            st.dataframe(style_dataframe(ar["top_parties"]), use_container_width=True)

    with wc_ap:
        if st.session_state["ap_summary"] is None:
            st.info("Upload AP file to view AP ageing.")
        else:
            ap = st.session_state["ap_summary"]
            y1, y2, y3 = st.columns(3)
            y1.metric("Total AP", f"{ap['total']:,.2f}")
            y2.metric("Overdue AP", f"{ap['overdue']:,.2f}")
            y3.metric("Overdue AP %", f"{ap['overdue_pct']:.2f}%")
            st.bar_chart(ap["by_bucket"].set_index("Age Bucket")[["Outstanding Amount"]])
            st.dataframe(style_dataframe(ap["by_bucket"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["by_branch"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["top_parties"]), use_container_width=True)


# ----------------------------
# INSIGHTS TAB
# ----------------------------
with tab_insights:
    st.subheader("Insights")

    insight_anom, insight_ai = st.tabs(["Anomalies", "AI Commentary"])

    with insight_anom:
        flags = st.session_state.get("anomaly_flags", [])
        if flags:
            for flag in flags:
                st.warning(flag)
        else:
            st.success("No major anomalies detected based on current rules.")

    with insight_ai:
        if st.session_state["mapped"] is None:
            st.warning("Please upload and validate data first.")
        elif not st.session_state["validation_passed"]:
            st.error("Resolve unmapped accounts before generating AI insights.")
        else:
            if st.button("Generate AI Insights", use_container_width=True):
                with st.spinner("Analyzing financials..."):
                    st.session_state["ai_commentary"] = generate_ai_commentary(
                        st.session_state["consolidated_pnl"],
                        st.session_state["consolidated_kpis"],
                        st.session_state["consolidated_bs"],
                        st.session_state["company_profile"],
                        anomaly_flags=st.session_state.get("anomaly_flags", []),
                        ar_summary=st.session_state.get("ar_summary"),
                        ap_summary=st.session_state.get("ap_summary"),
                        budget_summary=st.session_state.get("budget_summary"),
                        forecast_pnl_compare=st.session_state.get("forecast_pnl_compare"),
                    )

            if st.session_state["ai_commentary"]:
                st.write(st.session_state["ai_commentary"])


# ----------------------------
# DOWNLOADS TAB
# ----------------------------
with tab_downloads:
    st.subheader("Downloads")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped GL rows before downloading reports.")
    else:
        full_pack_bytes = create_excel_pack(
            consolidated_pnl=st.session_state["consolidated_pnl"],
            consolidated_bs=st.session_state["consolidated_bs"],
            consolidated_kpis=st.session_state["consolidated_kpis"],
            branch_summary=st.session_state["branch_summary"],
            branch_outputs=st.session_state["branch_outputs"],
            unmapped=st.session_state["unmapped"],
            executive_summary=st.session_state["executive_summary_df"],
            monthly_actuals=st.session_state["monthly_actuals"],
            monthly_branch_actuals=st.session_state["monthly_branch_actuals"],
            ar_df=st.session_state["ar_df"],
            ap_df=st.session_state["ap_df"],
            budget_compare=st.session_state["budget_compare"],
            forecast_compare=st.session_state["forecast_pnl_compare"],
            py_compare=st.session_state["previous_year_pnl_compare"],
            benchmark_compare=st.session_state["benchmark_compare"],
        )

        st.download_button(
            label="Download Full Management Pack",
            data=full_pack_bytes,
            file_name="full_management_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        if st.session_state["unmapped"] is not None and not st.session_state["unmapped"].empty:
            unmapped_csv = st.session_state["unmapped"].to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download Unmapped GL",
                data=unmapped_csv,
                file_name="unmapped_gl.csv",
                mime="text/csv",
                use_container_width=True,
            )
