from core.validation import *
from core.formatting import *
import os
import re
import json
from io import BytesIO
from pathlib import Path
from urllib.request import Request, urlopen
from urllib.parse import urlencode

import pandas as pd
import streamlit as st
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="AI CFO Copilot", layout="wide")

st.markdown("""
<style>
html, body, [class*="css"] {font-family: Arial, sans-serif;}
h1 {font-family: Arial, sans-serif !important; font-size: 34px !important; font-weight: 700 !important;}
h2 {font-family: Arial, sans-serif !important; font-size: 24px !important; font-weight: 700 !important;}
h3 {font-family: Arial, sans-serif !important; font-size: 20px !important; font-weight: 700 !important;}
div[data-testid="stDataFrame"] * {font-family: Arial, sans-serif !important; font-size: 13px !important;}
div[data-testid="stMetric"] * {font-family: Arial, sans-serif !important;}
button {font-family: Arial, sans-serif !important; font-size: 14px !important; font-weight: 600 !important;}
</style>
""", unsafe_allow_html=True)

# Floating animated AI CFO launcher
st.markdown("""
<style>
.floating-ai-cfo-wrap {
    position: fixed;
    right: 26px;
    bottom: 28px;
    z-index: 999999;
    display: flex;
    flex-direction: column;
    align-items: flex-end;
    gap: 8px;
}
.floating-ai-cfo-label {
    background: rgba(17, 24, 39, 0.92);
    color: #ffffff;
    padding: 8px 12px;
    border-radius: 999px;
    font-size: 13px;
    box-shadow: 0 12px 30px rgba(0,0,0,0.18);
    animation: aiLabelFloat 3.2s ease-in-out infinite;
}
.floating-ai-cfo-button {
    width: 72px;
    height: 72px;
    border-radius: 50%;
    background: linear-gradient(135deg, #0f766e, #2563eb, #7c3aed);
    color: #fff !important;
    text-decoration: none !important;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 34px;
    box-shadow: 0 18px 45px rgba(37, 99, 235, 0.38);
    animation: aiFloat 2.8s ease-in-out infinite, aiPulse 1.8s ease-in-out infinite;
    border: 2px solid rgba(255,255,255,0.9);
}
.floating-ai-cfo-button:hover {
    transform: scale(1.08);
    box-shadow: 0 20px 55px rgba(124, 58, 237, 0.45);
}
.floating-ai-cfo-dot {
    position: absolute;
    right: 8px;
    bottom: 8px;
    width: 16px;
    height: 16px;
    background: #22c55e;
    border-radius: 50%;
    border: 2px solid white;
    animation: aiDot 1.4s ease-in-out infinite;
}
@keyframes aiFloat {
    0%, 100% { transform: translateY(0px) rotate(0deg); }
    50% { transform: translateY(-12px) rotate(2deg); }
}
@keyframes aiPulse {
    0% { box-shadow: 0 0 0 0 rgba(37,99,235,0.45), 0 18px 45px rgba(37,99,235,0.38); }
    70% { box-shadow: 0 0 0 18px rgba(37,99,235,0), 0 18px 45px rgba(37,99,235,0.38); }
    100% { box-shadow: 0 0 0 0 rgba(37,99,235,0), 0 18px 45px rgba(37,99,235,0.38); }
}
@keyframes aiLabelFloat {
    0%, 100% { transform: translateY(0px); opacity: 0.95; }
    50% { transform: translateY(-6px); opacity: 1; }
}
@keyframes aiDot {
    0%, 100% { transform: scale(1); opacity: 1; }
    50% { transform: scale(1.35); opacity: 0.75; }
}
@media (max-width: 768px) {
    .floating-ai-cfo-wrap { right: 16px; bottom: 18px; }
    .floating-ai-cfo-label { display: none; }
    .floating-ai-cfo-button { width: 62px; height: 62px; font-size: 29px; }
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
    """Consistent table styling with numeric columns shown to 2 decimal places."""
    if df is None:
        return pd.DataFrame().style

    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    fmt = {col: "{:,.2f}" for col in numeric_cols}

    return (
        df.style
        .format(fmt)
        .set_properties(**{
            "font-family": "Arial",
            "font-size": "13px",
            "text-align": "left",
        })
    )


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], file_label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"{file_label} → Missing columns: {missing} | Found columns: {list(df.columns)}")


def safe_float(value, default=0.0):
    try:
        if pd.isna(value):
            return default
        return float(value)
    except Exception:
        return default


def get_report_period_label(profile: dict) -> str:
    """Human-readable period label used across reports and downloads."""
    profile = profile or {}
    label = str(profile.get("Report Period", "") or "").strip()
    fy = str(profile.get("Financial Year", "") or "").strip()
    period_type = str(profile.get("Reporting Period", "") or "").strip()

    if label and fy:
        return f"{label} | {fy}"
    if label:
        return label
    if fy and period_type:
        return f"{period_type} | {fy}"
    if fy:
        return fy
    return "Period not set"


def get_period_dates(profile: dict):
    """Return selected period start/end as pandas timestamps, or (None, None)."""
    profile = profile or {}
    start_raw = profile.get("Period Start Date")
    end_raw = profile.get("Period End Date")

    start = pd.to_datetime(start_raw, errors="coerce") if start_raw not in [None, ""] else pd.NaT
    end = pd.to_datetime(end_raw, errors="coerce") if end_raw not in [None, ""] else pd.NaT

    return (None if pd.isna(start) else start.normalize(), None if pd.isna(end) else end.normalize())


def validate_gl_dates_against_profile(gl_df: pd.DataFrame, profile: dict) -> list[dict]:
    """Return warning/recommendation items if GL dates are outside selected report period."""
    issues = []
    if gl_df is None or gl_df.empty or "Date" not in gl_df.columns:
        issues.append({
            "Area": "Current GL Report",
            "Issue": "Date column not provided or not readable in GL.",
            "Recommendation": "Add Date to the GL if you want period validation and monthly trend reporting."
        })
        return issues

    start, end = get_period_dates(profile)
    if start is None or end is None:
        issues.append({
            "Area": "Company Profile",
            "Issue": "Period Start Date and/or Period End Date not set.",
            "Recommendation": "Set the reporting period dates on Home so the app can validate whether GL rows belong to the selected period."
        })
        return issues

    dates = pd.to_datetime(gl_df["Date"], errors="coerce")
    valid_dates = dates.dropna()
    if valid_dates.empty:
        issues.append({
            "Area": "Current GL Report",
            "Issue": "GL Date column is present but dates could not be read.",
            "Recommendation": "Use a standard Excel date format such as 2026-04-30."
        })
        return issues

    outside_count = int(((valid_dates < start) | (valid_dates > end)).sum())
    if outside_count > 0:
        issues.append({
            "Area": "Current GL Report",
            "Issue": f"{outside_count} GL row(s) have dates outside the selected reporting period {start.date()} to {end.date()}.",
            "Recommendation": "Check whether the uploaded GL is for the correct month/period, or update the Home reporting period dates."
        })
    return issues


# ----------------------------
# External FX / benchmark helpers
# ----------------------------
def get_country_code(country: str) -> str:
    mapping = {
        "Australia": "AUS", "India": "IND", "United States": "USA", "United Kingdom": "GBR",
        "Canada": "CAN", "New Zealand": "NZL",
    }
    return mapping.get(str(country).strip(), "")


def currency_for_country(country: str) -> str:
    mapping = {
        "Australia": "AUD", "India": "INR", "United States": "USD", "United Kingdom": "GBP",
        "Canada": "CAD", "New Zealand": "NZD",
    }
    return mapping.get(str(country).strip(), "USD")


def fetch_json_url(url: str, timeout: int = 12):
    req = Request(url, headers={"User-Agent": "AI-CFO-Copilot/1.0"})
    with urlopen(req, timeout=timeout) as response:
        return json.loads(response.read().decode("utf-8"))


def fetch_fx_rate(base_currency: str, target_currency: str, date_value: str = "latest") -> dict:
    base_currency = str(base_currency).upper().strip()
    target_currency = str(target_currency).upper().strip()
    date_value = str(date_value or "latest").strip()
    if not base_currency or not target_currency:
        raise ValueError("Base currency and target currency are required.")
    if base_currency == target_currency:
        return {"Base": base_currency, "Target": target_currency, "Rate": 1.0, "Date": date_value, "Source": "Same currency"}
    path_date = "latest" if date_value.lower() in ["", "latest"] else date_value
    url = f"https://api.frankfurter.app/{path_date}?{urlencode({'from': base_currency, 'to': target_currency})}"
    data = fetch_json_url(url)
    rate = data.get("rates", {}).get(target_currency)
    if rate is None:
        raise ValueError(f"FX rate not returned for {base_currency} to {target_currency}.")
    return {"Base": base_currency, "Target": target_currency, "Rate": float(rate), "Date": data.get("date", date_value), "Source": "Frankfurter / ECB"}


def fetch_world_bank_indicator(country_code: str, indicator_code: str, indicator_name: str) -> dict:
    if not country_code:
        raise ValueError("Country code is missing.")
    url = f"https://api.worldbank.org/v2/country/{country_code}/indicator/{indicator_code}?format=json&per_page=8"
    data = fetch_json_url(url)
    rows = data[1] if isinstance(data, list) and len(data) > 1 else []
    for row in rows:
        value = row.get("value")
        if value is not None:
            return {"Indicator": indicator_name, "Code": indicator_code, "Year": row.get("date"), "Value": float(value), "Source": "World Bank"}
    return {"Indicator": indicator_name, "Code": indicator_code, "Year": "N/A", "Value": None, "Source": "World Bank"}


def fetch_country_indicators(country: str) -> pd.DataFrame:
    country_code = get_country_code(country)
    indicators = [
        ("NY.GDP.MKTP.KD.ZG", "GDP growth %"),
        ("FP.CPI.TOTL.ZG", "Inflation %"),
        ("SL.UEM.TOTL.ZS", "Unemployment %"),
        ("NE.EXP.GNFS.ZS", "Exports % of GDP"),
        ("NE.IMP.GNFS.ZS", "Imports % of GDP"),
        ("NV.IND.TOTL.ZS", "Industry value added % of GDP"),
    ]
    rows = [fetch_world_bank_indicator(country_code, code, name) for code, name in indicators]
    return pd.DataFrame(rows)


def get_builtin_industry_benchmarks(industry: str, country: str) -> pd.DataFrame:
    """Starter benchmark set. Users can override with uploaded benchmarks.
    These are broad placeholders for app testing, not official industry benchmarks.
    """
    base = {
        "Manufacturing": {"Gross Margin %": 30, "Operating Margin %": 10, "Opex as % of Revenue": 20, "AR Overdue %": 25, "AP Overdue %": 25},
        "Wholesale / Distribution": {"Gross Margin %": 22, "Operating Margin %": 6, "Opex as % of Revenue": 16, "AR Overdue %": 30, "AP Overdue %": 30},
        "Retail": {"Gross Margin %": 35, "Operating Margin %": 8, "Opex as % of Revenue": 28, "AR Overdue %": 10, "AP Overdue %": 25},
        "Professional Services": {"Gross Margin %": 55, "Operating Margin %": 18, "Opex as % of Revenue": 35, "AR Overdue %": 30, "AP Overdue %": 20},
        "Construction": {"Gross Margin %": 20, "Operating Margin %": 6, "Opex as % of Revenue": 14, "AR Overdue %": 35, "AP Overdue %": 35},
        "Logistics": {"Gross Margin %": 28, "Operating Margin %": 8, "Opex as % of Revenue": 22, "AR Overdue %": 30, "AP Overdue %": 25},
        "Hospitality": {"Gross Margin %": 60, "Operating Margin %": 10, "Opex as % of Revenue": 45, "AR Overdue %": 10, "AP Overdue %": 20},
        "Healthcare": {"Gross Margin %": 40, "Operating Margin %": 12, "Opex as % of Revenue": 30, "AR Overdue %": 25, "AP Overdue %": 20},
        "Technology": {"Gross Margin %": 65, "Operating Margin %": 20, "Opex as % of Revenue": 42, "AR Overdue %": 25, "AP Overdue %": 15},
        "Other": {"Gross Margin %": 30, "Operating Margin %": 10, "Opex as % of Revenue": 25, "AR Overdue %": 25, "AP Overdue %": 25},
    }
    values = base.get(industry, base["Other"])
    return pd.DataFrame([{"Metric": k, "Benchmark Value": v, "Country": country, "Industry": industry, "Source": "Starter benchmark - user should verify/replace"} for k, v in values.items()])


def merge_benchmark_sources(uploaded_df: pd.DataFrame | None, external_df: pd.DataFrame | None) -> pd.DataFrame | None:
    frames = []
    if uploaded_df is not None and not uploaded_df.empty:
        frames.append(uploaded_df[["Metric", "Benchmark Value"]].copy())
    if external_df is not None and not external_df.empty:
        frames.append(external_df[["Metric", "Benchmark Value"]].copy())
    if not frames:
        return None
    merged = pd.concat(frames, ignore_index=True)
    merged = merged.drop_duplicates(subset=["Metric"], keep="first")
    return merged

def show_required_columns(title, required_cols, optional_cols=None):
    st.markdown(f"**{title}**")
    req_df = pd.DataFrame({"Column": required_cols, "Required": ["Yes"] * len(required_cols)})
    if optional_cols:
        opt_df = pd.DataFrame({"Column": optional_cols, "Required": ["Optional"] * len(optional_cols)})
        display_df = pd.concat([req_df, opt_df], ignore_index=True)
    else:
        display_df = req_df
    st.dataframe(display_df, use_container_width=True, hide_index=True)


def calculate_validation_score(critical_count: int, warning_count: int, recommendation_count: int) -> int:
    score = 100 - (critical_count * 35) - (warning_count * 8) - (recommendation_count * 3)
    return max(0, min(100, score))


def render_validation_centre(critical_items=None, warning_items=None, recommendation_items=None, info_items=None, previews=None, block_processing=False):
    """Show upload validation results in a popup-style Validation Centre."""
    critical_items = critical_items or []
    warning_items = warning_items or []
    recommendation_items = recommendation_items or []
    info_items = info_items or []
    previews = previews or {}

    score = calculate_validation_score(len(critical_items), len(warning_items), len(recommendation_items))

    def _content():
        st.markdown("### Data Validation Centre")
        s1, s2, s3, s4 = st.columns(4)
        s1.metric("Readiness Score", f"{score}/100")
        s2.metric("Critical Errors", len(critical_items))
        s3.metric("Warnings", len(warning_items))
        s4.metric("Recommendations", len(recommendation_items))

        if not critical_items and not warning_items and not recommendation_items:
            st.success("No validation errors and no recommendations. Data is ready to generate reports.")
        elif critical_items:
            st.error("Critical errors found. Please fix these before reports can be generated.")
        else:
            st.warning("Data can be processed, but review the warnings/recommendations below.")

        if critical_items:
            st.markdown("#### Critical Errors")
            st.dataframe(pd.DataFrame(critical_items), use_container_width=True, hide_index=True)
        if warning_items:
            st.markdown("#### Warnings")
            st.dataframe(pd.DataFrame(warning_items), use_container_width=True, hide_index=True)
        if recommendation_items:
            st.markdown("#### Recommendations")
            st.dataframe(pd.DataFrame(recommendation_items), use_container_width=True, hide_index=True)
        if info_items:
            st.markdown("#### Information")
            st.dataframe(pd.DataFrame(info_items), use_container_width=True, hide_index=True)

        issue_frames = []
        if critical_items:
            issue_frames.append(pd.DataFrame(critical_items).assign(Severity="Critical"))
        if warning_items:
            issue_frames.append(pd.DataFrame(warning_items).assign(Severity="Warning"))
        if recommendation_items:
            issue_frames.append(pd.DataFrame(recommendation_items).assign(Severity="Recommendation"))
        if info_items:
            issue_frames.append(pd.DataFrame(info_items).assign(Severity="Info"))
        if issue_frames:
            issue_df = pd.concat(issue_frames, ignore_index=True)
            st.download_button(
                "Download Validation Review",
                data=dataframe_to_excel_bytes({"Validation Review": issue_df}),
                file_name="validation_review.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_validation_review_popup",
            )

        if previews:
            st.markdown("#### File Previews")
            for name, df in previews.items():
                with st.expander(f"Preview: {name}"):
                    st.dataframe(df.head(5), use_container_width=True)

        if block_processing:
            st.caption("Reports are blocked until critical errors are fixed.")
        else:
            st.caption("You can proceed. Recommendations do not change your mapping automatically.")

    if hasattr(st, "dialog"):
        @st.dialog("Validation Centre")
        def _dialog():
            _content()
        _dialog()
    else:
        with st.expander("Validation Centre", expanded=True):
            _content()


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
        ws.column_dimensions[col_letter].width = min(max_length + 3, 40)
    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 22


def dataframe_to_excel_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            safe_sheet = str(sheet_name)[:31]
            if df is None:
                df = pd.DataFrame()
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
            format_excel_sheet(writer.book[safe_sheet])
    return output.getvalue()


def make_sample_template_bytes(df: pd.DataFrame) -> bytes:
    return dataframe_to_excel_bytes({"Template": df})


def get_sample_templates():
    templates = {}
    templates["Current GL Report"] = pd.DataFrame([
        {"Account code": "4000", "Debit": 0, "Credit": 25000, "Branch": "Sydney", "Net": -25000, "Date": "2026-04-01", "Period": "April 2026", "Description": "Sales invoice"},
        {"Account code": "5100", "Debit": 8000, "Credit": 0, "Branch": "Sydney", "Net": 8000, "Date": "2026-04-02", "Period": "April 2026", "Description": "Freight domestic cost"},
        {"Account code": "5200", "Debit": 3000, "Credit": 0, "Branch": "Melbourne", "Net": 3000, "Date": "2026-04-03", "Period": "April 2026", "Description": "Freight international overhead"},
    ])
    templates["COA Mapping"] = pd.DataFrame([
        {"Account code": "4000", "Account Name": "Sales Revenue", "Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Statement": "Income Statement", "Sign Convention": "positive", "Display Order": 1},
        {"Account code": "5100", "Account Name": "Freight Domestic", "Reporting Group": "Cost of Sales", "Reporting Subgroup": "Freight Domestic", "Statement": "Income Statement", "Sign Convention": "positive", "Display Order": 2},
        {"Account code": "5200", "Account Name": "Freight International", "Reporting Group": "Operating Expense", "Reporting Subgroup": "Freight International", "Statement": "Income Statement", "Sign Convention": "positive", "Display Order": 4},
    ])
    templates["KPI Master"] = pd.DataFrame([
        {"KPI Name": "Revenue", "Formula Type": "direct", "Numerator Group": "Revenue", "Denominator Group": "", "Output Type": "value", "Display Order": 1},
        {"KPI Name": "COGS", "Formula Type": "direct", "Numerator Group": "Cost of Sales", "Denominator Group": "", "Output Type": "value", "Display Order": 2},
        {"KPI Name": "Gross Profit", "Formula Type": "derived", "Numerator Group": "Revenue", "Denominator Group": "Cost of Sales", "Output Type": "value", "Display Order": 3},
        {"KPI Name": "Gross Margin %", "Formula Type": "ratio", "Numerator Group": "Gross Profit", "Denominator Group": "Revenue", "Output Type": "percent", "Display Order": 4},
        {"KPI Name": "Operating Expenses", "Formula Type": "direct", "Numerator Group": "Operating Expense", "Denominator Group": "", "Output Type": "value", "Display Order": 5},
        {"KPI Name": "Operating Profit", "Formula Type": "derived", "Numerator Group": "Gross Profit", "Denominator Group": "Operating Expense", "Output Type": "value", "Display Order": 6},
        {"KPI Name": "Operating Margin %", "Formula Type": "ratio", "Numerator Group": "Operating Profit", "Denominator Group": "Revenue", "Output Type": "percent", "Display Order": 7},
        {"KPI Name": "Opex as % of Revenue", "Formula Type": "ratio", "Numerator Group": "Operating Expense", "Denominator Group": "Revenue", "Output Type": "percent", "Display Order": 8},
    ])
    templates["Latest Previous Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 50000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 22000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 28000},
    ])
    templates["Budget Data"] = pd.DataFrame([
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Revenue", "Amount": 100000},
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Cost of Sales", "Amount": 60000},
        {"Month": "2026-01", "Branch": "Melbourne", "Reporting Group": "Revenue", "Amount": 85000},
    ])
    templates["Forecast P&L"] = pd.DataFrame([
        {"Period": "April 2026", "Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 120000},
        {"Period": "April 2026", "Reporting Group": "Cost of Sales", "Reporting Subgroup": "Cost of Sales", "Report Value": 72000},
        {"Period": "April 2026", "Reporting Group": "Operating Expense", "Reporting Subgroup": "Rent", "Report Value": 15000},
    ])
    templates["Forecast Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 65000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 28000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 37000},
    ])
    templates["Previous Year P&L"] = pd.DataFrame([
        {"Period": "April 2025", "Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 98000},
        {"Period": "April 2025", "Reporting Group": "Cost of Sales", "Reporting Subgroup": "Cost of Sales", "Report Value": 59000},
        {"Period": "April 2025", "Reporting Group": "Operating Expense", "Reporting Subgroup": "Rent", "Report Value": 13000},
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
    templates["Prior Period P&L"] = templates["Previous Year P&L"].copy()
    templates["Prior Period Balance Sheet"] = templates["Latest Previous Balance Sheet"].copy()
    templates["Prior Period KPI Pack"] = pd.DataFrame([
        {"KPI": "Revenue", "Value": 98000, "Display Value": 98000, "Output Type": "value"},
        {"KPI": "Gross Margin %", "Value": 39.80, "Display Value": "39.80%", "Output Type": "percent"},
        {"KPI": "Operating Margin %", "Value": 26.53, "Display Value": "26.53%", "Output Type": "percent"},
    ])
    return templates


# ----------------------------
# Standardizers / normalizers
# ----------------------------
def standardize_key_columns(gl, coa, kpi=None, latest_bs=None):
    gl = clean_columns(gl)
    coa = clean_columns(coa)
    gl.rename(columns={
        "Account Code": "Account code", "account code": "Account code", "ACCOUNT CODE": "Account code",
        "Branch ": "Branch", "branch": "Branch", "BRANCH": "Branch",
        "Debit ": "Debit", "debit": "Debit", "DEBIT": "Debit",
        "Credit ": "Credit", "credit": "Credit", "CREDIT": "Credit",
        "net": "Net", "NET": "Net", "Description ": "Description",
        "Account Name": "Account Name", "account name": "Account Name", "ACCOUNT NAME": "Account Name",
        "Account Description": "Account Name", "account description": "Account Name",
        "GL Name": "Account Name", "gl name": "Account Name",
        "Posting Date": "Date", "Txn Date": "Date", "Date ": "Date",
    }, inplace=True)
    coa.rename(columns={
        "Account Code": "Account code", "account code": "Account code", "ACCOUNT CODE": "Account code",
        "Reporting group": "Reporting Group", "reporting group": "Reporting Group",
        "Reporting subgroup": "Reporting Subgroup", "reporting subgroup": "Reporting Subgroup",
        "Statement type": "Statement", "statement": "Statement",
        "Sign convention": "Sign Convention", "sign convention": "Sign Convention",
        "Display order": "Display Order", "display order": "Display Order", "DISPLAY ORDER": "Display Order",
        "Report Order": "Display Order", "report order": "Display Order",
        "Account Name": "Account Name", "account name": "Account Name", "ACCOUNT NAME": "Account Name",
        "Account Description": "Account Name", "account description": "Account Name",
        "GL Name": "Account Name", "gl name": "Account Name",
        "GL Description": "Account Name", "gl description": "Account Name",
    }, inplace=True)
    if kpi is not None:
        kpi = clean_columns(kpi)
        kpi.rename(columns={
            "Kpi Name": "KPI Name", "Kpi name": "KPI Name",
            "Formula type": "Formula Type", "Numerator group": "Numerator Group",
            "Denominator group": "Denominator Group", "Output type": "Output Type",
            "Display order": "Display Order",
        }, inplace=True)
    if latest_bs is not None:
        latest_bs = normalize_uploaded_bs(latest_bs, "Latest Previous Balance Sheet")
    return gl, coa, kpi, latest_bs


def normalize_uploaded_pnl(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={"Reporting group": "Reporting Group", "Reporting subgroup": "Reporting Subgroup", "Report value": "Report Value"}, inplace=True)
    validate_required_columns(df, ["Reporting Group", "Reporting Subgroup", "Report Value"], label)
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Reporting Subgroup"] = df["Reporting Subgroup"].astype(str).str.strip()
    df["Report Value"] = pd.to_numeric(df["Report Value"], errors="coerce").fillna(0)
    return df


def normalize_uploaded_bs(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={"Reporting group": "Reporting Group", "Reporting subgroup": "Reporting Subgroup", "Balance ": "Balance"}, inplace=True)
    validate_required_columns(df, ["Reporting Group", "Reporting Subgroup", "Balance"], label)
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Reporting Subgroup"] = df["Reporting Subgroup"].astype(str).str.strip()
    df["Balance"] = pd.to_numeric(df["Balance"], errors="coerce").fillna(0)
    return df


def normalize_plan_df(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={"Month ": "Month", "Branch ": "Branch", "Reporting group": "Reporting Group", "Amount ": "Amount", "Budget Amount": "Amount"}, inplace=True)
    validate_required_columns(df, ["Month", "Reporting Group", "Amount"], label)
    if "Branch" not in df.columns:
        df["Branch"] = "Consolidated"
    df["Month"] = df["Month"].astype(str).str.strip()
    df["Branch"] = df["Branch"].astype(str).str.strip().replace({"": "Consolidated", "nan": "Consolidated"})
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    return df


def normalize_benchmark_df(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={"Metric ": "Metric", "Benchmark": "Benchmark Value", "Benchmark %": "Benchmark Value"}, inplace=True)
    validate_required_columns(df, ["Metric", "Benchmark Value"], "Industry Benchmark File")
    df["Metric"] = df["Metric"].astype(str).str.strip()
    df["Benchmark Value"] = pd.to_numeric(df["Benchmark Value"], errors="coerce").fillna(0)
    return df


def normalize_ageing_df(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    df = clean_columns(df)
    rename_map = {
        "Customer": "Party Name", "Customer Name": "Party Name", "Supplier": "Party Name", "Supplier Name": "Party Name",
        "Vendor": "Party Name", "Vendor Name": "Party Name", "Invoice Number": "Document Number", "Bill Number": "Document Number",
        "Invoice No": "Document Number", "Bill No": "Document Number", "Outstanding": "Outstanding Amount",
        "Outstanding Balance": "Outstanding Amount", "Amount": "Outstanding Amount", "Due Date ": "Due Date",
        "Invoice Date ": "Document Date", "Bill Date": "Document Date", "Ageing Bucket": "Age Bucket", "Aging Bucket": "Age Bucket",
        "Age Bucket ": "Age Bucket", "Branch ": "Branch",
    }
    df.rename(columns=rename_map, inplace=True)
    validate_required_columns(df, ["Party Name", "Outstanding Amount"], f"{kind} Ageing")
    if "Branch" not in df.columns:
        df["Branch"] = "Unassigned"
    if "Document Number" not in df.columns:
        df["Document Number"] = ""
    if "Document Date" not in df.columns:
        df["Document Date"] = pd.NaT
    if "Due Date" not in df.columns:
        df["Due Date"] = pd.NaT
    if "Age Bucket" not in df.columns:
        df["Age Bucket"] = None
    df["Outstanding Amount"] = pd.to_numeric(df["Outstanding Amount"], errors="coerce").fillna(0)
    df["Document Date"] = pd.to_datetime(df["Document Date"], errors="coerce")
    df["Due Date"] = pd.to_datetime(df["Due Date"], errors="coerce")
    today = pd.Timestamp.today().normalize()
    def calc_bucket(row):
        existing = row.get("Age Bucket")
        if pd.notna(existing) and str(existing).strip():
            return str(existing).strip()
        due_date = row.get("Due Date")
        if pd.isna(due_date):
            return "Unknown"
        days_overdue = (today - due_date.normalize()).days
        if days_overdue <= 0:
            return "Current"
        if days_overdue <= 30:
            return "1-30"
        if days_overdue <= 60:
            return "31-60"
        if days_overdue <= 90:
            return "61-90"
        return "90+"
    df["Age Bucket"] = df.apply(calc_bucket, axis=1)
    df["Branch"] = df["Branch"].astype(str).str.strip()
    return df


# ----------------------------
# Finance calculations
# ----------------------------
REPORTING_GROUP_ORDER = {
    "Revenue": 1,
    "Sales": 1,
    "Cost of Sales": 2,
    "COGS": 2,
    "Cost of Goods Sold": 2,
    "Gross Profit": 3,
    "Operating Expense": 4,
    "Operating Expenses": 4,
    "Overheads": 4,
    "Opex": 4,
    "Operating Profit": 5,
    "EBITDA": 6,
    "Depreciation": 7,
    "EBIT": 8,
    "Other Income": 9,
    "Other Expenses": 10,
    "Finance Costs": 11,
    "Interest": 11,
    "Tax": 12,
    "Net Profit": 13,
    "Assets": 20,
    "Current Assets": 21,
    "Non Current Assets": 22,
    "Liabilities": 30,
    "Current Liabilities": 31,
    "Non Current Liabilities": 32,
    "Equity": 40,
}


def apply_reporting_order(df: pd.DataFrame) -> pd.DataFrame:
    """Sort reports in finance statement order instead of alphabetical order.

    If the COA Mapping has a Display Order column, that order is used first.
    Otherwise, the default REPORTING_GROUP_ORDER is used.
    """
    if df is None or df.empty:
        return df

    out = df.copy()

    if "Display Order" in out.columns:
        out["__Display Order"] = pd.to_numeric(out["Display Order"], errors="coerce")
    else:
        out["__Display Order"] = pd.NA

    if "Reporting Group" in out.columns:
        out["__Group Order"] = out["Reporting Group"].map(REPORTING_GROUP_ORDER).fillna(999)
    else:
        out["__Group Order"] = 999

    sort_cols = ["__Display Order", "__Group Order"]
    if "Reporting Group" in out.columns:
        sort_cols.append("Reporting Group")
    if "Reporting Subgroup" in out.columns:
        sort_cols.append("Reporting Subgroup")
    if "Account code" in out.columns:
        sort_cols.append("Account code")

    out = out.sort_values(sort_cols, na_position="last")
    return out.drop(columns=["__Display Order", "__Group Order"], errors="ignore").reset_index(drop=True)


def find_coa_duplicate_rows(coa: pd.DataFrame) -> pd.DataFrame:
    """Return duplicate COA Account code rows for user review.

    Duplicates are not removed silently. The app highlights them and asks the user
    to confirm before keeping the first mapping row for each duplicate Account code.
    """
    if coa is None or coa.empty or "Account code" not in coa.columns:
        return pd.DataFrame()

    temp = coa.copy()
    temp["Account code"] = temp["Account code"].astype(str).str.strip()
    dupes = temp[temp.duplicated("Account code", keep=False)].copy()
    if dupes.empty:
        return dupes

    dupes["Duplicate Review Note"] = "Duplicate Account code - review and decide which mapping should be kept"
    sort_cols = ["Account code"]
    if "Display Order" in dupes.columns:
        sort_cols.append("Display Order")
    return dupes.sort_values(sort_cols).reset_index(drop=True)


def resolve_coa_duplicate_rows(coa: pd.DataFrame, keep: str = "first") -> pd.DataFrame:
    """Resolve duplicate COA rows after user confirmation.

    This does not change the user's source Excel file. It only creates a cleaned
    copy for system processing.
    """
    if coa is None or coa.empty or "Account code" not in coa.columns:
        return coa
    cleaned = coa.copy()
    cleaned["Account code"] = cleaned["Account code"].astype(str).str.strip()
    return cleaned.drop_duplicates(subset=["Account code"], keep=keep).reset_index(drop=True)


def validate_coa_mapping_integrity(coa: pd.DataFrame, allow_duplicate_cleanup: bool = False) -> None:
    """Validate COA mapping integrity.

    Blank Account codes are always blocked. Duplicate Account codes are blocked
    unless the user has explicitly confirmed duplicate cleanup in the UI.
    """
    if coa is None or coa.empty or "Account code" not in coa.columns:
        return

    temp = coa.copy()
    temp["Account code"] = temp["Account code"].astype(str).str.strip()

    blank_codes = temp[temp["Account code"].isin(["", "nan", "None"])]
    if not blank_codes.empty:
        raise ValueError("COA Mapping has blank Account code rows. Remove or complete these rows.")

    dupes = find_coa_duplicate_rows(temp)
    if not dupes.empty and not allow_duplicate_cleanup:
        duplicate_codes = sorted(dupes["Account code"].astype(str).unique().tolist())
        raise ValueError(
            "COA Mapping has duplicate Account code rows. Review the duplicate table shown below. "
            "If you approve, tick the duplicate confirmation checkbox and the system will keep the first row for each duplicate Account code. "
            f"Duplicate Account codes: {duplicate_codes[:20]}"
        )



# ----------------------------
# COA mapping review helpers
# ----------------------------
CANONICAL_GROUPS = {
    "revenue": "Revenue",
    "sales": "Revenue",
    "income": "Revenue",
    "cost of sales": "COGS",
    "cogs": "COGS",
    "cost of goods sold": "COGS",
    "direct costs": "COGS",
    "gross profit": "Gross Profit",
    "operating expenses": "Overheads",
    "overheads": "Overheads",
    "opex": "Overheads",
    "expenses": "Overheads",
    "other income": "Other Income",
    "other expenses": "Other Expenses",
    "interest": "Interest",
    "finance costs": "Interest",
    "tax": "Tax",
    "net profit": "Net Profit",
}

KEYWORD_MAPPING_RULES = [
    {"keyword": "sales", "suggested": ["Revenue"], "severity": "High"},
    {"keyword": "revenue", "suggested": ["Revenue"], "severity": "High"},
    {"keyword": "income", "suggested": ["Revenue", "Other Income"], "severity": "Medium"},
    {"keyword": "cogs", "suggested": ["COGS"], "severity": "High"},
    {"keyword": "cost of sales", "suggested": ["COGS"], "severity": "High"},
    {"keyword": "cost of goods", "suggested": ["COGS"], "severity": "High"},
    {"keyword": "purchases", "suggested": ["COGS"], "severity": "Medium"},
    {"keyword": "raw material", "suggested": ["COGS"], "severity": "High"},
    {"keyword": "materials", "suggested": ["COGS"], "severity": "Medium"},
    {"keyword": "direct labour", "suggested": ["COGS"], "severity": "Medium"},
    {"keyword": "direct labor", "suggested": ["COGS"], "severity": "Medium"},
    {"keyword": "freight", "suggested": ["COGS", "Overheads"], "severity": "Medium"},
    {"keyword": "shipping", "suggested": ["COGS", "Overheads"], "severity": "Medium"},
    {"keyword": "delivery", "suggested": ["COGS", "Overheads"], "severity": "Medium"},
    {"keyword": "cartage", "suggested": ["COGS", "Overheads"], "severity": "Medium"},
    {"keyword": "rent", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "salary", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "wages", "suggested": ["Overheads", "COGS"], "severity": "Medium"},
    {"keyword": "admin", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "marketing", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "advertising", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "insurance", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "utilities", "suggested": ["Overheads"], "severity": "High"},
    {"keyword": "depreciation", "suggested": ["Overheads"], "severity": "Medium"},
    {"keyword": "interest", "suggested": ["Interest"], "severity": "High"},
    {"keyword": "finance charge", "suggested": ["Interest"], "severity": "High"},
    {"keyword": "tax", "suggested": ["Tax"], "severity": "High"},
]

VALID_PNL_GROUPS = {
    "Revenue", "COGS", "Gross Profit", "Overheads", "Operating Profit",
    "Other Income", "Other Expenses", "Interest", "Tax", "Net Profit"
}

BS_GROUP_KEYWORDS = ["asset", "liabil", "equity", "cash", "bank", "receivable", "payable", "inventory", "stock", "loan", "debt", "capital", "retained"]

def canonical_reporting_group(value: str) -> str:
    text = str(value or "").strip()
    key = text.lower()
    return CANONICAL_GROUPS.get(key, text)


def build_coa_mapping_review(coa: pd.DataFrame) -> pd.DataFrame:
    """Flag suspicious COA mappings without changing user data.

    This is advisory only. Finance classifications can vary by company, so the app
    detects and explains potential problems but lets the user decide.
    """
    if coa is None or coa.empty:
        return pd.DataFrame(columns=["Account code", "Account Name", "Current Mapping", "Suggested Mapping", "Severity", "Reason", "Status"])

    df = coa.copy()
    if "Account Name" not in df.columns:
        df["Account Name"] = ""
    if "Reporting Group" not in df.columns:
        return pd.DataFrame(columns=["Account code", "Account Name", "Current Mapping", "Suggested Mapping", "Severity", "Reason", "Status"])

    review_rows = []
    for _, row in df.iterrows():
        account_code = str(row.get("Account code", "")).strip()
        account_name = str(row.get("Account Name", "")).strip()
        subgroup = str(row.get("Reporting Subgroup", "")).strip()
        current_raw = str(row.get("Reporting Group", "")).strip()
        current = canonical_reporting_group(current_raw)
        haystack = f"{account_name} {subgroup} {account_code}".lower()

        # If no Account Name is provided, we cannot intelligently infer category.
        if not account_name and not subgroup:
            continue

        for rule in KEYWORD_MAPPING_RULES:
            keyword = rule["keyword"]
            if keyword in haystack:
                suggested = rule["suggested"]
                if current not in suggested:
                    review_rows.append({
                        "Account code": account_code,
                        "Account Name": account_name,
                        "Current Mapping": current_raw,
                        "Suggested Mapping": " / ".join(suggested),
                        "Severity": rule["severity"],
                        "Reason": f"Keyword '{keyword}' usually maps to {', '.join(suggested)}, but current group is '{current_raw}'.",
                        "Status": "Review",
                    })
                break

        # Flag likely BS items mapped into P&L groups.
        if any(k in haystack for k in BS_GROUP_KEYWORDS) and current in VALID_PNL_GROUPS:
            review_rows.append({
                "Account code": account_code,
                "Account Name": account_name,
                "Current Mapping": current_raw,
                "Suggested Mapping": "Balance Sheet group",
                "Severity": "Medium",
                "Reason": "Account name looks balance-sheet related but is mapped to a P&L group.",
                "Status": "Review",
            })

    if not review_rows:
        return pd.DataFrame(columns=["Account code", "Account Name", "Current Mapping", "Suggested Mapping", "Severity", "Reason", "Status"])

    review = pd.DataFrame(review_rows).drop_duplicates()
    sev_order = {"High": 1, "Medium": 2, "Low": 3}
    review["__Severity Order"] = review["Severity"].map(sev_order).fillna(9)
    return review.sort_values(["__Severity Order", "Account code"]).drop(columns="__Severity Order").reset_index(drop=True)


def build_financial_logic_review(consolidated_pnl: pd.DataFrame) -> pd.DataFrame:
    """Basic reasonableness checks after the P&L is generated."""
    rows = []
    if consolidated_pnl is None or consolidated_pnl.empty:
        return pd.DataFrame(columns=["Check", "Status", "Details"])

    data = consolidated_pnl.copy()
    data["Canonical Group"] = data["Reporting Group"].apply(canonical_reporting_group)
    values = data.groupby("Canonical Group")["Report Value"].sum().to_dict()

    revenue = safe_float(values.get("Revenue", 0))
    cogs = safe_float(values.get("COGS", values.get("Cost of Sales", 0)))
    gross_profit = safe_float(values.get("Gross Profit", 0))
    overheads = safe_float(values.get("Overheads", values.get("Operating Expenses", 0)))
    operating_profit = safe_float(values.get("Operating Profit", 0))

    def add(check, ok, details):
        rows.append({"Check": check, "Status": "OK" if ok else "Review", "Details": details})

    add("Revenue exists", revenue != 0, f"Revenue total is {revenue:,.2f}.")
    if revenue != 0 and cogs != 0:
        add("COGS compared with revenue", abs(cogs) <= abs(revenue) * 1.5, f"COGS total is {cogs:,.2f}; Revenue total is {revenue:,.2f}.")
    if gross_profit != 0 and revenue != 0:
        add("Gross profit compared with revenue", abs(gross_profit) <= abs(revenue) * 1.5, f"Gross Profit total is {gross_profit:,.2f}; Revenue total is {revenue:,.2f}.")
    if operating_profit != 0 and gross_profit != 0:
        add("Operating profit compared with gross profit", abs(operating_profit) <= abs(gross_profit) * 2, f"Operating Profit total is {operating_profit:,.2f}; Gross Profit total is {gross_profit:,.2f}.")
    if overheads != 0 and revenue != 0:
        ratio = abs(overheads) / abs(revenue) * 100
        add("Overheads as % of revenue", ratio <= 80, f"Overheads are {ratio:.2f}% of revenue.")

    return pd.DataFrame(rows)

def build_pnl_detail(report_df: pd.DataFrame) -> pd.DataFrame:
    """Account-level P&L detail so similar GL accounts stay separate."""
    cols = ["Reporting Group", "Reporting Subgroup", "Account code"]
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=cols + ["Report Value"])

    out = account_level_report_values(report_df)
    out = out.drop(columns=["Sign Convention"], errors="ignore")
    return apply_reporting_order(out)


def build_balance_sheet_detail(bs_df: pd.DataFrame) -> pd.DataFrame:
    """Account-level balance sheet detail."""
    cols = ["Reporting Group", "Reporting Subgroup", "Account code"]
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=cols + ["Balance"])

    out = account_level_report_values(bs_df)
    out = out.drop(columns=["Sign Convention"], errors="ignore")
    out = out.rename(columns={"Report Value": "Balance"})
    return apply_reporting_order(out)


def apply_sign_convention_to_gl(row) -> float:
    """
    Keep transaction-level value as raw Net. Do not use abs() at transaction level.

    The display sign is applied after grouping by Account code, so debit and credit
    movements inside the same GL account are netted first.
    """
    net = row.get("Net", 0)
    if pd.isna(net):
        return 0.0
    return float(net)


def apply_sign_after_account_group(df: pd.DataFrame, value_col: str = "Report Value") -> pd.DataFrame:
    """Apply Sign Convention after account-level netting."""
    if df is None or df.empty:
        return df

    out = df.copy()
    if "Sign Convention" not in out.columns:
        out["Sign Convention"] = "positive"

    def signed_value(row):
        value = safe_float(row.get(value_col, 0))
        sign = str(row.get("Sign Convention", "positive")).strip().lower()
        display_value = abs(value)
        return -display_value if sign == "negative" else display_value

    out[value_col] = out.apply(signed_value, axis=1)
    return out


def account_level_report_values(report_df: pd.DataFrame, extra_cols=None) -> pd.DataFrame:
    """
    Net transactions by Account code first, then apply display sign convention.
    This fixes accounts with both debit and credit movements.
    """
    if report_df is None or report_df.empty:
        return pd.DataFrame()

    extra_cols = extra_cols or []
    df = report_df.copy()

    for col in ["Account Name", "Display Order", "Sign Convention"]:
        if col not in df.columns:
            df[col] = "" if col != "Display Order" else pd.NA

    group_cols = [
        "Reporting Group",
        "Reporting Subgroup",
        "Account code",
        "Account Name",
        "Display Order",
        "Sign Convention",
    ]

    for col in extra_cols:
        if col in df.columns and col not in group_cols:
            group_cols.append(col)

    grouped = df.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    grouped = apply_sign_after_account_group(grouped, "Report Value")
    return grouped


def infer_pnl_section_from_row(row) -> str:
    """Infer whether a P&L row belongs to Revenue, COGS, Overheads, etc.

    This deliberately checks both Reporting Group and Reporting Subgroup because
    many COA files use Reporting Group as the GL/report line name and use
    Reporting Subgroup as the real financial section, for example:
    - Reporting Group = Sales Revenue Labour...
    - Reporting Subgroup = Income
    """
    group = str(row.get("Reporting Group", "") or "").strip()
    subgroup = str(row.get("Reporting Subgroup", "") or "").strip()
    text = f"{group} {subgroup}".lower()

    # Most specific checks first.
    if any(k in text for k in ["cost of goods", "cost of sales", "cogs", "direct cost", "direct costs"]):
        return "COGS"
    if any(k in text for k in ["other income", "sundry income", "non operating income"]):
        return "Other Income"
    if any(k in text for k in ["other expense", "other expenses", "non operating expense"]):
        return "Other Expenses"
    if any(k in text for k in ["interest", "finance cost", "finance costs", "borrowing cost"]):
        return "Interest"
    if "tax" in text:
        return "Tax"
    if any(k in text for k in ["sales", "revenue", "income"]):
        return "Revenue"
    if any(k in text for k in ["operating expense", "operating expenses", "overhead", "overheads", "opex", "expense", "expenses"]):
        return "Overheads"
    if "gross profit" in text:
        return "Calculated"
    if any(k in text for k in ["net profit", "profit after tax", "profit for the period"]):
        return "Calculated"
    if any(k in text for k in ["operating profit", "ebit", "ebitda"]):
        return "Calculated"
    return "Other"


def _sum_section(pnl_df: pd.DataFrame, section: str) -> float:
    if pnl_df is None or pnl_df.empty or "__Section" not in pnl_df.columns:
        return 0.0
    return float(pd.to_numeric(pnl_df.loc[pnl_df["__Section"] == section, "Report Value"], errors="coerce").fillna(0).sum())


def _make_pnl_total_row(label: str, value: float, order: float, line_type: str = "Total") -> dict:
    return {
        "Reporting Group": label,
        "Reporting Subgroup": "",
        "Display Order": order,
        "Report Value": round(float(value), 2),
        "Line Type": line_type,
    }


def add_pnl_subtotals(base_pnl: pd.DataFrame) -> pd.DataFrame:
    """Insert management-report totals into P&L.

    Output order:
    Revenue lines -> Total Revenue -> COGS lines -> Total COGS -> Gross Profit
    -> Overheads lines -> Total Overheads -> other sections -> Net Profit.
    """
    if base_pnl is None or base_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"])

    pnl = base_pnl.copy()
    if "Display Order" not in pnl.columns:
        pnl["Display Order"] = pd.NA
    pnl["Display Order"] = pd.to_numeric(pnl["Display Order"], errors="coerce")
    pnl["Report Value"] = pd.to_numeric(pnl["Report Value"], errors="coerce").fillna(0).round(2)
    pnl["__Section"] = pnl.apply(infer_pnl_section_from_row, axis=1)
    pnl["Line Type"] = "Detail"

    section_sort = {
        "Revenue": 1,
        "COGS": 2,
        "Overheads": 4,
        "Other Income": 6,
        "Other Expenses": 7,
        "Interest": 8,
        "Tax": 9,
        "Other": 10,
        "Calculated": 99,
    }
    pnl["__Section Order"] = pnl["__Section"].map(section_sort).fillna(99)
    pnl = pnl.sort_values(["__Section Order", "Display Order", "Reporting Group", "Reporting Subgroup"], na_position="last")

    revenue_total = _sum_section(pnl, "Revenue")
    cogs_raw = _sum_section(pnl, "COGS")
    overheads_raw = _sum_section(pnl, "Overheads")
    other_income = _sum_section(pnl, "Other Income")
    other_expenses_raw = _sum_section(pnl, "Other Expenses")
    interest_raw = _sum_section(pnl, "Interest")
    tax_raw = _sum_section(pnl, "Tax")

    # Costs can be uploaded/displayed as either positive or negative depending on Sign Convention.
    # For management P&L totals, we treat cost sections as deductions using absolute values.
    cogs_total = abs(cogs_raw)
    overheads_total = abs(overheads_raw)
    other_expenses_total = abs(other_expenses_raw)
    interest_total = abs(interest_raw)
    tax_total = abs(tax_raw)

    gross_profit = revenue_total - cogs_total
    net_profit = gross_profit - overheads_total + other_income - other_expenses_total - interest_total - tax_total

    output_rows = []

    def append_section(section: str):
        details = pnl[pnl["__Section"] == section].drop(columns=["__Section", "__Section Order"], errors="ignore")
        if not details.empty:
            output_rows.extend(details.to_dict("records"))

    append_section("Revenue")
    if revenue_total != 0:
        output_rows.append(_make_pnl_total_row("Total Revenue", revenue_total, 1.90))

    append_section("COGS")
    if cogs_total != 0:
        output_rows.append(_make_pnl_total_row("Total COGS", cogs_total, 2.90))

    if revenue_total != 0 or cogs_total != 0:
        output_rows.append(_make_pnl_total_row("Gross Profit", gross_profit, 3.00, "Subtotal"))

    append_section("Overheads")
    if overheads_total != 0:
        output_rows.append(_make_pnl_total_row("Total Overheads", overheads_total, 4.90))

    append_section("Other Income")
    if other_income != 0:
        output_rows.append(_make_pnl_total_row("Total Other Income", other_income, 6.90))

    append_section("Other Expenses")
    if other_expenses_total != 0:
        output_rows.append(_make_pnl_total_row("Total Other Expenses", other_expenses_total, 7.90))

    append_section("Interest")
    if interest_total != 0:
        output_rows.append(_make_pnl_total_row("Total Interest / Finance Costs", interest_total, 8.90))

    append_section("Tax")
    if tax_total != 0:
        output_rows.append(_make_pnl_total_row("Total Tax", tax_total, 9.90))

    append_section("Other")

    output_rows.append(_make_pnl_total_row("Net Profit", net_profit, 99.00, "Final Profit"))

    out = pd.DataFrame(output_rows)
    preferred_cols = ["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"]
    for col in preferred_cols:
        if col not in out.columns:
            out[col] = "" if col != "Report Value" else 0.0
    out["Report Value"] = pd.to_numeric(out["Report Value"], errors="coerce").fillna(0).round(2)
    return out[preferred_cols].reset_index(drop=True)


def build_pnl(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"])

    account_values = account_level_report_values(report_df)

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in account_values.columns:
        group_cols.append("Display Order")

    base_pnl = account_values.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    base_pnl = apply_reporting_order(base_pnl)
    return add_pnl_subtotals(base_pnl)


def build_balance_sheet_from_gl(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    account_values = account_level_report_values(bs_df)

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in account_values.columns:
        group_cols.append("Display Order")

    bs = (
        account_values.groupby(group_cols, dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Balance"})
    )
    return apply_reporting_order(bs)


def combine_opening_and_current_bs(opening_bs: pd.DataFrame, current_bs: pd.DataFrame) -> pd.DataFrame:
    if opening_bs is None or opening_bs.empty:
        return current_bs.copy()
    opening = opening_bs.copy()
    current = current_bs.copy()
    opening["Balance"] = pd.to_numeric(opening["Balance"], errors="coerce").fillna(0)
    current["Balance"] = pd.to_numeric(current["Balance"], errors="coerce").fillna(0)
    merged = opening.merge(current, on=["Reporting Group", "Reporting Subgroup"], how="outer", suffixes=("_opening", "_current")).fillna(0)
    merged["Balance"] = merged["Balance_opening"] + merged["Balance_current"]
    return apply_reporting_order(merged[["Reporting Group", "Reporting Subgroup", "Balance"]])


def build_kpis(report_df: pd.DataFrame, kpi_master: pd.DataFrame) -> pd.DataFrame:
    if kpi_master is None or kpi_master.empty:
        return None
    if report_df is not None and not report_df.empty:
        account_values = account_level_report_values(report_df)
        group_values = account_values.groupby("Reporting Group")["Report Value"].sum().to_dict()
    else:
        group_values = {}
    results, calculated = [], {}
    kpi_master = kpi_master.sort_values("Display Order").copy()
    for _, row in kpi_master.iterrows():
        kpi_name = str(row["KPI Name"]).strip()
        formula_type = str(row["Formula Type"]).strip().lower()
        numerator = str(row["Numerator Group"]).strip() if pd.notna(row["Numerator Group"]) else ""
        denominator = str(row["Denominator Group"]).strip() if pd.notna(row["Denominator Group"]) else ""
        output_type = str(row["Output Type"]).strip().lower()
        if formula_type == "direct":
            value = group_values.get(numerator, 0.0)
        elif formula_type == "derived":
            value = calculated.get(numerator, group_values.get(numerator, 0.0)) - calculated.get(denominator, group_values.get(denominator, 0.0))
        elif formula_type == "ratio":
            num_val = calculated.get(numerator, group_values.get(numerator, 0.0))
            den_val = calculated.get(denominator, group_values.get(denominator, 0.0))
            value = (num_val / den_val * 100) if den_val != 0 else 0.0
        else:
            value = 0.0
        calculated[kpi_name] = value
        results.append({"KPI": kpi_name, "Value": value, "Output Type": output_type})
    kpi_df = pd.DataFrame(results)
    kpi_df["Display Value"] = kpi_df.apply(lambda r: f"{r['Value']:.2f}%" if r["Output Type"] == "percent" else round(r["Value"], 2), axis=1)
    return kpi_df[["KPI", "Value", "Output Type", "Display Value"]]


def kpi_map_from_df(kpi_df: pd.DataFrame | None) -> dict:
    if kpi_df is None or kpi_df.empty:
        return {}
    return {row["KPI"]: row["Value"] for _, row in kpi_df.iterrows()}


def build_actuals_by_branch_reporting_group(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty:
        return pd.DataFrame(columns=["Branch", "Reporting Group", "Actual"])
    account_values = account_level_report_values(pnl_mapped, extra_cols=["Branch"])
    return (
        account_values.groupby(["Branch", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Actual"})
    )


def compare_plan_vs_actual(actuals_df: pd.DataFrame, plan_df: pd.DataFrame, label: str) -> pd.DataFrame:
    if plan_df is None or plan_df.empty:
        return pd.DataFrame(columns=["Branch", "Reporting Group", "Actual", label, "Variance", "Variance %"])
    plan_agg = plan_df.groupby(["Branch", "Reporting Group"], dropna=False)["Amount"].sum().reset_index().rename(columns={"Amount": label})
    merged = actuals_df.merge(plan_agg, on=["Branch", "Reporting Group"], how="outer").fillna(0)
    merged["Variance"] = merged["Actual"] - merged[label]
    merged["Variance %"] = merged.apply(lambda r: (r["Variance"] / r[label] * 100) if r[label] != 0 else 0.0, axis=1)
    return merged.sort_values(["Branch", "Reporting Group"]).reset_index(drop=True)


def summarize_plan_vs_actual(compare_df: pd.DataFrame, label: str) -> pd.DataFrame:
    if compare_df is None or compare_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Actual", label, "Variance", "Variance %"])
    out = compare_df.groupby("Reporting Group", dropna=False)[["Actual", label, "Variance"]].sum().reset_index()
    out["Variance %"] = out.apply(lambda r: (r["Variance"] / r[label] * 100) if r[label] != 0 else 0.0, axis=1)
    return out.sort_values("Reporting Group").reset_index(drop=True)


def compare_pnl_to_forecast(actual_pnl: pd.DataFrame, forecast_pnl: pd.DataFrame) -> pd.DataFrame:
    if actual_pnl is None or actual_pnl.empty or forecast_pnl is None or forecast_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Actual", "Forecast", "Variance", "Variance %"])
    actual = actual_pnl.copy().rename(columns={"Report Value": "Actual"})
    forecast = forecast_pnl.copy().rename(columns={"Report Value": "Forecast"})
    merged = actual.merge(forecast, on=["Reporting Group", "Reporting Subgroup"], how="outer").fillna(0)
    merged["Variance"] = merged["Actual"] - merged["Forecast"]
    merged["Variance %"] = merged.apply(lambda r: (r["Variance"] / r["Forecast"] * 100) if r["Forecast"] != 0 else 0.0, axis=1)
    return merged.sort_values(["Reporting Group", "Reporting Subgroup"]).reset_index(drop=True)


def compare_pnl_to_previous_year(actual_pnl: pd.DataFrame, previous_pnl: pd.DataFrame) -> pd.DataFrame:
    if actual_pnl is None or actual_pnl.empty or previous_pnl is None or previous_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Actual", "Previous Year", "Variance", "Variance %"])
    actual = actual_pnl.copy().rename(columns={"Report Value": "Actual"})
    previous = previous_pnl.copy().rename(columns={"Report Value": "Previous Year"})
    merged = actual.merge(previous, on=["Reporting Group", "Reporting Subgroup"], how="outer").fillna(0)
    merged["Variance"] = merged["Actual"] - merged["Previous Year"]
    merged["Variance %"] = merged.apply(lambda r: (r["Variance"] / r["Previous Year"] * 100) if r["Previous Year"] != 0 else 0.0, axis=1)
    return merged.sort_values(["Reporting Group", "Reporting Subgroup"]).reset_index(drop=True)


def build_ageing_summary(df: pd.DataFrame | None, kind: str) -> dict:
    if df is None or df.empty:
        return {"total": 0.0, "overdue": 0.0, "overdue_pct": 0.0, "by_bucket": pd.DataFrame(), "by_branch": pd.DataFrame(), "top_parties": pd.DataFrame(), "kind": kind}
    total = float(df["Outstanding Amount"].sum())
    overdue_df = df[df["Age Bucket"].isin(["1-30", "31-60", "61-90", "90+"])]
    overdue = float(overdue_df["Outstanding Amount"].sum())
    overdue_pct = (overdue / total * 100) if total != 0 else 0.0
    bucket_order = ["Current", "1-30", "31-60", "61-90", "90+", "Unknown"]
    by_bucket = df.groupby("Age Bucket", dropna=False)["Outstanding Amount"].sum().reset_index()
    by_bucket["Age Bucket"] = pd.Categorical(by_bucket["Age Bucket"], categories=bucket_order, ordered=True)
    by_bucket = by_bucket.sort_values("Age Bucket")
    by_branch = df.groupby("Branch", dropna=False)["Outstanding Amount"].sum().reset_index().sort_values("Outstanding Amount", ascending=False)
    top_parties = df.groupby("Party Name", dropna=False)["Outstanding Amount"].sum().reset_index().sort_values("Outstanding Amount", ascending=False).head(10)
    return {"total": total, "overdue": overdue, "overdue_pct": overdue_pct, "by_bucket": by_bucket, "by_branch": by_branch, "top_parties": top_parties, "kind": kind}


def build_monthly_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Reporting Group", "Amount"])
    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Month", "Reporting Group", "Amount"])
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    account_values = account_level_report_values(df, extra_cols=["Month"])
    return (
        account_values.groupby(["Month", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Reporting Group"])
    )


def build_monthly_branch_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    if "Branch" not in df.columns:
        df["Branch"] = "Consolidated"
    df["Branch"] = df["Branch"].fillna("Consolidated").astype(str).str.strip().replace("", "Consolidated")
    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    # Revenue names vary by client, e.g. "Sales Revenue Labour..." rather than exactly "Revenue".
    group_text = df["Reporting Group"].astype(str).str.strip().str.lower()
    subgroup_text = df["Reporting Subgroup"].astype(str).str.strip().str.lower() if "Reporting Subgroup" in df.columns else ""
    rev = df[
        group_text.str.contains("revenue|sales|income", na=False)
        | (subgroup_text.str.contains("income|revenue|sales", na=False) if hasattr(subgroup_text, "str") else False)
    ].copy()

    if rev.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    account_values = account_level_report_values(rev, extra_cols=["Month", "Branch"])
    if account_values.empty or "Month" not in account_values.columns or "Branch" not in account_values.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    return (
        account_values.groupby(["Month", "Branch"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Branch"])
    )


def build_py_comparison(current_kpis: pd.DataFrame | None, prior_kpis: pd.DataFrame | None) -> pd.DataFrame:
    if current_kpis is None or current_kpis.empty or prior_kpis is None or prior_kpis.empty or "KPI" not in prior_kpis.columns or "Value" not in prior_kpis.columns:
        return pd.DataFrame(columns=["Metric", "Current", "Prior Year", "Variance", "Variance %"])
    cur = current_kpis[["KPI", "Value"]].rename(columns={"KPI": "Metric", "Value": "Current"})
    py = prior_kpis[["KPI", "Value"]].rename(columns={"KPI": "Metric", "Value": "Prior Year"})
    merged = cur.merge(py, on="Metric", how="inner")
    merged["Variance"] = merged["Current"] - merged["Prior Year"]
    merged["Variance %"] = merged.apply(lambda r: (r["Variance"] / r["Prior Year"] * 100) if r["Prior Year"] != 0 else 0.0, axis=1)
    return merged


def build_benchmark_comparison(current_kpis: pd.DataFrame | None, benchmark_df: pd.DataFrame | None, ar_summary=None, ap_summary=None) -> pd.DataFrame:
    rows = []
    if current_kpis is not None and not current_kpis.empty:
        for _, row in current_kpis.iterrows():
            rows.append({"Metric": row["KPI"], "Current Value": row["Value"]})
    if ar_summary is not None:
        rows.append({"Metric": "AR Overdue %", "Current Value": ar_summary["overdue_pct"]})
    if ap_summary is not None:
        rows.append({"Metric": "AP Overdue %", "Current Value": ap_summary["overdue_pct"]})
    current_df = pd.DataFrame(rows)
    if current_df.empty or benchmark_df is None or benchmark_df.empty:
        return pd.DataFrame(columns=["Metric", "Current Value", "Benchmark Value", "Gap"])
    merged = current_df.merge(benchmark_df, on="Metric", how="inner")
    merged["Gap"] = merged["Current Value"] - merged["Benchmark Value"]
    return merged.sort_values("Metric")


def rag_status(metric_name: str, current_value: float, benchmark_value=None) -> str:
    metric_name = str(metric_name).lower()
    if benchmark_value not in [None, ""]:
        gap = current_value - safe_float(benchmark_value)
        if "margin" in metric_name:
            return "Green" if gap >= 0 else ("Amber" if gap >= -3 else "Red")
        if "overdue" in metric_name:
            return "Green" if gap <= 0 else ("Amber" if gap <= 5 else "Red")
    if "gross margin" in metric_name:
        return "Green" if current_value >= 25 else ("Amber" if current_value >= 18 else "Red")
    if "operating margin" in metric_name:
        return "Green" if current_value >= 10 else ("Amber" if current_value >= 5 else "Red")
    if "opex" in metric_name:
        return "Green" if current_value <= 25 else ("Amber" if current_value <= 35 else "Red")
    if "overdue" in metric_name:
        return "Green" if current_value <= 20 else ("Amber" if current_value <= 35 else "Red")
    return "Amber"


def build_executive_summary(current_kpis, ar_summary=None, ap_summary=None, budget_summary=None, benchmark_compare=None, forecast_pnl_compare=None, previous_year_pnl_compare=None) -> pd.DataFrame:
    rows = []
    current_kpi_map = kpi_map_from_df(current_kpis)
    for metric in ["Revenue", "Gross Margin %", "Operating Margin %", "Opex as % of Revenue"]:
        current_value = safe_float(current_kpi_map.get(metric, 0))
        benchmark_value = ""
        if benchmark_compare is not None and not benchmark_compare.empty:
            match = benchmark_compare[benchmark_compare["Metric"] == metric]
            if not match.empty:
                benchmark_value = safe_float(match.iloc[0]["Benchmark Value"])
        rows.append({"Metric": metric, "Current Value": current_value, "Benchmark Value": benchmark_value, "Status": rag_status(metric, current_value, benchmark_value)})
    if ar_summary is not None:
        rows.append({"Metric": "AR Overdue %", "Current Value": safe_float(ar_summary["overdue_pct"]), "Benchmark Value": "", "Status": rag_status("AR Overdue %", safe_float(ar_summary["overdue_pct"]))})
    if ap_summary is not None:
        rows.append({"Metric": "AP Overdue %", "Current Value": safe_float(ap_summary["overdue_pct"]), "Benchmark Value": "", "Status": rag_status("AP Overdue %", safe_float(ap_summary["overdue_pct"]))})
    if budget_summary is not None and not budget_summary.empty and "Budget" in budget_summary.columns and budget_summary["Budget"].sum() != 0:
        pct = budget_summary["Variance"].sum() / budget_summary["Budget"].sum() * 100
        rows.append({"Metric": "Budget Variance %", "Current Value": pct, "Benchmark Value": "", "Status": "Green" if pct >= 0 else ("Amber" if pct >= -10 else "Red")})
    if forecast_pnl_compare is not None and not forecast_pnl_compare.empty and forecast_pnl_compare["Forecast"].sum() != 0:
        pct = forecast_pnl_compare["Variance"].sum() / forecast_pnl_compare["Forecast"].sum() * 100
        rows.append({"Metric": "Forecast Variance %", "Current Value": pct, "Benchmark Value": "", "Status": "Green" if pct >= 0 else ("Amber" if pct >= -10 else "Red")})
    if previous_year_pnl_compare is not None and not previous_year_pnl_compare.empty and previous_year_pnl_compare["Previous Year"].sum() != 0:
        pct = previous_year_pnl_compare["Variance"].sum() / previous_year_pnl_compare["Previous Year"].sum() * 100
        rows.append({"Metric": "Previous Year Variance %", "Current Value": pct, "Benchmark Value": "", "Status": "Green" if pct >= 0 else ("Amber" if pct >= -10 else "Red")})
    return pd.DataFrame(rows)


def detect_anomalies(consolidated_kpis, prior_kpis=None, ar_summary=None, ap_summary=None, budget_summary=None, forecast_pnl_compare=None):
    flags = []
    k = kpi_map_from_df(consolidated_kpis)
    if k.get("Revenue", 0) <= 0:
        flags.append("Revenue is zero or negative.")
    if k.get("Gross Margin %", 0) < 20:
        flags.append(f"Gross margin is low at {k.get('Gross Margin %', 0):.2f}%.")
    if k.get("Operating Margin %", 0) < 5:
        flags.append(f"Operating margin is weak at {k.get('Operating Margin %', 0):.2f}%.")
    if k.get("Opex as % of Revenue", 0) > 40:
        flags.append(f"Operating expenses are high at {k.get('Opex as % of Revenue', 0):.2f}% of revenue.")
    if ar_summary is not None and ar_summary["overdue_pct"] > 40:
        flags.append(f"AR overdue is high at {ar_summary['overdue_pct']:.2f}% of total receivables.")
    if ap_summary is not None and ap_summary["overdue_pct"] > 40:
        flags.append(f"AP overdue is high at {ap_summary['overdue_pct']:.2f}% of total payables.")
    if budget_summary is not None and not budget_summary.empty and "Budget" in budget_summary.columns and budget_summary["Budget"].sum() != 0:
        pct = budget_summary["Variance"].sum() / budget_summary["Budget"].sum() * 100
        if pct < -10:
            flags.append(f"Actual performance is {pct:.2f}% below budget.")
    if forecast_pnl_compare is not None and not forecast_pnl_compare.empty and forecast_pnl_compare["Forecast"].sum() != 0:
        pct = forecast_pnl_compare["Variance"].sum() / forecast_pnl_compare["Forecast"].sum() * 100
        if pct < -10:
            flags.append(f"Actual performance is {pct:.2f}% below forecast.")
    return flags


def create_excel_pack(consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary, branch_outputs, unmapped, executive_summary=None, monthly_actuals=None, monthly_branch_actuals=None, ar_df=None, ap_df=None, budget_compare=None, forecast_compare=None, py_compare=None, benchmark_compare=None, forecast_bs=None, fx_rate_info=None, country_indicators=None, external_benchmark_df=None, consolidated_pnl_detail=None, consolidated_bs_detail=None, coa_mapping_review=None, financial_logic_review=None):
    df_dict = {"Executive Summary": executive_summary if executive_summary is not None else pd.DataFrame(), "Consolidated P&L": consolidated_pnl}
    if consolidated_pnl_detail is not None and not consolidated_pnl_detail.empty:
        df_dict["P&L Detail by GL"] = consolidated_pnl_detail
    if consolidated_bs is not None and not consolidated_bs.empty:
        df_dict["Consolidated BS"] = consolidated_bs
    if consolidated_bs_detail is not None and not consolidated_bs_detail.empty:
        df_dict["BS Detail by GL"] = consolidated_bs_detail
    if forecast_bs is not None and not forecast_bs.empty:
        df_dict["Forecast BS"] = forecast_bs
    if consolidated_kpis is not None:
        df_dict["Consolidated KPIs"] = consolidated_kpis
    if branch_summary is not None and not branch_summary.empty:
        df_dict["Branch Summary KPIs"] = branch_summary
    if monthly_actuals is not None and not monthly_actuals.empty:
        df_dict["Monthly Trends"] = monthly_actuals
    if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
        df_dict["Branch Monthly Trends"] = monthly_branch_actuals
    if branch_outputs:
        for branch, reports in branch_outputs.items():
            df_dict[f"{str(branch)[:20]} P&L"] = reports.get("pnl", pd.DataFrame())
            if reports.get("pnl_detail") is not None and not reports.get("pnl_detail").empty:
                df_dict[f"{str(branch)[:18]} GL Detail"] = reports.get("pnl_detail")
            if reports.get("kpis") is not None:
                df_dict[f"{str(branch)[:20]} KPIs"] = reports["kpis"]
    if unmapped is not None and not unmapped.empty:
        df_dict["Unmapped Accounts"] = unmapped
    if ar_df is not None and not ar_df.empty:
        df_dict["AR Ageing"] = ar_df
    if ap_df is not None and not ap_df.empty:
        df_dict["AP Ageing"] = ap_df
    if budget_compare is not None and not budget_compare.empty:
        df_dict["Budget vs Actual"] = budget_compare
    if forecast_compare is not None and not forecast_compare.empty:
        df_dict["Actual vs Forecast"] = forecast_compare
    if py_compare is not None and not py_compare.empty:
        df_dict["Actual vs PY"] = py_compare
    if benchmark_compare is not None and not benchmark_compare.empty:
        df_dict["Benchmark Comparison"] = benchmark_compare
    if coa_mapping_review is not None and not coa_mapping_review.empty:
        df_dict["COA Mapping Review"] = coa_mapping_review
    if financial_logic_review is not None and not financial_logic_review.empty:
        df_dict["Financial Logic Review"] = financial_logic_review
    if external_benchmark_df is not None and not external_benchmark_df.empty:
        df_dict["Benchmark Source"] = external_benchmark_df
    if country_indicators is not None and not country_indicators.empty:
        df_dict["Country Indicators"] = country_indicators
    if fx_rate_info is not None:
        df_dict["FX Rate"] = pd.DataFrame([fx_rate_info])
    return dataframe_to_excel_bytes(df_dict)


def save_run_to_history(company_profile, consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary):
    company_name = company_profile.get("Company Name", "").strip()
    if not company_name:
        return
    company_slug = slugify_company_name(company_name)
    financial_year = company_profile.get("Financial Year", "unknown_year").strip().replace(" ", "_") or "unknown_year"
    reporting_period = company_profile.get("Report Period", company_profile.get("Reporting Period", "unknown_period")).strip().replace(" ", "_") or "unknown_period"
    run_folder = HISTORY_ROOT / company_slug / f"{financial_year}_{reporting_period}"
    run_folder.mkdir(parents=True, exist_ok=True)
    consolidated_pnl.to_excel(run_folder / "consolidated_pnl.xlsx", index=False)
    if consolidated_bs is not None and not consolidated_bs.empty:
        consolidated_bs.to_excel(run_folder / "consolidated_bs.xlsx", index=False)
    if consolidated_kpis is not None:
        consolidated_kpis.to_excel(run_folder / "consolidated_kpis.xlsx", index=False)
    if branch_summary is not None and not branch_summary.empty:
        branch_summary.to_excel(run_folder / "branch_summary.xlsx", index=False)


def list_saved_company_runs(company_name: str):
    company_folder = HISTORY_ROOT / slugify_company_name(company_name)
    if not company_folder.exists():
        return []
    return sorted([item.name for item in company_folder.iterdir() if item.is_dir()], reverse=True)


def restore_run_from_history(company_name: str, run_name: str):
    run_folder = HISTORY_ROOT / slugify_company_name(company_name) / run_name
    restored = {}
    if (run_folder / "consolidated_pnl.xlsx").exists():
        restored["prior_pnl"] = pd.read_excel(run_folder / "consolidated_pnl.xlsx")
    if (run_folder / "consolidated_bs.xlsx").exists():
        restored["prior_bs"] = pd.read_excel(run_folder / "consolidated_bs.xlsx")
    if (run_folder / "consolidated_kpis.xlsx").exists():
        restored["prior_kpis"] = pd.read_excel(run_folder / "consolidated_kpis.xlsx")
    return restored


def generate_ai_commentary(pnl_df, kpi_df, bs_df, profile, anomaly_flags=None, ar_summary=None, ap_summary=None, budget_summary=None, forecast_pnl_compare=None):
    if OpenAI is None:
        return "AI Commentary failed: openai package is not installed. Add openai to requirements.txt."
    if not os.getenv("OPENAI_API_KEY"):
        return "AI Commentary failed: OPENAI_API_KEY is not set in Streamlit secrets/environment."
    try:
        client = OpenAI()
        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        pnl_summary = pnl_df.to_string(index=False)[:3000] if pnl_df is not None and not pnl_df.empty else "No P&L data available."
        kpi_summary = kpi_df[["KPI", "Display Value"]].to_string(index=False)[:2000] if kpi_df is not None and not kpi_df.empty else "No KPI data available."
        bs_summary = bs_df.to_string(index=False)[:2000] if bs_df is not None and not bs_df.empty else "No Balance Sheet data available."
        anomaly_text = "\n".join(anomaly_flags) if anomaly_flags else "No anomaly flags detected."
        prompt = f"""
Prepare concise CFO commentary using only the data below.
Company profile: {profile}
Anomaly flags: {anomaly_text}
P&L: {pnl_summary}
KPIs: {kpi_summary}
Balance Sheet: {bs_summary}
Write: Executive Summary, Key Insights, Risks, Opportunities, Recommended Actions.
"""
        response = client.chat.completions.create(
            model=model_name,
            messages=[{"role": "developer", "content": "You are a concise CFO advisor."}, {"role": "user", "content": prompt}],
            temperature=0.3,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI Commentary failed: {str(e)}"



def dataframe_context(label: str, df: pd.DataFrame | None, max_rows: int = 30, max_chars: int = 4500) -> str:
    """Convert an in-memory dataframe into compact text for AI context."""
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return f"{label}: Not available."
    try:
        preview = df.head(max_rows).to_string(index=False)
        return f"{label} (first {min(len(df), max_rows)} of {len(df)} rows):\n{preview[:max_chars]}"
    except Exception as exc:
        return f"{label}: Could not render dataframe context ({exc})."


def build_ai_cfo_context() -> str:
    """Build a grounded context pack from the currently uploaded/processed finance data."""
    profile = st.session_state.get("company_profile", {}) or {}
    ar_summary = st.session_state.get("ar_summary")
    ap_summary = st.session_state.get("ap_summary")
    validation_report = st.session_state.get("last_validation_report") or {}

    summary_lines = [
        f"Company Profile: {profile}",
        f"Reporting Structure: {st.session_state.get('reporting_structure')}",
        f"Validation Score: {validation_report.get('score', 'Not available')}",
        f"Validation Critical Items: {len(validation_report.get('critical', []))}",
        f"Validation Warnings: {len(validation_report.get('warnings', []))}",
        f"Validation Recommendations: {len(validation_report.get('recommendations', []))}",
    ]
    if ar_summary is not None:
        summary_lines.append(f"AR Summary: Total={ar_summary.get('total', 0):,.2f}, Overdue={ar_summary.get('overdue', 0):,.2f}, Overdue %={ar_summary.get('overdue_pct', 0):.2f}%")
    if ap_summary is not None:
        summary_lines.append(f"AP Summary: Total={ap_summary.get('total', 0):,.2f}, Overdue={ap_summary.get('overdue', 0):,.2f}, Overdue %={ap_summary.get('overdue_pct', 0):.2f}%")

    context_parts = [
        "\n".join(summary_lines),
        dataframe_context("Consolidated P&L", st.session_state.get("consolidated_pnl")),
        dataframe_context("Consolidated KPIs", st.session_state.get("consolidated_kpis")),
        dataframe_context("Consolidated Balance Sheet", st.session_state.get("consolidated_bs")),
        dataframe_context("Branch KPI Summary", st.session_state.get("branch_summary")),
        dataframe_context("Budget vs Actual Summary", st.session_state.get("budget_summary")),
        dataframe_context("Forecast P&L Comparison", st.session_state.get("forecast_pnl_compare")),
        dataframe_context("Previous Year P&L Comparison", st.session_state.get("previous_year_pnl_compare")),
        dataframe_context("Benchmark Comparison", st.session_state.get("benchmark_compare")),
        dataframe_context("COA Mapping Review", st.session_state.get("coa_mapping_review")),
        dataframe_context("Financial Logic Review", st.session_state.get("financial_logic_review")),
        dataframe_context("AR Ageing Buckets", (ar_summary or {}).get("by_bucket") if ar_summary else None),
        dataframe_context("AP Ageing Buckets", (ap_summary or {}).get("by_bucket") if ap_summary else None),
        dataframe_context("Monthly Actuals", st.session_state.get("monthly_actuals")),
        dataframe_context("Unmapped GL Rows", st.session_state.get("unmapped"), max_rows=20),
    ]
    return "\n\n---\n\n".join(context_parts)[:28000]



def generic_ai_cfo_help_context() -> str:
    """Static product guidance the chatbot can use before any upload."""
    return """
AI CFO Copilot app guidance:
- Mandatory current-period uploads: Current GL Report and COA Mapping.
- GL mandatory columns in Consolidated Only mode: Account code, Debit, Credit. Branch is optional and defaults to Consolidated.
- GL mandatory columns in Branch / Business Unit mode: Account code, Debit, Credit, Branch.
- GL optional but useful columns: Net, Date, Description.
- COA Mapping mandatory columns: Account code, Reporting Group, Reporting Subgroup, Statement.
- COA optional columns: Sign Convention, Display Order, Account Name.
- Forecast P&L upload columns: Reporting Group, Reporting Subgroup, Report Value.
- Forecast Balance Sheet upload columns: Reporting Group, Reporting Subgroup, Balance.
- Previous Year P&L upload columns: Reporting Group, Reporting Subgroup, Report Value.
- Budget upload columns: Month, Branch, Reporting Group, Amount.
- AR/AP Ageing mandatory columns: Party Name, Outstanding Amount. Optional: Document Number, Document Date, Due Date, Branch, Age Bucket.
- Benchmark upload columns: Metric, Benchmark Value.
- Recommended flow: set company profile, choose reporting structure, download templates, replace sample rows, validate and upload, review validation centre, then view dashboard/reports.
- The chatbot can answer generic upload questions before upload. After upload, it can also answer data-specific CFO questions.
- Internet/benchmark capability: the app can use fetched FX rates, World Bank country indicators, starter industry benchmark data, and optionally web search if TAVILY_API_KEY is configured by the deployment owner.
""".strip()


def fetch_tavily_search_context(query: str, max_results: int = 5) -> str:
    """Optional web search context using Tavily if TAVILY_API_KEY is configured. Uses stdlib only."""
    api_key = os.getenv("TAVILY_API_KEY")
    if not api_key:
        return "Web search: Not configured. Set TAVILY_API_KEY to enable live web search."
    try:
        import json as _json
        payload = _json.dumps({
            "api_key": api_key,
            "query": query,
            "search_depth": "basic",
            "include_answer": True,
            "max_results": max_results,
        }).encode("utf-8")
        req = Request(
            "https://api.tavily.com/search",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urlopen(req, timeout=20) as response:
            data = _json.loads(response.read().decode("utf-8"))
        lines = []
        if data.get("answer"):
            lines.append(f"Search answer: {data.get('answer')}")
        for item in data.get("results", [])[:max_results]:
            title = item.get("title", "Untitled")
            url = item.get("url", "")
            content = item.get("content", "")[:700]
            lines.append(f"- {title}\n  URL: {url}\n  Summary: {content}")
        return "Web search results:\n" + "\n".join(lines) if lines else "Web search returned no usable results."
    except Exception as exc:
        return f"Web search failed: {exc}"


def build_external_ai_context(question: str) -> str:
    """Build optional internet/benchmark context from configured APIs and already loaded external data."""
    profile = st.session_state.get("company_profile", {}) or {}
    country = profile.get("Country", "") or "Australia"
    industry = profile.get("Industry", "") or "Other"
    currency = profile.get("Currency", "") or "AUD"
    q = (question or "").lower()

    parts = []
    parts.append(dataframe_context("Loaded Country Indicators", st.session_state.get("country_indicators")))
    parts.append(dataframe_context("Loaded External Benchmark Data", st.session_state.get("external_benchmark_df")))
    fx_info = st.session_state.get("fx_rate_info")
    if fx_info:
        parts.append(f"Loaded FX Rate: {fx_info}")

    wants_external = any(word in q for word in [
        "benchmark", "industry", "country", "market", "inflation", "gdp", "forex", "fx", "exchange", "external", "internet", "web", "compare"
    ])

    if wants_external:
        try:
            country_ind = fetch_country_indicators(country)
            parts.append(dataframe_context(f"Fresh World Bank Country Indicators for {country}", country_ind))
        except Exception as exc:
            parts.append(f"World Bank country indicator fetch failed: {exc}")
        try:
            starter_bench = get_builtin_industry_benchmarks(industry, country)
            parts.append(dataframe_context(f"Starter Industry Benchmarks for {industry} / {country}", starter_bench))
        except Exception as exc:
            parts.append(f"Starter benchmark load failed: {exc}")
        try:
            if currency and currency != "AUD":
                parts.append(f"FX context note: profile currency is {currency}. Use the app's FX section for exact conversion rate before financial comparison.")
        except Exception:
            pass
        parts.append(fetch_tavily_search_context(question))

    return "\n\n---\n\n".join([p for p in parts if p])[:18000]


def fallback_chatbot_answer(question: str, has_data: bool) -> str:
    """Useful fallback when OpenAI key is not available."""
    q = (question or "").lower()
    if any(w in q for w in ["branch", "business unit", "division"]):
        return "Branch is optional. Use **Consolidated Only** if the GL has no branch/business unit column; the app will use `Consolidated`. Use **Branch / Business Unit Reporting** only when you want branch-wise P&L/KPIs and your GL has a Branch column."
    if any(w in q for w in ["template", "upload", "column", "format"]):
        return "Use the **Download Sample Templates** section. For GL, the required columns are `Account code`, `Debit`, `Credit`; `Branch` is required only for Branch / Business Unit Reporting. COA needs `Account code`, `Reporting Group`, `Reporting Subgroup`, and `Statement`."
    if any(w in q for w in ["forecast", "3 way", "three way"]):
        return "For now, upload Forecast P&L and Forecast Balance Sheet directly. A driver-based 3-way model should later generate P&L, BS and Cash Flow from assumptions like revenue growth, gross margin, DSO, DPO, inventory days, capex, debt and tax."
    if any(w in q for w in ["benchmark", "industry", "country", "forex", "fx"]):
        return "The app supports uploaded benchmark files, starter industry benchmarks, World Bank country indicators, and FX-rate fetching. For live web search, configure `TAVILY_API_KEY`; otherwise the app uses built-in/API sources only."
    if not has_data:
        return "I can answer generic questions now. For company-specific analysis, upload and validate GL + COA first, then I can review P&L, KPIs, AR/AP, budget, forecast, benchmarks and mapping warnings."
    return "I can help analyse the uploaded financial data. Ask about margin movement, revenue, branch performance, AR/AP risk, budget variance, forecast variance, benchmarks, or mapping issues."


def answer_ai_cfo_question(question: str, mode: str = "Auto") -> str:
    """Chatbot answer supporting generic pre-upload help, uploaded-data analysis, and optional external benchmark context."""
    has_data = st.session_state.get("mapped") is not None
    if OpenAI is None or not os.getenv("OPENAI_API_KEY"):
        return fallback_chatbot_answer(question, has_data)

    generic_context = generic_ai_cfo_help_context()
    uploaded_context = build_ai_cfo_context() if has_data else "Uploaded finance data: Not available yet. User has not validated and uploaded files."
    external_context = build_external_ai_context(question)
    prior_messages = st.session_state.get("ai_cfo_chat_messages", [])[-10:]
    chat_history = "\n".join([f"{m.get('role', 'user')}: {m.get('content', '')}" for m in prior_messages])[:7000]

    system_prompt = """
You are an AI CFO chatbot inside a finance reporting web app.
You have three jobs:
1. Before upload: answer generic questions about upload templates, columns, validation, benchmarks, forecasts, and how to use the app.
2. After upload: answer data-specific CFO questions using the uploaded financial context.
3. When benchmark/external research is requested: use the external context provided. Do not pretend live internet search is available unless web search context is present.

Rules:
- Do not invent company-specific numbers. Use uploaded-data context only for data-specific analysis.
- If required data is missing, say exactly which upload or field is needed.
- External benchmark data can be broad and may need verification; clearly label it as external/starter/API-based.
- Be practical, CFO-style, and concise. Give recommended actions.
- Do not provide legal, tax, audit, or assurance conclusions.
""".strip()

    user_prompt = f"""
Chat mode selected by user: {mode}

Generic app guidance:
{generic_context}

Uploaded data context:
{uploaded_context}

External/benchmark context:
{external_context}

Recent chat history:
{chat_history}

Current user question:
{question}
"""
    try:
        client = OpenAI()
        model_name = os.getenv("OPENAI_MODEL", "gpt-4o-mini")
        response = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "developer", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
            temperature=0.25,
        )
        return response.choices[0].message.content
    except Exception as e:
        return f"AI CFO Chat failed: {str(e)}"


def prepare_data(gl_file, mapping_file, kpi_file=None, latest_bs_file=None, allow_duplicate_coa_cleanup: bool = False, reporting_structure: str = "Consolidated Only"):
    gl = pd.read_excel(gl_file)
    coa = pd.read_excel(mapping_file)
    kpi_master = pd.read_excel(kpi_file) if kpi_file is not None else None
    latest_bs = pd.read_excel(latest_bs_file) if latest_bs_file is not None else None
    gl, coa, kpi_master, latest_bs = standardize_key_columns(gl, coa, kpi_master, latest_bs)
    branch_required = reporting_structure == "Branch / Business Unit Reporting"
    gl_required_cols = ["Account code", "Debit", "Credit"] + (["Branch"] if branch_required else [])
    validate_required_columns(gl, gl_required_cols, "Current GL Report")
    validate_required_columns(coa, ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "COA Mapping")
    if kpi_master is not None:
        validate_required_columns(kpi_master, ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], "KPI Master")
    if "Branch" not in gl.columns:
        gl["Branch"] = "Consolidated"
    gl["Account code"] = gl["Account code"].astype(str).str.strip()
    coa["Account code"] = coa["Account code"].astype(str).str.strip()
    validate_coa_mapping_integrity(coa, allow_duplicate_cleanup=allow_duplicate_coa_cleanup)
    if allow_duplicate_coa_cleanup:
        coa = resolve_coa_duplicate_rows(coa, keep="first")
    gl["Branch"] = gl["Branch"].astype(str).str.strip().replace({"": "Consolidated", "nan": "Consolidated"})
    gl["Debit"] = pd.to_numeric(gl["Debit"], errors="coerce").fillna(0)
    gl["Credit"] = pd.to_numeric(gl["Credit"], errors="coerce").fillna(0)
    if "Net" not in gl.columns:
        gl["Net"] = gl["Debit"] - gl["Credit"]
    else:
        gl["Net"] = pd.to_numeric(gl["Net"], errors="coerce").fillna(gl["Debit"] - gl["Credit"])
    if "Date" in gl.columns:
        gl["Date"] = pd.to_datetime(gl["Date"], errors="coerce")
    data = gl.merge(coa, on="Account code", how="left", validate="many_to_one")
    unmapped = data[data["Reporting Group"].isna()].copy()
    mapped = data[data["Reporting Group"].notna()].copy()
    if "Sign Convention" not in mapped.columns:
        mapped["Sign Convention"] = "positive"
    mapped["Report Value"] = mapped.apply(apply_sign_convention_to_gl, axis=1)
    pnl_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()
    bs_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "balance sheet"].copy()
    return gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped


# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped", "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary", "detected_branches", "validation_passed", "company_profile", "bs_disclaimer", "ai_commentary", "prior_pnl", "prior_bs", "prior_kpis", "save_run_preference", "anomaly_flags", "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df", "budget_compare", "budget_summary", "benchmark_df", "py_compare", "benchmark_compare", "monthly_actuals", "monthly_branch_actuals", "executive_summary_df", "forecast_pnl", "forecast_bs", "previous_year_pnl", "forecast_pnl_compare", "previous_year_pnl_compare", "fx_rate_info", "country_indicators", "external_benchmark_df", "consolidated_pnl_detail", "consolidated_bs_detail", "coa_duplicate_rows", "coa_mapping_review", "financial_logic_review", "last_validation_report", "reporting_structure", "ai_cfo_chat_messages"
]:
    if key not in st.session_state:
        st.session_state[key] = None
if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}
if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False
if st.session_state["ai_cfo_chat_messages"] is None:
    st.session_state["ai_cfo_chat_messages"] = []
if "ai_cfo_panel_open" not in st.session_state or st.session_state["ai_cfo_panel_open"] is None:
    st.session_state["ai_cfo_panel_open"] = False

# ----------------------------
# UI - Product-style navigation
# ----------------------------
st.markdown("""
<style>
:root {--brand-1:#0f766e;--brand-2:#2563eb;--brand-3:#7c3aed;--ink:#111827;--muted:#6b7280;--line:#e5e7eb;--soft:#f8fafc;}
.block-container {max-width:1320px; padding-top:1.2rem;}
.main-card {padding:1.15rem 1.25rem;border:1px solid rgba(226,232,240,0.9);border-radius:18px;background:rgba(255,255,255,0.94);box-shadow:0 10px 30px rgba(15,23,42,0.06);margin-bottom:1rem;}
.hero-card {position:relative;overflow:hidden;min-height:295px;padding:2.1rem;border-radius:28px;color:white;background:linear-gradient(90deg, rgba(15,23,42,0.94) 0%, rgba(15,118,110,0.86) 48%, rgba(37,99,235,0.68) 100%),url('https://images.unsplash.com/photo-1554224155-6726b3ff858f?auto=format&fit=crop&w=1400&q=80');background-size:cover;background-position:center;box-shadow:0 22px 70px rgba(15,23,42,0.24);margin-bottom:1.25rem;}
.hero-kicker {display:inline-flex;align-items:center;gap:0.45rem;padding:0.42rem 0.72rem;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.24);border-radius:999px;font-weight:700;font-size:0.86rem;margin-bottom:0.9rem;}
.hero-title {font-size:2.65rem;line-height:1.05;font-weight:800;letter-spacing:-0.045em;max-width:780px;margin:0 0 0.8rem 0;}
.hero-subtitle {font-size:1.04rem;color:rgba(255,255,255,0.88);max-width:720px;margin-bottom:1.2rem;}
.hero-pill-row {display:flex;flex-wrap:wrap;gap:0.55rem;margin-top:1rem;}
.hero-pill {padding:0.52rem 0.78rem;border-radius:999px;background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.2);color:white;font-weight:700;font-size:0.86rem;}
.feature-card {padding:1.1rem;border-radius:18px;border:1px solid rgba(226,232,240,0.95);background:linear-gradient(180deg,#ffffff 0%,#f8fafc 100%);box-shadow:0 8px 26px rgba(15,23,42,0.05);height:100%;}
.feature-icon {width:42px;height:42px;border-radius:14px;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,rgba(15,118,110,0.12),rgba(37,99,235,0.12));color:#0f766e;font-size:1.35rem;margin-bottom:0.72rem;}
.feature-title {font-weight:800;color:var(--ink);margin-bottom:0.25rem;}
.feature-text {color:var(--muted);font-size:0.91rem;line-height:1.45;}
.setup-panel {border-radius:24px;padding:1.25rem;border:1px solid rgba(37,99,235,0.14);background:linear-gradient(180deg,rgba(239,246,255,0.82),rgba(255,255,255,0.96));box-shadow:0 12px 34px rgba(37,99,235,0.08);}
.section-title-row {display:flex;align-items:center;gap:0.65rem;margin:1.1rem 0 0.7rem 0;}
.section-title-icon {width:34px;height:34px;border-radius:12px;display:flex;align-items:center;justify-content:center;background:#eff6ff;color:#2563eb;font-size:1.08rem;}
.section-title-text {font-size:1.22rem;font-weight:800;color:#f8fafc;text-shadow:0 1px 1px rgba(0,0,0,0.25);}
.workflow-step {padding:0.8rem 1rem;border-radius:16px;background:#111827;border:1px solid rgba(148,163,184,0.30);text-align:center;font-weight:800;color:#f8fafc;box-shadow:0 6px 18px rgba(0,0,0,0.18);}
.workflow-step-done {background:#064e3b;border-color:#34d399;color:#d1fae5;}
.workflow-step-active {background:#1e3a8a;border-color:#60a5fa;color:#eff6ff;}
.small-muted {color:#6b7280;font-size:0.9rem;}
.alert-card {padding:0.9rem 1rem;border-radius:14px;background:#3f3f0f;border:1px solid #a3a329;margin-bottom:0.55rem;box-shadow:0 4px 14px rgba(251,146,60,0.08);color:#fef9c3;}
.alert-card b {color:#ffffff;}
.status-strip {display:grid;grid-template-columns:repeat(4,1fr);gap:0.8rem;margin-bottom:1rem;}
.status-mini {background:#ffffff;border:1px solid #e5e7eb;border-radius:16px;padding:0.85rem 1rem;box-shadow:0 7px 22px rgba(15,23,42,0.04);}
.status-mini b {font-size:0.82rem;color:#6b7280;display:block;margin-bottom:0.25rem;}
.status-mini span {font-weight:800;color:#111827;}
.upload-intro {padding:1rem 1.1rem;border-radius:18px;background:#111827;border:1px solid rgba(148,163,184,0.35);margin-bottom:1rem;color:#f8fafc;}
.upload-intro b {color:#ffffff;}
.upload-intro .small-muted {color:#cbd5e1;}
@media (max-width:900px) {.hero-title{font-size:2rem;}.hero-card{padding:1.4rem;min-height:auto;}.status-strip{grid-template-columns:repeat(2,1fr);}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div style="display:flex;align-items:center;gap:0.65rem;margin-bottom:0.2rem;">
  <div style="width:40px;height:40px;border-radius:14px;background:linear-gradient(135deg,#0f766e,#2563eb);display:flex;align-items:center;justify-content:center;color:white;font-size:1.35rem;font-weight:800;">▣</div>
  <div>
    <div style="font-size:2rem;font-weight:850;letter-spacing:-0.04em;color:#f8fafc;line-height:1;">AI CFO Copilot</div>
    <div style="color:#cbd5e1;font-size:0.95rem;margin-top:0.2rem;">Finance reporting, validation, forecasting, benchmarking and AI analysis in one guided workspace</div>
  </div>
</div>
""", unsafe_allow_html=True)

pages = [
    "🏠 Home",
    "📁 Data Upload",
    "📊 Dashboard",
    "📈 Reports",
    "💰 Working Capital",
    "🧠 Insights",
    "📤 Downloads",
]

page_query = st.query_params.get("page", "")
page_key_map = {
    "home": "🏠 Home",
    "upload": "📁 Data Upload",
    "dashboard": "📊 Dashboard",
    "reports": "📈 Reports",
    "working_capital": "💰 Working Capital",
    "insights": "🧠 Insights",
    "downloads": "📤 Downloads",
}
page_slug_map = {v: k for k, v in page_key_map.items()}
selected_page = page_key_map.get(page_query, "🏠 Home")

# Hide Streamlit's default sidebar completely. Navigation is handled by top action buttons.
st.markdown("""
<style>
[data-testid="stSidebar"] {display: none !important;}
[data-testid="collapsedControl"] {display: none !important;}
.block-container {padding-top: 2rem;}
.top-nav-row {
    padding: 0.75rem 0 0.35rem 0;
    border-bottom: 1px solid #eef0f3;
    margin-bottom: 1rem;
}
.nav-hint {
    color: #6b7280;
    font-size: 0.9rem;
    margin-bottom: 0.4rem;
}
div[data-testid="stButton"] > button {
    border-radius: 14px !important;
    border: 1px solid rgba(148,163,184,0.35) !important;
    background: linear-gradient(135deg, #111827, #1f2937) !important;
    color: #f8fafc !important;
    box-shadow: 0 8px 24px rgba(0,0,0,0.22) !important;
    min-height: 2.55rem !important;
}
div[data-testid="stButton"] > button:hover {
    border-color: #60a5fa !important;
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: #ffffff !important;
    box-shadow: 0 12px 30px rgba(37,99,235,0.28) !important;
}
div[data-testid="stButton"] > button p, div[data-testid="stButton"] > button span {
    color: #f8fafc !important;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="top-nav-row">', unsafe_allow_html=True)
st.markdown('<div class="nav-hint">CFO Workflow</div>', unsafe_allow_html=True)
nav_cols = st.columns(len(pages))
for nav_page, nav_col in zip(pages, nav_cols):
    is_current = selected_page == nav_page
    button_label = ("✓ " if is_current else "") + nav_page
    if nav_col.button(button_label, key=f"nav_{page_slug_map[nav_page]}", use_container_width=True):
        st.query_params["page"] = page_slug_map[nav_page]
        st.rerun()
st.markdown('</div>', unsafe_allow_html=True)

profile = st.session_state.get("company_profile", {}) or {}
report = st.session_state.get("last_validation_report") or {}
score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)
meta_cols = st.columns(4)
meta_cols[0].caption(f"Company: {profile.get('Company Name', 'Not set')}")
meta_cols[1].caption(f"Industry: {profile.get('Industry', 'Not set')}")
meta_cols[2].caption(f"Period: {get_report_period_label(profile)}")
meta_cols[3].caption(f"Readiness: {score}/100")

# Workflow status shown on every page
profile_done = bool((st.session_state.get("company_profile") or {}).get("Company Name"))
data_loaded = st.session_state.get("mapped") is not None
validation_ok = bool(st.session_state.get("validation_passed")) if data_loaded else False
reports_ready = st.session_state.get("consolidated_pnl") is not None
insights_ready = bool(st.session_state.get("ai_commentary"))

steps = [
    ("1 Configure", profile_done),
    ("2 Upload", data_loaded),
    ("3 Validate", validation_ok),
    ("4 Reports", reports_ready),
    ("5 Insights", insights_ready),
]
step_cols = st.columns(len(steps))
for idx, ((label, done), col) in enumerate(zip(steps, step_cols)):
    cls = "workflow-step-done" if done else ("workflow-step-active" if (idx == 0 and not profile_done) or (idx == 1 and profile_done and not data_loaded) or (idx == 2 and data_loaded and not validation_ok) else "")
    col.markdown(f'<div class="workflow-step {cls}">{"✓ " if done else ""}{label}</div>', unsafe_allow_html=True)

st.markdown("""
<style>
/* Floating AI CFO launcher. Clicking toggles the panel. */
.st-key-open_ai_cfo_global button {
    position: fixed !important;
    right: 26px !important;
    bottom: 28px !important;
    z-index: 999999 !important;
    width: 72px !important;
    height: 72px !important;
    min-height: 72px !important;
    border-radius: 50% !important;
    background: linear-gradient(135deg, #0f766e, #2563eb, #7c3aed) !important;
    color: #ffffff !important;
    font-size: 30px !important;
    border: 2px solid rgba(255,255,255,0.9) !important;
    box-shadow: 0 18px 45px rgba(37,99,235,0.38) !important;
    animation: aiFloat 2.8s ease-in-out infinite, aiPulse 1.8s ease-in-out infinite;
}
.st-key-open_ai_cfo_global button p { color: #ffffff !important; font-size: 30px !important; }
.st-key-open_ai_cfo_global button:hover {
    transform: scale(1.08) !important;
    box-shadow: 0 20px 55px rgba(124,58,237,0.45) !important;
}

/* Chatbot overlay panel. It is fixed and does NOT change page layout. */
.st-key-ai_cfo_overlay_panel {
    position: fixed !important;
    right: 24px !important;
    bottom: 112px !important;
    width: min(430px, calc(100vw - 32px)) !important;
    max-height: calc(100vh - 150px) !important;
    z-index: 999998 !important;
    overflow: auto !important;
    border: 1px solid rgba(96,165,250,0.38) !important;
    background: linear-gradient(180deg, rgba(15,23,42,0.98), rgba(17,24,39,0.98)) !important;
    border-radius: 24px !important;
    padding: 1rem !important;
    box-shadow: 0 24px 70px rgba(0,0,0,0.48) !important;
}
.st-key-ai_cfo_overlay_panel * { color: #f8fafc; }
.st-key-ai_cfo_overlay_panel label, .st-key-ai_cfo_overlay_panel p { color: #dbeafe !important; }
.st-key-ai_cfo_overlay_panel textarea, .st-key-ai_cfo_overlay_panel input {
    background: rgba(255,255,255,0.08) !important;
    color: #ffffff !important;
    border: 1px solid rgba(148,163,184,0.45) !important;
    border-radius: 12px !important;
}
.st-key-ai_cfo_overlay_panel textarea::placeholder, .st-key-ai_cfo_overlay_panel input::placeholder { color: #cbd5e1 !important; }
.ai-overlay-title {font-size:1.2rem;font-weight:850;color:#f8fafc;margin-bottom:0.25rem;}
.ai-overlay-sub {color:#cbd5e1;font-size:0.9rem;margin-bottom:0.85rem;line-height:1.35;}
.ai-bubble-user {background:#2563eb;color:#fff;border-radius:16px 16px 4px 16px;padding:0.68rem 0.78rem;margin:0.35rem 0 0.35rem auto;max-width:88%;font-size:0.92rem;}
.ai-bubble-assistant {background:rgba(255,255,255,0.08);color:#f8fafc;border:1px solid rgba(148,163,184,0.24);border-radius:16px 16px 16px 4px;padding:0.68rem 0.78rem;margin:0.35rem auto 0.35rem 0;max-width:94%;font-size:0.92rem;}
.ai-panel-note {color:#93c5fd;font-size:0.78rem;margin-top:0.45rem;}
@media (max-width: 768px) {
    .st-key-open_ai_cfo_global button { right: 16px !important; bottom: 18px !important; width: 62px !important; height: 62px !important; min-height: 62px !important; }
    .st-key-ai_cfo_overlay_panel { right: 12px !important; left: 12px !important; bottom: 92px !important; width: auto !important; max-height: calc(100vh - 120px) !important; }
}
</style>
""", unsafe_allow_html=True)

if st.button("🤖", key="open_ai_cfo_global", help="Open / close AI CFO Assistant"):
    st.session_state["ai_cfo_panel_open"] = not st.session_state.get("ai_cfo_panel_open", False)
    st.rerun()

# Global AI CFO overlay panel. Fixed position, so it does not disturb page alignment.
if st.session_state.get("ai_cfo_panel_open"):
    with st.container(key="ai_cfo_overlay_panel"):
        st.markdown('<div class="ai-overlay-title">🤖 AI CFO Assistant</div>', unsafe_allow_html=True)
        st.markdown('<div class="ai-overlay-sub">Ask upload questions before data, or CFO-style questions after upload. This panel floats above the page and will not move your dashboard or reports.</div>', unsafe_allow_html=True)

        close_col, clear_col = st.columns(2)
        if close_col.button("Close", use_container_width=True, key="close_ai_cfo_panel"):
            st.session_state["ai_cfo_panel_open"] = False
            st.rerun()
        if clear_col.button("Clear", use_container_width=True, key="clear_ai_cfo_panel"):
            st.session_state["ai_cfo_chat_messages"] = []
            st.rerun()

        chat_mode_inline = st.selectbox(
            "Mode",
            ["Auto", "General Help", "Data-specific CFO Analysis", "Internet & Benchmark Research"],
            key="global_ai_cfo_mode"
        )

        if not st.session_state.get("ai_cfo_chat_messages"):
            st.markdown('<div class="ai-bubble-assistant">Hi, I’m your AI CFO. Ask me about upload formats, mapping, validation, benchmarks, forecasts, or your uploaded financial data.</div>', unsafe_allow_html=True)

        for msg in st.session_state.get("ai_cfo_chat_messages", [])[-8:]:
            role = msg.get("role", "assistant")
            content = msg.get("content", "")
            if role == "user":
                st.markdown(f'<div class="ai-bubble-user">{content}</div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="ai-bubble-assistant">{content}</div>', unsafe_allow_html=True)

        with st.form("ai_cfo_overlay_form", clear_on_submit=True):
            user_question = st.text_area("Ask a question", placeholder="Example: Why is gross margin down?", height=85, key="ai_cfo_overlay_question")
            send_col, hint_col = st.columns([0.42, 0.58])
            send_clicked = send_col.form_submit_button("Send", use_container_width=True)
            hint_col.markdown('<div class="ai-panel-note">Use Close or click 🤖 again to hide this chat.</div>', unsafe_allow_html=True)

        if send_clicked and user_question.strip():
            st.session_state["ai_cfo_chat_messages"].append({"role": "user", "content": user_question.strip()})
            with st.spinner("AI CFO is thinking..."):
                inline_answer = answer_ai_cfo_question(user_question.strip(), mode=chat_mode_inline)
            st.session_state["ai_cfo_chat_messages"].append({"role": "assistant", "content": inline_answer})
            st.rerun()

if selected_page == "🏠 Home":
    profile = st.session_state.get("company_profile", {}) or {}
    report = st.session_state.get("last_validation_report") or {}
    critical = report.get("critical", [])
    warnings = report.get("warnings", [])
    recommendations = report.get("recommendations", [])
    score = report.get("score", 100 if isinstance(st.session_state.get("mapped"), pd.DataFrame) and not st.session_state.get("mapped").empty else 0)

    st.markdown("""
    <div class="hero-card">
        <div class="hero-kicker">✨ AI-powered finance workspace</div>
        <div class="hero-title">Turn messy finance data into board-ready decisions.</div>
        <div class="hero-subtitle">
            Upload GL, COA, budgets, forecast packs and AR/AP ageing. The app validates your data, flags mapping risks,
            builds P&L, balance sheet, KPI packs, dashboards and lets users ask an AI CFO questions.
        </div>
        <div class="hero-pill-row">
            <span class="hero-pill">📊 Management dashboards</span>
            <span class="hero-pill">✅ Data validation centre</span>
            <span class="hero-pill">🧠 AI CFO chatbot</span>
            <span class="hero-pill">🌍 FX & benchmarks</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown(f"""
    <div class="status-strip">
        <div class="status-mini"><b>Data Readiness</b><span>{score}/100</span></div>
        <div class="status-mini"><b>Critical Errors</b><span>{len(critical)}</span></div>
        <div class="status-mini"><b>Warnings</b><span>{len(warnings)}</span></div>
        <div class="status-mini"><b>Recommendations</b><span>{len(recommendations)}</span></div>
    </div>
    """, unsafe_allow_html=True)

    f1, f2, f3, f4 = st.columns(4)
    with f1:
        st.markdown('<div class="feature-card"><div class="feature-icon">🏢</div><div class="feature-title">1. Company Setup</div><div class="feature-text">Capture company, country, industry, period and reporting structure before any upload.</div></div>', unsafe_allow_html=True)
    with f2:
        st.markdown('<div class="feature-card"><div class="feature-icon">📁</div><div class="feature-title">2. Validate Uploads</div><div class="feature-text">Upload GL and COA. The Validation Centre checks missing columns, duplicates and mapping risks.</div></div>', unsafe_allow_html=True)
    with f3:
        st.markdown('<div class="feature-card"><div class="feature-icon">📈</div><div class="feature-title">3. Generate Reports</div><div class="feature-text">Create P&L, balance sheet, branch packs, KPI summary, variance and working capital views.</div></div>', unsafe_allow_html=True)
    with f4:
        st.markdown('<div class="feature-card"><div class="feature-icon">🤖</div><div class="feature-title">4. Ask AI CFO</div><div class="feature-text">Ask generic upload questions before data, then data-specific CFO questions after upload.</div></div>', unsafe_allow_html=True)

    st.markdown('<div class="section-title-row"><div class="section-title-icon">🏢</div><div class="section-title-text">Company Setup</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="setup-panel">', unsafe_allow_html=True)

    industry_options = ["Select Industry", "Manufacturing", "Wholesale / Distribution", "Retail", "Professional Services", "Construction", "Logistics", "Hospitality", "Healthcare", "Technology", "Other"]
    country_options = ["Select Country", "Australia", "India", "United States", "United Kingdom", "Canada", "New Zealand", "Other"]
    currency_options = ["Select Currency", "AUD", "INR", "USD", "GBP", "CAD", "NZD", "Other"]
    period_options = ["Monthly", "Quarterly", "Annual"]
    structure_options = ["Consolidated Only", "Branch / Business Unit Reporting"]

    def option_index(options, value, default=0):
        return options.index(value) if value in options else default

    c1, c2 = st.columns(2)
    with c1:
        company_name = st.text_input("Company Name *", value=profile.get("Company Name", ""), key="home_company_name")
        industry = st.selectbox("Industry", industry_options, index=option_index(industry_options, profile.get("Industry", "Select Industry")), key="home_industry")
        country = st.selectbox("Country", country_options, index=option_index(country_options, profile.get("Country", "Select Country")), key="home_country")
        state_region = st.text_input("State / Region", value=profile.get("State / Region", ""), key="home_state_region")
        financial_year = st.text_input("Financial Year", value=profile.get("Financial Year", ""), placeholder="Example: FY2026 or 2025-26", key="home_financial_year")
        report_period = st.text_input("Report Period *", value=profile.get("Report Period", ""), placeholder="Example: April 2026 or Q1 FY2026", key="home_report_period")
        period_start_date = st.date_input("Period Start Date", value=pd.to_datetime(profile.get("Period Start Date"), errors="coerce").date() if profile.get("Period Start Date") else None, key="home_period_start_date")
    with c2:
        currency = st.selectbox("Currency", currency_options, index=option_index(currency_options, profile.get("Currency", "Select Currency")), key="home_currency")
        tax_identifier = st.text_input("Tax Identifier / ABN / GSTIN (Optional)", value=profile.get("Tax Identifier", ""), key="home_tax_identifier")
        reporting_period = st.selectbox("Period Type", period_options, index=option_index(period_options, profile.get("Reporting Period", "Monthly")), key="home_reporting_period")
        period_end_date = st.date_input("Period End Date", value=pd.to_datetime(profile.get("Period End Date"), errors="coerce").date() if profile.get("Period End Date") else None, key="home_period_end_date")
        reporting_structure = st.radio("Reporting Structure", structure_options, index=option_index(structure_options, profile.get("Reporting Structure", "Consolidated Only")), key="home_reporting_structure", help="If Consolidated Only is selected, Branch is optional in GL and the app uses a default Consolidated unit.")
        benchmark_group = st.text_input("Benchmark Group (Optional)", value=profile.get("Benchmark Group", ""), key="home_benchmark_group")

    business_notes = st.text_area("Business Notes (Optional)", value=profile.get("Business Notes", ""), key="home_business_notes")
    save_run_preference = st.checkbox("Save this run for future comparison", value=st.session_state["save_run_preference"], key="home_save_run")

    h1, h2, h3 = st.columns([1.2, 1, 1])
    with h1:
        if st.button("Save Company Profile", use_container_width=True, key="home_save_company_profile"):
            if not company_name.strip():
                st.error("Company Name is mandatory.")
            elif industry == "Select Industry" or country == "Select Country":
                st.error("Please select at least Industry and Country.")
            elif not report_period.strip():
                st.error("Report Period is mandatory. Example: April 2026 or Q1 FY2026.")
            elif period_start_date and period_end_date and pd.to_datetime(period_start_date) > pd.to_datetime(period_end_date):
                st.error("Period Start Date cannot be after Period End Date.")
            else:
                st.session_state["company_profile"] = {"Company Name": company_name.strip(), "Industry": industry, "Country": country, "State / Region": state_region, "Financial Year": financial_year, "Report Period": report_period.strip(), "Period Start Date": str(period_start_date) if period_start_date else "", "Period End Date": str(period_end_date) if period_end_date else "", "Currency": currency if currency != "Select Currency" else "", "Tax Identifier": tax_identifier, "Reporting Period": reporting_period, "Reporting Structure": reporting_structure, "Benchmark Group": benchmark_group, "Business Notes": business_notes}
                st.session_state["reporting_structure"] = reporting_structure
                st.session_state["save_run_preference"] = save_run_preference
                st.success("Company profile saved. You can now go to Data Upload.")
    with h2:
        if st.button("Go to Data Upload", use_container_width=True, key="home_go_upload"):
            st.query_params["page"] = "upload"
            st.rerun()
    with h3:
        if st.button("Ask AI CFO", use_container_width=True, key="home_go_ai"):
            st.session_state["ai_cfo_panel_open"] = True
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("External Data & Benchmarks (Optional)", expanded=False):
        profile = st.session_state.get("company_profile", {})
        selected_country = profile.get("Country", "Australia") if profile else "Australia"
        selected_industry = profile.get("Industry", "Other") if profile else "Other"
        selected_currency = profile.get("Currency", "AUD") if profile else "AUD"
        default_target_currency = selected_currency if selected_currency and selected_currency != "Select Currency" else currency_for_country(selected_country)

        st.info("Optional setup for FX, country indicators and starter industry benchmarks. Uploaded benchmark files are still added from Data Upload, but external benchmark setup belongs here with the company profile.")
        fx1, fx2, fx3 = st.columns(3)
        with fx1:
            fx_base = st.selectbox("FX Base Currency", ["AUD", "INR", "USD", "GBP", "CAD", "NZD", "EUR"], index=0)
        with fx2:
            currency_options = ["AUD", "INR", "USD", "GBP", "CAD", "NZD", "EUR"]
            target_index = currency_options.index(default_target_currency) if default_target_currency in currency_options else 0
            fx_target = st.selectbox("FX Target Currency", currency_options, index=target_index)
        with fx3:
            fx_date = st.text_input("FX Date", value="latest", help="Use latest or YYYY-MM-DD")

        b1, b2, b3 = st.columns(3)
        with b1:
            if st.button("Fetch FX Rate", use_container_width=True):
                try:
                    st.session_state["fx_rate_info"] = fetch_fx_rate(fx_base, fx_target, fx_date)
                    st.success("FX rate fetched.")
                except Exception as e:
                    st.error(f"FX fetch failed: {e}")
        with b2:
            if st.button("Fetch Country Indicators", use_container_width=True):
                try:
                    st.session_state["country_indicators"] = fetch_country_indicators(selected_country)
                    st.success("Country indicators fetched.")
                except Exception as e:
                    st.error(f"Country indicator fetch failed: {e}")
        with b3:
            if st.button("Load Starter Industry Benchmarks", use_container_width=True):
                st.session_state["external_benchmark_df"] = get_builtin_industry_benchmarks(selected_industry, selected_country)
                st.success("Starter benchmark set loaded.")

        if st.session_state.get("fx_rate_info"):
            st.markdown("**FX Rate**")
            st.dataframe(pd.DataFrame([st.session_state["fx_rate_info"]]), use_container_width=True, hide_index=True)
        if st.session_state.get("country_indicators") is not None:
            st.markdown("**Country Indicators**")
            st.dataframe(st.session_state["country_indicators"], use_container_width=True, hide_index=True)
        if st.session_state.get("external_benchmark_df") is not None:
            st.markdown("**Loaded Benchmark Data**")
            st.dataframe(st.session_state["external_benchmark_df"], use_container_width=True, hide_index=True)



    profile = st.session_state.get("company_profile", {}) or {}
    if profile:
        st.markdown('<div class="section-title-row"><div class="section-title-icon">📋</div><div class="section-title-text">Current Company Profile</div></div>', unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(profile.items(), columns=["Field", "Value"]), use_container_width=True, hide_index=True)

    if critical or warnings or recommendations:
        st.markdown('<div class="section-title-row"><div class="section-title-icon">⚠️</div><div class="section-title-text">Items Needing Attention</div></div>', unsafe_allow_html=True)
        for item in (critical + warnings + recommendations)[:8]:
            st.markdown(f'<div class="alert-card"><b>{item.get("Area", "Review")}</b><br>{item.get("Issue", "")}</div>', unsafe_allow_html=True)
    elif st.session_state.get("mapped") is not None:
        st.success("No validation errors and no recommendations found in the last validation run.")
    else:
        st.info("Save company profile, then upload data to activate the Validation Centre.")

elif selected_page == "📁 Data Upload":
    st.markdown('<div class="section-title-row"><div class="section-title-icon">📁</div><div class="section-title-text">Data Upload</div></div>', unsafe_allow_html=True)
    profile = st.session_state.get("company_profile", {}) or {}
    if not profile or not profile.get("Company Name"):
        st.warning("Please complete Company Setup on the Home page before uploading files.")
        if st.button("Go to Home Setup", use_container_width=True, key="upload_go_home_setup"):
            st.query_params["page"] = "home"
            st.rerun()
    else:
        st.markdown(f"""
        <div class="upload-intro">
            <b>Uploading for:</b> {profile.get("Company Name", "")} &nbsp; | &nbsp;
            <b>Industry:</b> {profile.get("Industry", "")} &nbsp; | &nbsp;
            <b>Country:</b> {profile.get("Country", "")} &nbsp; | &nbsp;
            <b>Reporting:</b> {profile.get("Reporting Structure", "Consolidated Only")}
            <br><span class="small-muted">Company details are managed on the Home page. This section is only for files, templates, FX, benchmarks and validation.</span>
        </div>
        """, unsafe_allow_html=True)

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

        current_reporting_structure = st.session_state.get("company_profile", {}).get("Reporting Structure", "Consolidated Only")
        if current_reporting_structure == "Consolidated Only":
            st.info("Branch / Business Unit is optional for this company. If missing in GL, the app will use Consolidated automatically.")
        else:
            st.info("Branch / Business Unit Reporting is enabled. Branch column is mandatory in GL.")

        duplicate_resolution_confirmed = st.checkbox(
            "If duplicate COA Account codes are found, I have reviewed them and approve keeping the first row for each duplicate Account code",
            value=False,
            help="The original Excel file is not changed. This only cleans a processing copy after you approve."
        )

        if st.button("Validate & Upload Files", use_container_width=True):
            critical_items, warning_items, recommendation_items, info_items = [], [], [], []
            validation_success, loaded_files, previews = [], {}, {}

            profile = st.session_state.get("company_profile", {})
            reporting_structure = profile.get("Reporting Structure", "Consolidated Only")
            branch_required = reporting_structure == "Branch / Business Unit Reporting"

            def add_item(bucket, area, issue, recommendation=""):
                bucket.append({"Area": area, "Issue": str(issue), "Recommendation": recommendation})

            def log_success(file):
                validation_success.append({"File": file, "Status": "Valid"})

            if not profile or not profile.get("Company Name", "").strip():
                add_item(critical_items, "Company Profile", "Company profile is not saved.", "Save Company Profile before uploading files.")
            if gl_file is None:
                add_item(critical_items, "Current GL Report", "Mandatory file missing.", "Upload the Current GL Report template/file.")
            if mapping_file is None:
                add_item(critical_items, "COA Mapping", "Mandatory file missing.", "Upload the COA Mapping template/file.")

            gl_required_cols = ["Account code", "Debit", "Credit"] + (["Branch"] if branch_required else [])
            if not branch_required:
                add_item(info_items, "Reporting Structure", "Consolidated Only selected.", "Branch column is optional. Missing/blank Branch values will be treated as Consolidated.")
            else:
                add_item(info_items, "Reporting Structure", "Branch / Business Unit Reporting selected.", "Branch column is mandatory in the GL.")

            file_checks = [
                ("Current GL Report", gl_file, lambda f: standardize_key_columns(pd.read_excel(f), pd.DataFrame())[0], gl_required_cols, "gl"),
                ("COA Mapping", mapping_file, lambda f: standardize_key_columns(pd.DataFrame(), pd.read_excel(f))[1], ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "coa"),
                ("KPI Master", kpi_file, lambda f: standardize_key_columns(pd.DataFrame(), pd.DataFrame(), pd.read_excel(f))[2], ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], "kpi"),
                ("Latest Previous Balance Sheet", latest_bs_file, lambda f: normalize_uploaded_bs(pd.read_excel(f), "Latest Previous Balance Sheet"), ["Reporting Group", "Reporting Subgroup", "Balance"], "latest_bs"),
                ("Budget Data", budget_file, lambda f: normalize_plan_df(pd.read_excel(f), "Budget Data"), ["Month", "Reporting Group", "Amount"], "budget"),
                ("Forecast P&L", forecast_pnl_file, lambda f: normalize_uploaded_pnl(pd.read_excel(f), "Forecast P&L"), ["Reporting Group", "Reporting Subgroup", "Report Value"], "forecast_pnl"),
                ("Forecast Balance Sheet", forecast_bs_file, lambda f: normalize_uploaded_bs(pd.read_excel(f), "Forecast Balance Sheet"), ["Reporting Group", "Reporting Subgroup", "Balance"], "forecast_bs"),
                ("Previous Year P&L", previous_year_pnl_file, lambda f: normalize_uploaded_pnl(pd.read_excel(f), "Previous Year P&L"), ["Reporting Group", "Reporting Subgroup", "Report Value"], "previous_year_pnl"),
                ("AR Ageing", ar_file, lambda f: normalize_ageing_df(pd.read_excel(f), "AR"), ["Party Name", "Outstanding Amount"], "ar"),
                ("AP Ageing", ap_file, lambda f: normalize_ageing_df(pd.read_excel(f), "AP"), ["Party Name", "Outstanding Amount"], "ap"),
                ("Industry Benchmark File", benchmark_file, lambda f: normalize_benchmark_df(pd.read_excel(f)), ["Metric", "Benchmark Value"], "benchmark"),
            ]

            for file_label, file_obj, loader, required, key in file_checks:
                if file_obj is None:
                    if key in ["kpi", "latest_bs", "budget", "forecast_pnl", "forecast_bs", "previous_year_pnl", "ar", "ap", "benchmark"]:
                        add_item(info_items, file_label, "Optional file not uploaded.", "Upload this file if you want this analysis included.")
                    continue
                try:
                    df = loader(file_obj)
                    validate_required_columns(df, required, file_label)
                    if key == "gl" and not branch_required and "Branch" not in df.columns:
                        df["Branch"] = "Consolidated"
                    if key == "gl":
                        for period_issue in validate_gl_dates_against_profile(df, profile):
                            add_item(warning_items, period_issue["Area"], period_issue["Issue"], period_issue["Recommendation"])
                    if key == "coa":
                        duplicate_rows = find_coa_duplicate_rows(df)
                        st.session_state["coa_duplicate_rows"] = duplicate_rows
                        if not duplicate_rows.empty:
                            add_item(warning_items, "COA Mapping", f"Duplicate Account code rows found: {duplicate_rows['Account code'].nunique()} duplicated code(s).", "Review duplicate mapping rows. Tick the duplicate confirmation box to keep the first row for each duplicate for this run.")
                            if not duplicate_resolution_confirmed:
                                add_item(critical_items, "COA Mapping", "Duplicate Account code rows require user confirmation before processing.", "Fix the COA file or tick the duplicate confirmation checkbox after reviewing duplicates.")
                        mapping_review = build_coa_mapping_review(resolve_coa_duplicate_rows(df, keep="first") if not duplicate_rows.empty else df)
                        loaded_files["coa_mapping_review"] = mapping_review
                        if mapping_review is not None and not mapping_review.empty:
                            for _, r in mapping_review.head(10).iterrows():
                                add_item(recommendation_items, "COA Mapping Review", f"{r.get('Account code', '')} {r.get('Account Name', '')}: mapped to {r.get('Current Mapping', '')}", f"Suggested review: {r.get('Suggested Mapping', '')}. {r.get('Reason', '')}")
                    loaded_files[key] = df
                    previews[file_label] = df
                    log_success(file_label)
                except Exception as e:
                    raw_df = None
                    try:
                        raw_df = clean_columns(pd.read_excel(file_obj))
                    except Exception:
                        pass
                    found_cols = f" Found columns: {list(raw_df.columns)}" if raw_df is not None else ""
                    add_item(critical_items, file_label, f"{e}.{found_cols}", "Correct the file using the sample template and upload again.")

            if critical_items:
                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=True)
                st.stop()

            try:
                gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
                    gl_file,
                    mapping_file,
                    kpi_file,
                    latest_bs_file,
                    allow_duplicate_coa_cleanup=duplicate_resolution_confirmed,
                    reporting_structure=reporting_structure,
                )
                consolidated_pnl = build_pnl(pnl_mapped)
                consolidated_pnl_detail = build_pnl_detail(pnl_mapped)
                coa_mapping_review = build_coa_mapping_review(coa)
                financial_logic_review = build_financial_logic_review(consolidated_pnl)
                if financial_logic_review is not None and not financial_logic_review.empty:
                    for _, r in financial_logic_review.head(10).iterrows():
                        add_item(recommendation_items, "Financial Logic Review", r.get("Issue", "Financial logic item"), r.get("Recommendation", "Review financial classification and source data."))
                if unmapped is not None and not unmapped.empty:
                    add_item(warning_items, "GL Mapping", f"{len(unmapped)} GL row(s) are unmapped.", "Review unmapped account codes and update COA Mapping.")

                current_bs = build_balance_sheet_from_gl(bs_mapped)
                consolidated_bs_detail = build_balance_sheet_detail(bs_mapped)
                bs_disclaimer = None
                if latest_bs is not None:
                    consolidated_bs = combine_opening_and_current_bs(latest_bs, current_bs)
                else:
                    consolidated_bs = current_bs
                    bs_disclaimer = "Balance Sheet may not fully match because opening balances were not provided."
                    add_item(info_items, "Balance Sheet", "Latest previous Balance Sheet not uploaded.", "Upload opening/latest BS if you want stronger balance-sheet continuity.")
                consolidated_kpis = build_kpis(pnl_mapped, kpi_master) if kpi_master is not None else None
                detected_branches = sorted(pnl_mapped["Branch"].dropna().unique().tolist()) if not pnl_mapped.empty else []
                branch_outputs, branch_summary_rows = {}, []
                for branch in detected_branches:
                    branch_df = pnl_mapped[pnl_mapped["Branch"] == branch].copy()
                    branch_pnl = build_pnl(branch_df)
                    branch_pnl_detail = build_pnl_detail(branch_df)
                    branch_kpis = build_kpis(branch_df, kpi_master) if kpi_master is not None else None
                    branch_outputs[branch] = {"pnl": branch_pnl, "pnl_detail": branch_pnl_detail, "kpis": branch_kpis}
                    if branch_kpis is not None:
                        row = {"Branch": branch}
                        for _, r in branch_kpis.iterrows():
                            row[r["KPI"]] = r["Display Value"]
                        branch_summary_rows.append(row)
                branch_summary = pd.DataFrame(branch_summary_rows) if branch_summary_rows else pd.DataFrame()
                ar_df, ap_df = loaded_files.get("ar"), loaded_files.get("ap")
                ar_summary = build_ageing_summary(ar_df, "AR") if ar_df is not None else None
                ap_summary = build_ageing_summary(ap_df, "AP") if ap_df is not None else None
                budget_df = loaded_files.get("budget")
                uploaded_benchmark_df = loaded_files.get("benchmark")
                benchmark_df = merge_benchmark_sources(uploaded_benchmark_df, st.session_state.get("external_benchmark_df"))
                forecast_pnl = loaded_files.get("forecast_pnl")
                forecast_bs = loaded_files.get("forecast_bs")
                previous_year_pnl = loaded_files.get("previous_year_pnl")
                actuals_df = build_actuals_by_branch_reporting_group(pnl_mapped)
                budget_compare = compare_plan_vs_actual(actuals_df, budget_df, "Budget") if budget_df is not None else None
                budget_summary = summarize_plan_vs_actual(budget_compare, "Budget") if budget_compare is not None else None
                forecast_pnl_compare = compare_pnl_to_forecast(consolidated_pnl, forecast_pnl) if forecast_pnl is not None else None
                previous_year_pnl_compare = compare_pnl_to_previous_year(consolidated_pnl, previous_year_pnl) if previous_year_pnl is not None else None
                py_compare = build_py_comparison(consolidated_kpis, st.session_state.get("prior_kpis"))
                benchmark_compare = build_benchmark_comparison(consolidated_kpis, benchmark_df, ar_summary, ap_summary)
                monthly_actuals = build_monthly_actuals(pnl_mapped)
                monthly_branch_actuals = build_monthly_branch_actuals(pnl_mapped)
                executive_summary_df = build_executive_summary(consolidated_kpis, ar_summary=ar_summary, ap_summary=ap_summary, budget_summary=budget_summary, benchmark_compare=benchmark_compare, forecast_pnl_compare=forecast_pnl_compare, previous_year_pnl_compare=previous_year_pnl_compare)
                anomaly_flags = detect_anomalies(consolidated_kpis, prior_kpis=st.session_state.get("prior_kpis"), ar_summary=ar_summary, ap_summary=ap_summary, budget_summary=budget_summary, forecast_pnl_compare=forecast_pnl_compare) if consolidated_kpis is not None else []
                for flag in anomaly_flags:
                    add_item(recommendation_items, "Anomaly Review", flag, "Review source data and mapping before relying on final reports.")

                last_validation_report = {
                    "critical": critical_items,
                    "warnings": warning_items,
                    "recommendations": recommendation_items,
                    "info": info_items,
                    "score": calculate_validation_score(len(critical_items), len(warning_items), len(recommendation_items)),
                }

                for k, v in {
                    "gl": gl, "coa": coa, "kpi_master": kpi_master, "latest_bs": latest_bs, "mapped": mapped, "pnl_mapped": pnl_mapped, "bs_mapped": bs_mapped, "unmapped": unmapped, "consolidated_pnl": consolidated_pnl, "consolidated_pnl_detail": consolidated_pnl_detail, "consolidated_bs": consolidated_bs, "consolidated_bs_detail": consolidated_bs_detail, "consolidated_kpis": consolidated_kpis, "branch_outputs": branch_outputs, "branch_summary": branch_summary, "detected_branches": detected_branches, "validation_passed": unmapped.empty, "bs_disclaimer": bs_disclaimer, "ai_commentary": None, "ar_df": ar_df, "ap_df": ap_df, "ar_summary": ar_summary, "ap_summary": ap_summary, "budget_df": budget_df, "budget_compare": budget_compare, "budget_summary": budget_summary, "benchmark_df": benchmark_df, "benchmark_compare": benchmark_compare, "py_compare": py_compare, "monthly_actuals": monthly_actuals, "monthly_branch_actuals": monthly_branch_actuals, "executive_summary_df": executive_summary_df, "forecast_pnl": forecast_pnl, "forecast_bs": forecast_bs, "previous_year_pnl": previous_year_pnl, "forecast_pnl_compare": forecast_pnl_compare, "previous_year_pnl_compare": previous_year_pnl_compare, "anomaly_flags": anomaly_flags, "coa_mapping_review": coa_mapping_review, "financial_logic_review": financial_logic_review, "last_validation_report": last_validation_report, "reporting_structure": reporting_structure,
                }.items():
                    st.session_state[k] = v
                if st.session_state["save_run_preference"]:
                    save_run_to_history(st.session_state["company_profile"], consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary)

                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=False)
                if unmapped.empty:
                    st.success("Files loaded successfully. Validation Centre has been opened.")
                else:
                    st.warning("Files loaded, but unmapped GL rows were found. Review the Validation Centre.")
            except Exception as e:
                add_item(critical_items, "Processing", f"Files passed upload validation, but processing failed: {e}", "Review the error details, mapping file, and source files.")
                render_validation_centre(critical_items, warning_items, recommendation_items, info_items, previews, block_processing=True)
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
            prior_pnl_file = st.file_uploader("Prior Period P&L (Optional)", type=["xlsx"])
        with c2:
            prior_bs_file = st.file_uploader("Prior Period Balance Sheet (Optional)", type=["xlsx"])
            prior_kpi_file = st.file_uploader("Prior Period KPI Pack (Optional)", type=["xlsx"])
        if st.button("Load Prior Period Inputs", use_container_width=True):
            try:
                loaded_any = False
                if prior_pnl_file is not None:
                    st.session_state["prior_pnl"] = normalize_uploaded_pnl(pd.read_excel(prior_pnl_file), "Prior Period P&L")
                    loaded_any = True
                if prior_bs_file is not None:
                    st.session_state["prior_bs"] = normalize_uploaded_bs(pd.read_excel(prior_bs_file), "Prior Period Balance Sheet")
                    loaded_any = True
                if prior_kpi_file is not None:
                    pk = clean_columns(pd.read_excel(prior_kpi_file))
                    pk.rename(columns={"Kpi": "KPI", "Display value": "Display Value"}, inplace=True)
                    validate_required_columns(pk, ["KPI", "Value"], "Prior Period KPI Pack")
                    st.session_state["prior_kpis"] = pk
                    loaded_any = True
                st.success("Prior period data loaded successfully.") if loaded_any else st.info("No prior period file uploaded.")
            except Exception as e:
                st.error(f"Error loading prior period data: {e}")

    if st.session_state.get("gl") is not None:
        with st.expander("Validation Summary"):
            report = st.session_state.get("last_validation_report") or {}
            if report:
                c1, c2, c3, c4 = st.columns(4)
                c1.metric("Readiness Score", f"{report.get('score', 100)}/100")
                c2.metric("Critical", len(report.get("critical", [])))
                c3.metric("Warnings", len(report.get("warnings", [])))
                c4.metric("Recommendations", len(report.get("recommendations", [])))
                if st.button("Open Validation Centre Again", use_container_width=True):
                    render_validation_centre(report.get("critical", []), report.get("warnings", []), report.get("recommendations", []), report.get("info", []), {}, block_processing=False)
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("GL Rows", len(st.session_state["gl"]))
            m2.metric("Mapped Rows", len(st.session_state["mapped"]))
            m3.metric("Unmapped Rows", len(st.session_state["unmapped"]))
            m4.metric("Reporting Units", len(st.session_state["detected_branches"] or []))
            unmapped = st.session_state.get("unmapped")
            if unmapped is not None and not unmapped.empty:
                cols_to_show = [c for c in ["Account code", "Account Name", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
                st.dataframe(style_dataframe(unmapped[cols_to_show]), use_container_width=True)

    with st.expander("Required Columns Guide"):
        g1, g2 = st.columns(2)
        with g1:
            show_required_columns("Current GL Report", ["Account code", "Debit", "Credit"], ["Branch / Business Unit", "Net", "Date", "Period", "Description", "Account Name"])
            show_required_columns("COA Mapping", ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], ["Account Name", "Sign Convention", "Display Order"])
            show_required_columns("KPI Master", ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], [])
            show_required_columns("Latest Previous Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Budget Data", ["Month", "Reporting Group", "Amount"], ["Branch / Business Unit"])
            show_required_columns("Forecast P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], ["Period"])
        with g2:
            show_required_columns("Forecast Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Previous Year P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], ["Period"])
            show_required_columns("AR Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("AP Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("Industry Benchmark File", ["Metric", "Benchmark Value"], [])
            show_required_columns("Prior Period KPI Pack", ["KPI", "Value"], ["Display Value", "Output Type"])

    with st.expander("Download Sample Templates", expanded=True):
        st.info("Download a template, replace the sample rows with your own data, and upload the same file back into the app.")
        cols = st.columns(3)
        for idx, (name, df) in enumerate(get_sample_templates().items()):
            with cols[idx % 3]:
                st.download_button(
                    label=f"Download {name}",
                    data=make_sample_template_bytes(df),
                    file_name=f"{name.lower().replace(' ', '_').replace('&', 'and')}_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"tpl_{name}",
                )

elif selected_page == "📊 Dashboard":
    st.subheader("Dashboard")
    st.caption(f"Reporting period: {get_report_period_label(st.session_state.get('company_profile', {}))}")
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
        k = kpi_map_from_df(st.session_state["consolidated_kpis"])
        ar, ap = st.session_state.get("ar_summary"), st.session_state.get("ap_summary")
        st.markdown("### Core KPI Snapshot")
        k1, k2, k3, k4, k5 = st.columns(5)
        k1.metric("Revenue", f"{k.get('Revenue', 0):,.2f}")
        k2.metric("Gross Profit", f"{k.get('Gross Profit', 0):,.2f}")
        k3.metric("Gross Margin %", f"{k.get('Gross Margin %', 0):.2f}%")
        k4.metric("Operating Profit", f"{k.get('Operating Profit', 0):,.2f}")
        k5.metric("Operating Margin %", f"{k.get('Operating Margin %', 0):.2f}%")
        k6, k7, k8, k9, k10 = st.columns(5)
        k6.metric("Opex %", f"{k.get('Opex as % of Revenue', 0):.2f}%")
        k7.metric("Total AR", f"{ar['total']:,.2f}" if ar else "0.00")
        k8.metric("AR Overdue %", f"{ar['overdue_pct']:.2f}%" if ar else "0.00%")
        k9.metric("Total AP", f"{ap['total']:,.2f}" if ap else "0.00")
        k10.metric("AP Overdue %", f"{ap['overdue_pct']:.2f}%" if ap else "0.00%")
        st.markdown("### Key Charts")
        if st.session_state["budget_summary"] is not None and not st.session_state["budget_summary"].empty:
            st.markdown("**Budget vs Actual**")
            st.bar_chart(st.session_state["budget_summary"].set_index("Reporting Group")[["Actual", "Budget"]])
        if st.session_state["forecast_pnl_compare"] is not None and not st.session_state["forecast_pnl_compare"].empty:
            st.markdown("**Actual vs Forecast P&L**")
            st.bar_chart(st.session_state["forecast_pnl_compare"].groupby("Reporting Group")[["Actual", "Forecast"]].sum())
        if st.session_state["previous_year_pnl_compare"] is not None and not st.session_state["previous_year_pnl_compare"].empty:
            st.markdown("**Actual vs Previous Year P&L**")
            st.bar_chart(st.session_state["previous_year_pnl_compare"].groupby("Reporting Group")[["Actual", "Previous Year"]].sum())
        if st.session_state["benchmark_compare"] is not None and not st.session_state["benchmark_compare"].empty:
            st.markdown("**Industry Benchmark Comparison**")
            st.bar_chart(st.session_state["benchmark_compare"].set_index("Metric")[["Current Value", "Benchmark Value"]])

        if st.session_state.get("fx_rate_info"):
            with st.expander("FX Rate Used"):
                st.dataframe(pd.DataFrame([st.session_state["fx_rate_info"]]), use_container_width=True, hide_index=True)
        if st.session_state.get("country_indicators") is not None:
            with st.expander("Country Indicators"):
                st.dataframe(st.session_state["country_indicators"], use_container_width=True, hide_index=True)
        branch_rows = []
        if st.session_state["branch_outputs"]:
            for branch, reports in st.session_state["branch_outputs"].items():
                bk = kpi_map_from_df(reports.get("kpis"))
                branch_rows.append({"Branch": branch, "Revenue": bk.get("Revenue", 0), "Gross Margin %": bk.get("Gross Margin %", 0), "Operating Margin %": bk.get("Operating Margin %", 0)})
        branch_df = pd.DataFrame(branch_rows)
        if not branch_df.empty:
            c1, c2 = st.columns(2)
            with c1:
                st.markdown("**Revenue by Branch**")
                st.bar_chart(branch_df.set_index("Branch")[["Revenue"]])
            with c2:
                st.markdown("**Operating Margin % by Branch**")
                st.bar_chart(branch_df.set_index("Branch")[["Operating Margin %"]])

elif selected_page == "📈 Reports":
    st.subheader("Reports")
    st.caption(f"Reporting period: {get_report_period_label(st.session_state.get('company_profile', {}))}")
    sub_pnl, sub_bs, sub_kpi, sub_trends, sub_variance = st.tabs(["P&L", "Balance Sheet", "KPIs", "Trends", "Variance"])
    with sub_pnl:
        if st.session_state["consolidated_pnl"] is None:
            st.info("No P&L available yet.")
        else:
            st.markdown("### Consolidated P&L Summary")
            st.dataframe(style_dataframe(st.session_state["consolidated_pnl"]), use_container_width=True)
            if st.session_state.get("consolidated_pnl_detail") is not None and not st.session_state["consolidated_pnl_detail"].empty:
                st.markdown("### P&L Detail by GL Account")
                st.info("This view maps by exact Account code, so similar accounts such as Freight Domestic and Freight International remain separate.")
                st.dataframe(style_dataframe(st.session_state["consolidated_pnl_detail"]), use_container_width=True)
            if st.session_state["branch_outputs"]:
                st.markdown("### Branch P&L")
                for branch, reports in st.session_state["branch_outputs"].items():
                    with st.expander(str(branch)):
                        st.markdown("**Summary**")
                        st.dataframe(style_dataframe(reports["pnl"]), use_container_width=True)
                        if reports.get("pnl_detail") is not None and not reports["pnl_detail"].empty:
                            st.markdown("**Detail by GL Account**")
                            st.dataframe(style_dataframe(reports["pnl_detail"]), use_container_width=True)
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
            if st.session_state.get("consolidated_bs_detail") is not None and not st.session_state["consolidated_bs_detail"].empty:
                st.markdown("### Balance Sheet Detail by GL Account")
                st.dataframe(style_dataframe(st.session_state["consolidated_bs_detail"]), use_container_width=True)
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
            for group, title in [("revenue", "Revenue Trend"), ("gross profit", "Gross Profit Trend"), ("operating profit", "Operating Profit Trend")]:
                temp = monthly_actuals[monthly_actuals["Reporting Group"].astype(str).str.strip().str.lower() == group]
                if not temp.empty:
                    st.markdown(f"### {title}")
                    st.line_chart(temp.set_index("Month")[["Amount"]])
            if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
                st.markdown("### Branch Revenue Trend")
                st.line_chart(monthly_branch_actuals.pivot(index="Month", columns="Branch", values="Amount").fillna(0))
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

elif selected_page == "💰 Working Capital":
    st.subheader("Working Capital")
    st.caption(f"Reporting period: {get_report_period_label(st.session_state.get('company_profile', {}))}")
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
            if not ar["by_bucket"].empty:
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
            if not ap["by_bucket"].empty:
                st.bar_chart(ap["by_bucket"].set_index("Age Bucket")[["Outstanding Amount"]])
            st.dataframe(style_dataframe(ap["by_bucket"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["by_branch"]), use_container_width=True)
            st.dataframe(style_dataframe(ap["top_parties"]), use_container_width=True)

elif selected_page == "🧠 Insights":
    st.subheader("Insights")
    st.caption("Insights now focuses only on performance anomalies and AI commentary. Mapping, duplicates, and upload recommendations are handled in the Validation Centre during upload.")

    insight_anom, insight_ai = st.tabs(["Anomalies", "AI Commentary"])

    with insight_anom:
        flags = st.session_state.get("anomaly_flags", [])
        if flags:
            for flag in flags:
                st.warning(flag)
        else:
            st.success("No major financial anomalies detected based on current rules.")

        logic_review = st.session_state.get("financial_logic_review")
        if logic_review is not None and not logic_review.empty:
            st.markdown("### Financial Logic Notes")
            st.info("These are calculation-level checks. Upload-format, duplicate COA, and COA mapping recommendations stay inside the Validation Centre.")
            st.dataframe(style_dataframe(logic_review), use_container_width=True)

    with insight_ai:
        if st.session_state["mapped"] is None:
            st.warning("Please upload and validate data first.")
        elif not st.session_state["validation_passed"]:
            st.error("Resolve unmapped accounts before generating AI insights.")
        else:
            if st.button("Generate AI Insights", use_container_width=True):
                with st.spinner("Analyzing financials..."):
                    st.session_state["ai_commentary"] = generate_ai_commentary(st.session_state["consolidated_pnl"], st.session_state["consolidated_kpis"], st.session_state["consolidated_bs"], st.session_state["company_profile"], anomaly_flags=st.session_state.get("anomaly_flags", []), ar_summary=st.session_state.get("ar_summary"), ap_summary=st.session_state.get("ap_summary"), budget_summary=st.session_state.get("budget_summary"), forecast_pnl_compare=st.session_state.get("forecast_pnl_compare"))
            if st.session_state["ai_commentary"]:
                st.write(st.session_state["ai_commentary"])

elif selected_page == "💬 Ask AI CFO":
    st.subheader("AI CFO Chatbot")

    has_data = st.session_state.get("mapped") is not None
    profile = st.session_state.get("company_profile", {}) or {}

    st.markdown("""
    <div class="section-card">
        <h3>💬 Ask questions before or after upload</h3>
        <p class="small-muted">
        Before upload, ask about templates, required columns, validation, forecast setup, or benchmarks.
        After upload, ask data-specific questions about P&L, KPIs, AR/AP, budget, forecast, branch performance and mapping review.
        </p>
    </div>
    """, unsafe_allow_html=True)

    status_cols = st.columns(4)
    with status_cols[0]:
        st.metric("Data Status", "Loaded" if has_data else "Not uploaded")
    with status_cols[1]:
        st.metric("Company", profile.get("Company Name", "Not set"))
    with status_cols[2]:
        st.metric("Industry", profile.get("Industry", "Not set"))
    with status_cols[3]:
        st.metric("Country", profile.get("Country", "Not set"))

    chat_mode = st.radio(
        "Chat mode",
        ["Auto", "General Help", "Data-specific CFO Analysis", "Internet & Benchmark Research"],
        horizontal=True,
        help="Auto chooses based on your question. Internet & Benchmark Research uses available external APIs and optional Tavily search if configured.",
    )

    if not has_data:
        st.info("You can ask generic questions now. Upload and validate GL + COA to unlock company-specific CFO analysis.")
    else:
        st.success("Uploaded financial data is available for data-specific CFO questions.")

    st.markdown("### Suggested questions")
    qcols = st.columns(3)
    with qcols[0]:
        if st.button("What files should I upload?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "What files should I upload and which columns are mandatory?"
        if st.button("How should I map freight accounts?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "How should I decide whether freight domestic and freight international should be COGS or overheads?"
    with qcols[1]:
        if st.button("Compare me with benchmarks", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "Compare the company against available country and industry benchmarks. Tell me what data is missing if needed."
        if st.button("Explain forecast uploads", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "How should I upload forecast P&L and forecast balance sheet, and how will the app compare them?"
    with qcols[2]:
        if st.button("Why is GP changing?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "Why is gross profit or gross margin changing based on the uploaded data?"
        if st.button("What should management focus on?", use_container_width=True):
            st.session_state["_ai_cfo_pending_prompt"] = "What are the top management focus areas based on the uploaded data and validation review?"

    tool_cols = st.columns([1, 1, 3])
    with tool_cols[0]:
        if st.button("Clear Chat", use_container_width=True):
            st.session_state["ai_cfo_chat_messages"] = []
            st.rerun()
    with tool_cols[1]:
        if st.button("Fetch Benchmark Context", use_container_width=True):
            try:
                selected_country = profile.get("Country", "Australia") or "Australia"
                selected_industry = profile.get("Industry", "Other") or "Other"
                st.session_state["country_indicators"] = fetch_country_indicators(selected_country)
                st.session_state["external_benchmark_df"] = get_builtin_industry_benchmarks(selected_industry, selected_country)
                st.success("Benchmark context loaded. Ask the chatbot to compare benchmarks now.")
            except Exception as exc:
                st.error(f"Benchmark context fetch failed: {exc}")

    # Chat transcript
    st.markdown("### Conversation")
    if not st.session_state.get("ai_cfo_chat_messages"):
        with st.chat_message("assistant"):
            st.markdown(
                "Hi, I’m your AI CFO assistant. I can help with upload formats, mapping decisions, validation issues, benchmarks, forecasts, and uploaded-data analysis. What would you like to check?"
            )

    for msg in st.session_state.get("ai_cfo_chat_messages", []):
        with st.chat_message(msg.get("role", "assistant")):
            st.markdown(msg.get("content", ""))

    pending_prompt = st.session_state.pop("_ai_cfo_pending_prompt", None)
    typed_prompt = st.chat_input("Ask about uploads, benchmarks, mapping, forecasts, or your uploaded financial data...")
    user_question = pending_prompt or typed_prompt

    if user_question:
        st.session_state["ai_cfo_chat_messages"].append({"role": "user", "content": user_question})
        with st.chat_message("user"):
            st.markdown(user_question)
        with st.chat_message("assistant"):
            with st.spinner("AI CFO is thinking..."):
                answer = answer_ai_cfo_question(user_question, mode=chat_mode)
            st.markdown(answer)
        st.session_state["ai_cfo_chat_messages"].append({"role": "assistant", "content": answer})

    with st.expander("What can the chatbot use?"):
        st.write("Before upload: app guidance, template rules, validation rules, generic finance logic and available external benchmark APIs.")
        st.write("After upload: your uploaded P&L, balance sheet, KPIs, branch data, AR/AP, budget, forecast, benchmarks, validation issues and mapping review.")
        st.write("Live web search is optional and requires `TAVILY_API_KEY` in deployment secrets. Without it, the app uses FX APIs, World Bank indicators, starter benchmarks and uploaded benchmark files.")
        st.write("The chatbot does not permanently train itself and does not replace finance review.")


elif selected_page == "📤 Downloads":
    st.subheader("Downloads")
    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped GL rows before downloading reports.")
    else:
        full_pack_bytes = create_excel_pack(consolidated_pnl=st.session_state["consolidated_pnl"], consolidated_bs=st.session_state["consolidated_bs"], consolidated_kpis=st.session_state["consolidated_kpis"], branch_summary=st.session_state["branch_summary"], branch_outputs=st.session_state["branch_outputs"], unmapped=st.session_state["unmapped"], executive_summary=st.session_state["executive_summary_df"], monthly_actuals=st.session_state["monthly_actuals"], monthly_branch_actuals=st.session_state["monthly_branch_actuals"], ar_df=st.session_state["ar_df"], ap_df=st.session_state["ap_df"], budget_compare=st.session_state["budget_compare"], forecast_compare=st.session_state["forecast_pnl_compare"], py_compare=st.session_state["previous_year_pnl_compare"], benchmark_compare=st.session_state["benchmark_compare"], forecast_bs=st.session_state["forecast_bs"], fx_rate_info=st.session_state.get("fx_rate_info"), country_indicators=st.session_state.get("country_indicators"), external_benchmark_df=st.session_state.get("external_benchmark_df"))
        st.download_button(label="Download Full Management Pack", data=full_pack_bytes, file_name="full_management_pack.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        if st.session_state["unmapped"] is not None and not st.session_state["unmapped"].empty:
            st.download_button(label="Download Unmapped GL", data=st.session_state["unmapped"].to_csv(index=False).encode("utf-8"), file_name="unmapped_gl.csv", mime="text/csv", use_container_width=True)
