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
        "text-align": "left",
    })


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
        {"Account code": "4000", "Debit": 0, "Credit": 25000, "Branch": "Sydney", "Net": -25000, "Date": "2026-03-01", "Description": "Sales invoice"},
        {"Account code": "5100", "Debit": 8000, "Credit": 0, "Branch": "Sydney", "Net": 8000, "Date": "2026-03-02", "Description": "Freight domestic cost"},
        {"Account code": "5200", "Debit": 3000, "Credit": 0, "Branch": "Melbourne", "Net": 3000, "Date": "2026-03-03", "Description": "Freight international overhead"},
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
        {"Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 120000},
        {"Reporting Group": "Cost of Sales", "Reporting Subgroup": "Cost of Sales", "Report Value": 72000},
        {"Reporting Group": "Operating Expense", "Reporting Subgroup": "Rent", "Report Value": 15000},
    ])
    templates["Forecast Balance Sheet"] = pd.DataFrame([
        {"Reporting Group": "Assets", "Reporting Subgroup": "Cash", "Balance": 65000},
        {"Reporting Group": "Liabilities", "Reporting Subgroup": "Trade Payables", "Balance": 28000},
        {"Reporting Group": "Equity", "Reporting Subgroup": "Retained Earnings", "Balance": 37000},
    ])
    templates["Previous Year P&L"] = pd.DataFrame([
        {"Reporting Group": "Revenue", "Reporting Subgroup": "Sales", "Report Value": 98000},
        {"Reporting Group": "Cost of Sales", "Reporting Subgroup": "Cost of Sales", "Report Value": 59000},
        {"Reporting Group": "Operating Expense", "Reporting Subgroup": "Rent", "Report Value": 13000},
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

    detail = report_df.copy()
    if "Account Name" not in detail.columns:
        detail["Account Name"] = ""
    if "Display Order" not in detail.columns:
        detail["Display Order"] = pd.NA

    group_cols = ["Reporting Group", "Reporting Subgroup", "Account code", "Account Name", "Display Order"]
    out = detail.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    return apply_reporting_order(out)


def build_balance_sheet_detail(bs_df: pd.DataFrame) -> pd.DataFrame:
    """Account-level balance sheet detail."""
    cols = ["Reporting Group", "Reporting Subgroup", "Account code"]
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=cols + ["Balance"])

    detail = bs_df.copy()
    if "Account Name" not in detail.columns:
        detail["Account Name"] = ""
    if "Display Order" not in detail.columns:
        detail["Display Order"] = pd.NA

    group_cols = ["Reporting Group", "Reporting Subgroup", "Account code", "Account Name", "Display Order"]
    out = detail.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index().rename(columns={"Report Value": "Balance"})
    return apply_reporting_order(out)


def apply_sign_convention_to_gl(row) -> float:
    net = row.get("Net", 0)
    sign = str(row.get("Sign Convention", "positive")).strip().lower()
    if pd.isna(net):
        return 0.0
    val = abs(float(net))
    return -val if sign == "negative" else val


def build_pnl(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Report Value"])

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in report_df.columns:
        group_cols.append("Display Order")

    pnl = report_df.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    return apply_reporting_order(pnl)


def build_balance_sheet_from_gl(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in bs_df.columns:
        group_cols.append("Display Order")

    bs = (
        bs_df.groupby(group_cols, dropna=False)["Report Value"]
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
    group_values = report_df.groupby("Reporting Group")["Report Value"].sum().to_dict() if report_df is not None and not report_df.empty else {}
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
    return pnl_mapped.groupby(["Branch", "Reporting Group"], dropna=False)["Report Value"].sum().reset_index().rename(columns={"Report Value": "Actual"})


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
    return df.groupby(["Month", "Reporting Group"], dropna=False)["Report Value"].sum().reset_index().rename(columns={"Report Value": "Amount"}).sort_values(["Month", "Reporting Group"])


def build_monthly_branch_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])
    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])
    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    rev = df[df["Reporting Group"].astype(str).str.strip().str.lower() == "revenue"]
    return rev.groupby(["Month", "Branch"], dropna=False)["Report Value"].sum().reset_index().rename(columns={"Report Value": "Amount"}).sort_values(["Month", "Branch"])


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
    reporting_period = company_profile.get("Reporting Period", "unknown_period").strip().replace(" ", "_") or "unknown_period"
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
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped", "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary", "detected_branches", "validation_passed", "company_profile", "bs_disclaimer", "ai_commentary", "prior_pnl", "prior_bs", "prior_kpis", "save_run_preference", "anomaly_flags", "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df", "budget_compare", "budget_summary", "benchmark_df", "py_compare", "benchmark_compare", "monthly_actuals", "monthly_branch_actuals", "executive_summary_df", "forecast_pnl", "forecast_bs", "previous_year_pnl", "forecast_pnl_compare", "previous_year_pnl_compare", "fx_rate_info", "country_indicators", "external_benchmark_df", "consolidated_pnl_detail", "consolidated_bs_detail", "coa_duplicate_rows", "coa_mapping_review", "financial_logic_review", "last_validation_report", "reporting_structure"
]:
    if key not in st.session_state:
        st.session_state[key] = None
if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}
if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False

# ----------------------------
# UI
# ----------------------------
st.title("AI CFO Copilot")
st.caption("Automated branch-wise P&L, KPI packs, dashboarding, working capital, budget/forecast comparison, and AI insights")

tab_setup, tab_dashboard, tab_financials, tab_working_capital, tab_insights, tab_downloads = st.tabs(["Setup", "Dashboard", "Financials", "Working Capital", "Insights", "Downloads"])

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
            reporting_structure = st.radio(
                "Reporting Structure",
                ["Consolidated Only", "Branch / Business Unit Reporting"],
                index=0,
                help="If Consolidated Only is selected, Branch is optional in GL and the app uses a default Consolidated unit."
            )
            benchmark_group = st.text_input("Benchmark Group (Optional)")
        business_notes = st.text_area("Business Notes (Optional)")
        save_run_preference = st.checkbox("Save this run for future comparison", value=st.session_state["save_run_preference"])
        if st.button("Save Company Profile", use_container_width=True):
            if not company_name.strip():
                st.error("Company Name is mandatory.")
            elif industry == "Select Industry" or country == "Select Country":
                st.error("Please select at least Industry and Country.")
            else:
                st.session_state["company_profile"] = {"Company Name": company_name.strip(), "Industry": industry, "Country": country, "State / Region": state_region, "Financial Year": financial_year, "Currency": currency if currency != "Select Currency" else "", "Tax Identifier": tax_identifier, "Reporting Period": reporting_period, "Reporting Structure": reporting_structure, "Benchmark Group": benchmark_group, "Business Notes": business_notes}
                st.session_state["reporting_structure"] = reporting_structure
                st.session_state["save_run_preference"] = save_run_preference
                st.success("Company profile saved successfully.")
        if st.session_state["company_profile"]:
            st.dataframe(style_dataframe(pd.DataFrame(st.session_state["company_profile"].items(), columns=["Field", "Value"])), use_container_width=True)


    with st.expander("External FX & Benchmark Data", expanded=False):
        profile = st.session_state.get("company_profile", {})
        selected_country = profile.get("Country", "Australia") if profile else "Australia"
        selected_industry = profile.get("Industry", "Other") if profile else "Other"
        selected_currency = profile.get("Currency", "AUD") if profile else "AUD"
        default_target_currency = selected_currency if selected_currency and selected_currency != "Select Currency" else currency_for_country(selected_country)

        st.info("Optional: fetch live FX rates and country indicators. Industry benchmarks here are starter defaults; upload your own benchmark file to override them.")
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
                    if key == "coa":
                        duplicate_rows = find_duplicate_coa_rows(df)
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
            show_required_columns("Current GL Report", ["Account code", "Debit", "Credit"], ["Branch / Business Unit", "Net", "Date", "Description", "Account Name"])
            show_required_columns("COA Mapping", ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], ["Account Name", "Sign Convention", "Display Order"])
            show_required_columns("KPI Master", ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"], [])
            show_required_columns("Latest Previous Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Budget Data", ["Month", "Reporting Group", "Amount"], ["Branch / Business Unit"])
            show_required_columns("Forecast P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], [])
        with g2:
            show_required_columns("Forecast Balance Sheet", ["Reporting Group", "Reporting Subgroup", "Balance"], [])
            show_required_columns("Previous Year P&L", ["Reporting Group", "Reporting Subgroup", "Report Value"], [])
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

with tab_financials:
    st.subheader("Financials")
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

with tab_insights:
    st.subheader("Insights")
    insight_anom, insight_mapping, insight_ai = st.tabs(["Anomalies", "Mapping Review", "AI Commentary"])
    with insight_anom:
        flags = st.session_state.get("anomaly_flags", [])
        if flags:
            for flag in flags:
                st.warning(flag)
        else:
            st.success("No major anomalies detected based on current rules.")
    with insight_mapping:
        mapping_review = st.session_state.get("coa_mapping_review")
        logic_review = st.session_state.get("financial_logic_review")
        if mapping_review is None or mapping_review.empty:
            st.success("No obvious COA mapping issues detected by keyword rules.")
        else:
            st.warning("Review these mappings. The system only suggests; it does not reclassify automatically.")
            st.dataframe(style_dataframe(mapping_review), use_container_width=True)
        if logic_review is not None and not logic_review.empty:
            st.markdown("### Financial Logic Review")
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

with tab_downloads:
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
