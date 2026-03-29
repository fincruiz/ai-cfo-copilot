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
        raise ValueError(f"{file_label} is missing required columns: {', '.join(missing)}")


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

    templates["Forecast Data"] = pd.DataFrame([
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Revenue", "Amount": 98000},
        {"Month": "2026-01", "Branch": "Sydney", "Reporting Group": "Gross Profit", "Amount": 40000},
        {"Month": "2026-01", "Branch": "Melbourne", "Reporting Group": "Revenue", "Amount": 82000},
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

    templates["Prior Period KPI Pack"] = pd.DataFrame([
        {"KPI": "Revenue", "Value": 22000, "Display Value": 22000, "Output Type": "value"},
        {"KPI": "Gross Margin %", "Value": 68.18, "Display Value": "68.18%", "Output Type": "percent"},
        {"KPI": "Operating Margin %", "Value": 31.82, "Display Value": "31.82%", "Output Type": "percent"},
    ])

    return templates


# ----------------------------
# Standardization helpers
# ----------------------------
def standardize_key_columns(gl, coa, kpi=None, latest_bs=None):
    gl = clean_columns(gl)
    coa = clean_columns(coa)

    gl.rename(columns={
        "Account Code": "Account code",
        "account code": "Account code",
        "Account code ": "Account code",
        "branch": "Branch",
        "net": "Net",
        "Debit ": "Debit",
        "Credit ": "Credit",
        "Description ": "Description",
        "Date ": "Date",
        "Posting Date": "Date",
        "Txn Date": "Date",
    }, inplace=True)

    coa.rename(columns={
        "Account Code": "Account code",
        "account code": "Account code",
        "Reporting group": "Reporting Group",
        "Reporting subgroup": "Reporting Subgroup",
        "Sign convention": "Sign Convention",
        "Statement type": "Statement",
    }, inplace=True)

    if kpi is not None:
        kpi = clean_columns(kpi)
        kpi.rename(columns={
            "Formula type": "Formula Type",
            "Numerator group": "Numerator Group",
            "Denominator group": "Denominator Group",
            "Output type": "Output Type",
            "Display order": "Display Order",
            "Kpi name": "KPI Name",
        }, inplace=True)

    if latest_bs is not None:
        latest_bs = clean_columns(latest_bs)
        latest_bs.rename(columns={
            "Balance ": "Balance",
            "Reporting group": "Reporting Group",
            "Reporting subgroup": "Reporting Subgroup",
        }, inplace=True)

    return gl, coa, kpi, latest_bs


def normalize_plan_df(df: pd.DataFrame, label: str) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={
        "Month ": "Month",
        "Branch ": "Branch",
        "Reporting group": "Reporting Group",
        "Amount ": "Amount",
        "Budget Amount": "Amount",
        "Forecast Amount": "Amount",
    }, inplace=True)

    validate_required_columns(df, ["Month", "Branch", "Reporting Group", "Amount"], label)
    df["Month"] = df["Month"].astype(str).str.strip()
    df["Branch"] = df["Branch"].astype(str).str.strip()
    df["Reporting Group"] = df["Reporting Group"].astype(str).str.strip()
    df["Amount"] = pd.to_numeric(df["Amount"], errors="coerce").fillna(0)
    return df


def normalize_benchmark_df(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    df.rename(columns={
        "Metric ": "Metric",
        "Benchmark": "Benchmark Value",
        "Benchmark %": "Benchmark Value",
        "Benchmark Value ": "Benchmark Value",
    }, inplace=True)

    validate_required_columns(df, ["Metric", "Benchmark Value"], "Benchmark Data")
    df["Metric"] = df["Metric"].astype(str).str.strip()
    df["Benchmark Value"] = pd.to_numeric(df["Benchmark Value"], errors="coerce").fillna(0)
    return df


def normalize_ageing_df(df: pd.DataFrame, kind: str) -> pd.DataFrame:
    df = clean_columns(df)
    rename_map = {
        "Customer": "Party Name",
        "Customer Name": "Party Name",
        "Supplier": "Party Name",
        "Supplier Name": "Party Name",
        "Vendor": "Party Name",
        "Vendor Name": "Party Name",
        "Invoice Number": "Document Number",
        "Bill Number": "Document Number",
        "Invoice No": "Document Number",
        "Bill No": "Document Number",
        "Outstanding": "Outstanding Amount",
        "Outstanding Balance": "Outstanding Amount",
        "Amount": "Outstanding Amount",
        "Due Date ": "Due Date",
        "Invoice Date ": "Document Date",
        "Bill Date": "Document Date",
        "Ageing Bucket": "Age Bucket",
        "Aging Bucket": "Age Bucket",
        "Age Bucket ": "Age Bucket",
        "Branch ": "Branch",
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
        elif days_overdue <= 30:
            return "1-30"
        elif days_overdue <= 60:
            return "31-60"
        elif days_overdue <= 90:
            return "61-90"
        else:
            return "90+"

    df["Age Bucket"] = df.apply(calc_bucket, axis=1)
    df["Branch"] = df["Branch"].astype(str).str.strip()
    return df


# ----------------------------
# Core finance helpers
# ----------------------------
def apply_sign_convention_to_gl(row) -> float:
    net = row["Net"]
    sign = str(row.get("Sign Convention", "")).strip().lower()

    if pd.isna(net):
        return 0.0

    abs_net = abs(float(net))
    if sign == "negative":
        return -abs_net
    return abs_net


def build_pnl(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Report Value"])

    return (
        report_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .sort_values(["Reporting Group", "Reporting Subgroup"])
    )


def build_balance_sheet_from_gl(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    return (
        bs_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Balance"})
        .sort_values(["Reporting Group", "Reporting Subgroup"])
    )


def combine_opening_and_current_bs(opening_bs: pd.DataFrame, current_bs: pd.DataFrame) -> pd.DataFrame:
    if opening_bs is None or opening_bs.empty:
        return current_bs.copy()

    opening = opening_bs.copy()
    current = current_bs.copy()

    opening["Balance"] = pd.to_numeric(opening["Balance"], errors="coerce").fillna(0)
    current["Balance"] = pd.to_numeric(current["Balance"], errors="coerce").fillna(0)

    merged = opening.merge(
        current,
        on=["Reporting Group", "Reporting Subgroup"],
        how="outer",
        suffixes=("_opening", "_current"),
    ).fillna(0)

    merged["Balance"] = merged["Balance_opening"] + merged["Balance_current"]
    return merged[["Reporting Group", "Reporting Subgroup", "Balance"]].sort_values(
        ["Reporting Group", "Reporting Subgroup"]
    ).reset_index(drop=True)


def build_kpis(report_df: pd.DataFrame, kpi_master: pd.DataFrame) -> pd.DataFrame:
    group_values = report_df.groupby("Reporting Group")["Report Value"].sum().to_dict()
    results = []
    calculated = {}

    kpi_master = kpi_master.sort_values("Display Order").copy()

    for _, row in kpi_master.iterrows():
        kpi_name = str(row["KPI Name"]).strip()
        formula_type = str(row["Formula Type"]).strip().lower()
        numerator = str(row["Numerator Group"]).strip() if pd.notna(row["Numerator Group"]) else ""
        denominator = str(row["Denominator Group"]).strip() if pd.notna(row["Denominator Group"]) else ""
        output_type = str(row["Output Type"]).strip().lower()

        value = 0.0

        if formula_type == "direct":
            value = group_values.get(numerator, 0.0)
        elif formula_type == "derived":
            value = calculated.get(numerator, group_values.get(numerator, 0.0)) - \
                    calculated.get(denominator, group_values.get(denominator, 0.0))
        elif formula_type == "ratio":
            num_val = calculated.get(numerator, group_values.get(numerator, 0.0))
            den_val = calculated.get(denominator, group_values.get(denominator, 0.0))
            value = (num_val / den_val * 100) if den_val != 0 else 0.0

        calculated[kpi_name] = value
        results.append({"KPI": kpi_name, "Value": value, "Output Type": output_type})

    kpi_df = pd.DataFrame(results)
    kpi_df["Display Value"] = kpi_df.apply(
        lambda r: f"{r['Value']:.2f}%" if r["Output Type"] == "percent" else round(r["Value"], 2),
        axis=1,
    )
    return kpi_df[["KPI", "Value", "Output Type", "Display Value"]]


def kpi_map_from_df(kpi_df: pd.DataFrame | None) -> dict:
    if kpi_df is None or kpi_df.empty:
        return {}
    return {row["KPI"]: row["Value"] for _, row in kpi_df.iterrows()}


def build_actuals_by_branch_reporting_group(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty:
        return pd.DataFrame(columns=["Branch", "Reporting Group", "Actual"])

    return (
        pnl_mapped.groupby(["Branch", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Actual"})
    )


def compare_plan_vs_actual(actuals_df: pd.DataFrame, plan_df: pd.DataFrame, label: str) -> pd.DataFrame:
    if plan_df is None or plan_df.empty:
        return pd.DataFrame(columns=["Branch", "Reporting Group", "Actual", label, "Variance", "Variance %"])

    plan_agg = (
        plan_df.groupby(["Branch", "Reporting Group"], dropna=False)["Amount"]
        .sum()
        .reset_index()
        .rename(columns={"Amount": label})
    )

    merged = actuals_df.merge(plan_agg, on=["Branch", "Reporting Group"], how="outer").fillna(0)
    merged["Variance"] = merged["Actual"] - merged[label]
    merged["Variance %"] = merged.apply(
        lambda r: (r["Variance"] / r[label] * 100) if r[label] != 0 else 0.0,
        axis=1,
    )
    return merged.sort_values(["Branch", "Reporting Group"]).reset_index(drop=True)


def summarize_plan_vs_actual(compare_df: pd.DataFrame, label: str) -> pd.DataFrame:
    if compare_df is None or compare_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Actual", label, "Variance", "Variance %"])

    out = (
        compare_df.groupby("Reporting Group", dropna=False)[["Actual", label, "Variance"]]
        .sum()
        .reset_index()
    )
    out["Variance %"] = out.apply(
        lambda r: (r["Variance"] / r[label] * 100) if r[label] != 0 else 0.0,
        axis=1,
    )
    return out.sort_values("Reporting Group").reset_index(drop=True)


def build_ageing_summary(df: pd.DataFrame | None, kind: str) -> dict:
    if df is None or df.empty:
        return {
            "total": 0.0,
            "overdue": 0.0,
            "overdue_pct": 0.0,
            "by_bucket": pd.DataFrame(columns=["Age Bucket", "Outstanding Amount"]),
            "by_branch": pd.DataFrame(columns=["Branch", "Outstanding Amount"]),
            "top_parties": pd.DataFrame(columns=["Party Name", "Outstanding Amount"]),
            "kind": kind,
        }

    total = float(df["Outstanding Amount"].sum())
    overdue_df = df[df["Age Bucket"].isin(["1-30", "31-60", "61-90", "90+"])]
    overdue = float(overdue_df["Outstanding Amount"].sum())
    overdue_pct = (overdue / total * 100) if total != 0 else 0.0

    bucket_order = ["Current", "1-30", "31-60", "61-90", "90+", "Unknown"]

    by_bucket = df.groupby("Age Bucket", dropna=False)["Outstanding Amount"].sum().reset_index()
    by_bucket["Age Bucket"] = pd.Categorical(by_bucket["Age Bucket"], categories=bucket_order, ordered=True)
    by_bucket = by_bucket.sort_values("Age Bucket")

    by_branch = (
        df.groupby("Branch", dropna=False)["Outstanding Amount"]
        .sum()
        .reset_index()
        .sort_values("Outstanding Amount", ascending=False)
    )

    top_parties = (
        df.groupby("Party Name", dropna=False)["Outstanding Amount"]
        .sum()
        .reset_index()
        .sort_values("Outstanding Amount", ascending=False)
        .head(10)
    )

    return {
        "total": total,
        "overdue": overdue,
        "overdue_pct": overdue_pct,
        "by_bucket": by_bucket,
        "by_branch": by_branch,
        "top_parties": top_parties,
        "kind": kind,
    }


# ----------------------------
# Monthly trends / executive summary helpers
# ----------------------------
def build_monthly_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Reporting Group", "Amount"])

    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["Month", "Reporting Group", "Amount"])

    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    out = (
        df.groupby(["Month", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Reporting Group"])
    )
    return out


def build_monthly_branch_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()

    if df.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    df["Month"] = df["Date"].dt.to_period("M").astype(str)
    revenue_df = df[df["Reporting Group"].astype(str).str.strip().str.lower() == "revenue"].copy()

    out = (
        revenue_df.groupby(["Month", "Branch"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Branch"])
    )
    return out


def build_py_comparison(current_kpis: pd.DataFrame | None, prior_kpis: pd.DataFrame | None) -> pd.DataFrame:
    if current_kpis is None or current_kpis.empty or prior_kpis is None or prior_kpis.empty:
        return pd.DataFrame(columns=["Metric", "Current", "Prior Year", "Variance", "Variance %"])

    cur = current_kpis[["KPI", "Value"]].rename(columns={"KPI": "Metric", "Value": "Current"})
    py = prior_kpis[["KPI", "Value"]].rename(columns={"KPI": "Metric", "Value": "Prior Year"})

    merged = cur.merge(py, on="Metric", how="inner")
    merged["Variance"] = merged["Current"] - merged["Prior Year"]
    merged["Variance %"] = merged.apply(
        lambda r: (r["Variance"] / r["Prior Year"] * 100) if r["Prior Year"] != 0 else 0.0,
        axis=1,
    )
    return merged


def build_benchmark_comparison(current_kpis: pd.DataFrame | None, benchmark_df: pd.DataFrame | None,
                               ar_summary=None, ap_summary=None) -> pd.DataFrame:
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
    return merged.sort_values("Metric").reset_index(drop=True)


def rag_status(metric_name: str, current_value: float, benchmark_value=None) -> str:
    metric_name = str(metric_name).strip().lower()

    if benchmark_value not in [None, ""]:
        benchmark_value = safe_float(benchmark_value)
        gap = current_value - benchmark_value
        if "margin" in metric_name:
            if gap >= 0:
                return "Green"
            elif gap >= -3:
                return "Amber"
            return "Red"
        if "overdue" in metric_name:
            if gap <= 0:
                return "Green"
            elif gap <= 5:
                return "Amber"
            return "Red"

    if "gross margin" in metric_name:
        if current_value >= 25:
            return "Green"
        elif current_value >= 18:
            return "Amber"
        return "Red"

    if "operating margin" in metric_name:
        if current_value >= 10:
            return "Green"
        elif current_value >= 5:
            return "Amber"
        return "Red"

    if "opex" in metric_name:
        if current_value <= 25:
            return "Green"
        elif current_value <= 35:
            return "Amber"
        return "Red"

    if "overdue" in metric_name:
        if current_value <= 20:
            return "Green"
        elif current_value <= 35:
            return "Amber"
        return "Red"

    return "Amber"


def build_executive_summary(current_kpis, ar_summary=None, ap_summary=None, budget_summary=None, forecast_summary=None,
                            benchmark_compare=None) -> pd.DataFrame:
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

    if forecast_summary is not None and not forecast_summary.empty and forecast_summary["Forecast"].sum() != 0:
        total_variance_pct = forecast_summary["Variance"].sum() / forecast_summary["Forecast"].sum() * 100
        rows.append({
            "Metric": "Forecast Variance %",
            "Current Value": total_variance_pct,
            "Benchmark Value": "",
            "Status": "Green" if total_variance_pct >= 0 else ("Amber" if total_variance_pct >= -10 else "Red"),
        })

    return pd.DataFrame(rows)


# ----------------------------
# Pack creation
# ----------------------------
def create_excel_pack(
    consolidated_pnl,
    consolidated_bs,
    consolidated_kpis,
    branch_summary,
    branch_outputs,
    unmapped,
    executive_summary=None,
    monthly_actuals=None,
    monthly_branch_actuals=None,
    ar_df=None,
    ap_df=None,
    budget_compare=None,
    forecast_compare=None,
    py_compare=None,
    benchmark_compare=None,
):
    df_dict = {
        "Executive Summary": executive_summary if executive_summary is not None else pd.DataFrame(),
        "Consolidated P&L": consolidated_pnl,
    }

    if consolidated_bs is not None and not consolidated_bs.empty:
        df_dict["Consolidated BS"] = consolidated_bs
    if consolidated_kpis is not None:
        df_dict["Consolidated KPIs"] = consolidated_kpis
    if branch_summary is not None and not branch_summary.empty:
        df_dict["Branch Summary KPIs"] = branch_summary
    if monthly_actuals is not None and not monthly_actuals.empty:
        df_dict["Monthly Trends"] = monthly_actuals
    if monthly_branch_actuals is not None and not monthly_branch_actuals.empty:
        df_dict["Branch Monthly Trends"] = monthly_branch_actuals

    for branch, reports in branch_outputs.items():
        df_dict[f"{branch} P&L"] = reports["pnl"]
        if reports["kpis"] is not None:
            df_dict[f"{branch} KPIs"] = reports["kpis"]

    if not unmapped.empty:
        df_dict["Unmapped Accounts"] = unmapped
    if ar_df is not None and not ar_df.empty:
        df_dict["AR Ageing"] = ar_df
    if ap_df is not None and not ap_df.empty:
        df_dict["AP Ageing"] = ap_df
    if budget_compare is not None and not budget_compare.empty:
        df_dict["Budget vs Actual"] = budget_compare
    if forecast_compare is not None and not forecast_compare.empty:
        df_dict["Forecast vs Actual"] = forecast_compare
    if py_compare is not None and not py_compare.empty:
        df_dict["Actual vs PY"] = py_compare
    if benchmark_compare is not None and not benchmark_compare.empty:
        df_dict["Benchmark Comparison"] = benchmark_compare

    return dataframe_to_excel_bytes(df_dict)


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
# AI / anomalies
# ----------------------------
def detect_anomalies(
    consolidated_kpis,
    branch_outputs,
    prior_kpis=None,
    ar_summary=None,
    ap_summary=None,
    budget_summary=None,
    forecast_summary=None,
):
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

    if forecast_summary is not None and not forecast_summary.empty and forecast_summary["Forecast"].sum() != 0:
        total_variance_pct = forecast_summary["Variance"].sum() / forecast_summary["Forecast"].sum() * 100
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
    forecast_summary=None,
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

        forecast_text = "No forecast data available."
        if forecast_summary is not None and not forecast_summary.empty:
            forecast_text = forecast_summary.to_string(index=False)[:1500]

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


def build_prior_period_from_gl(prior_gl_file, coa, kpi_master):
    prior_gl = pd.read_excel(prior_gl_file)
    prior_gl = clean_columns(prior_gl)
    prior_gl.rename(columns={
        "Account Code": "Account code",
        "account code": "Account code",
        "branch": "Branch",
        "net": "Net",
    }, inplace=True)

    validate_required_columns(prior_gl, ["Account code", "Debit", "Credit", "Branch"], "Prior Period GL")

    prior_gl["Account code"] = prior_gl["Account code"].astype(str).str.strip()
    prior_gl["Debit"] = pd.to_numeric(prior_gl["Debit"], errors="coerce").fillna(0)
    prior_gl["Credit"] = pd.to_numeric(prior_gl["Credit"], errors="coerce").fillna(0)

    if "Net" not in prior_gl.columns:
        prior_gl["Net"] = prior_gl["Debit"] - prior_gl["Credit"]
    else:
        prior_gl["Net"] = pd.to_numeric(prior_gl["Net"], errors="coerce").fillna(prior_gl["Debit"] - prior_gl["Credit"])

    merged = prior_gl.merge(coa, on="Account code", how="left")
    merged = merged[merged["Reporting Group"].notna()].copy()

    if "Sign Convention" not in merged.columns:
        merged["Sign Convention"] = "positive"

    merged["Report Value"] = merged.apply(apply_sign_convention_to_gl, axis=1)

    prior_pnl = build_pnl(merged[merged["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy())
    prior_bs = build_balance_sheet_from_gl(merged[merged["Statement"].astype(str).str.strip().str.lower() == "balance sheet"].copy())
    prior_kpis = build_kpis(
        merged[merged["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy(),
        kpi_master
    ) if kpi_master is not None else None

    return prior_pnl, prior_bs, prior_kpis


# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped",
    "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs", "branch_summary",
    "detected_branches", "validation_passed", "company_profile", "bs_disclaimer", "ai_commentary",
    "prior_pnl", "prior_bs", "prior_kpis", "save_run_preference", "anomaly_flags",
    "ar_df", "ap_df", "ar_summary", "ap_summary", "budget_df", "forecast_df",
    "budget_compare", "forecast_compare", "budget_summary", "forecast_summary",
    "benchmark_df", "py_compare", "benchmark_compare", "monthly_actuals", "monthly_branch_actuals",
    "executive_summary_df"
]:
    if key not in st.session_state:
        st.session_state[key] = None

if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}
if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False


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
            forecast_file = st.file_uploader("Forecast Data (Optional)", type=["xlsx"])
        with c3:
            ar_file = st.file_uploader("AR Ageing (Optional)", type=["xlsx"])
            ap_file = st.file_uploader("AP Ageing (Optional)", type=["xlsx"])
            benchmark_file = st.file_uploader("Industry Benchmark File (Optional)", type=["xlsx"])

        if st.button("Validate & Load Current Files", use_container_width=True):
            try:
                profile = st.session_state["company_profile"]
                if not profile or not profile.get("Company Name", "").strip():
                    st.error("Please save Company Profile first. Company Name is mandatory.")
                elif not (gl_file and mapping_file):
                    st.error("Please upload Current GL Report and COA Mapping.")
                else:
                    gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
                        gl_file, mapping_file, kpi_file, latest_bs_file
                    )

                    consolidated_pnl = build_pnl(pnl_mapped)
                    current_bs = build_balance_sheet_from_gl(bs_mapped)

                    bs_disclaimer = None
                    if latest_bs is not None:
                        consolidated_bs = combine_opening_and_current_bs(latest_bs, current_bs)
                    else:
                        consolidated_bs = current_bs
                        bs_disclaimer = "Balance Sheet may not fully match because opening balances were not provided."

                    consolidated_kpis = build_kpis(pnl_mapped, kpi_master) if kpi_master is not None else None

                    detected_branches = sorted(pnl_mapped["Branch"].dropna().unique().tolist())

                    branch_outputs = {}
                    branch_summary_rows = []
                    for branch in detected_branches:
                        branch_df = pnl_mapped[pnl_mapped["Branch"] == branch].copy()
                        branch_pnl = build_pnl(branch_df)
                        branch_kpis = build_kpis(branch_df, kpi_master) if kpi_master is not None else None
                        branch_outputs[branch] = {"pnl": branch_pnl, "kpis": branch_kpis}

                        if branch_kpis is not None:
                            row = {"Branch": branch}
                            for _, r in branch_kpis.iterrows():
                                row[r["KPI"]] = r["Display Value"]
                            branch_summary_rows.append(row)

                    branch_summary = pd.DataFrame(branch_summary_rows) if branch_summary_rows else pd.DataFrame()

                    ar_df = normalize_ageing_df(pd.read_excel(ar_file), "AR") if ar_file is not None else None
                    ap_df = normalize_ageing_df(pd.read_excel(ap_file), "AP") if ap_file is not None else None
                    ar_summary = build_ageing_summary(ar_df, "AR") if ar_df is not None else None
                    ap_summary = build_ageing_summary(ap_df, "AP") if ap_df is not None else None

                    budget_df = normalize_plan_df(pd.read_excel(budget_file), "Budget Data") if budget_file is not None else None
                    forecast_df = normalize_plan_df(pd.read_excel(forecast_file), "Forecast Data") if forecast_file is not None else None
                    benchmark_df = normalize_benchmark_df(pd.read_excel(benchmark_file)) if benchmark_file is not None else None

                    actuals_df = build_actuals_by_branch_reporting_group(pnl_mapped)
                    budget_compare = compare_plan_vs_actual(actuals_df, budget_df, "Budget") if budget_df is not None else None
                    forecast_compare = compare_plan_vs_actual(actuals_df, forecast_df, "Forecast") if forecast_df is not None else None
                    budget_summary = summarize_plan_vs_actual(budget_compare, "Budget") if budget_compare is not None else None
                    forecast_summary = summarize_plan_vs_actual(forecast_compare, "Forecast") if forecast_compare is not None else None

                    py_compare = build_py_comparison(consolidated_kpis, st.session_state.get("prior_kpis"))
                    benchmark_compare = build_benchmark_comparison(consolidated_kpis, benchmark_df, ar_summary, ap_summary)

                    monthly_actuals = build_monthly_actuals(pnl_mapped)
                    monthly_branch_actuals = build_monthly_branch_actuals(pnl_mapped)

                    executive_summary_df = build_executive_summary(
                        consolidated_kpis,
                        ar_summary=ar_summary,
                        ap_summary=ap_summary,
                        budget_summary=budget_summary,
                        forecast_summary=forecast_summary,
                        benchmark_compare=benchmark_compare,
                    )

                    st.session_state["gl"] = gl
                    st.session_state["coa"] = coa
                    st.session_state["kpi_master"] = kpi_master
                    st.session_state["latest_bs"] = latest_bs
                    st.session_state["mapped"] = mapped
                    st.session_state["pnl_mapped"] = pnl_mapped
                    st.session_state["bs_mapped"] = bs_mapped
                    st.session_state["unmapped"] = unmapped
                    st.session_state["consolidated_pnl"] = consolidated_pnl
                    st.session_state["consolidated_bs"] = consolidated_bs
                    st.session_state["consolidated_kpis"] = consolidated_kpis
                    st.session_state["branch_outputs"] = branch_outputs
                    st.session_state["branch_summary"] = branch_summary
                    st.session_state["detected_branches"] = detected_branches
                    st.session_state["validation_passed"] = unmapped.empty
                    st.session_state["bs_disclaimer"] = bs_disclaimer
                    st.session_state["ai_commentary"] = None
                    st.session_state["ar_df"] = ar_df
                    st.session_state["ap_df"] = ap_df
                    st.session_state["ar_summary"] = ar_summary
                    st.session_state["ap_summary"] = ap_summary
                    st.session_state["budget_df"] = budget_df
                    st.session_state["forecast_df"] = forecast_df
                    st.session_state["benchmark_df"] = benchmark_df
                    st.session_state["budget_compare"] = budget_compare
                    st.session_state["forecast_compare"] = forecast_compare
                    st.session_state["budget_summary"] = budget_summary
                    st.session_state["forecast_summary"] = forecast_summary
                    st.session_state["py_compare"] = py_compare
                    st.session_state["benchmark_compare"] = benchmark_compare
                    st.session_state["monthly_actuals"] = monthly_actuals
                    st.session_state["monthly_branch_actuals"] = monthly_branch_actuals
                    st.session_state["executive_summary_df"] = executive_summary_df

                    if st.session_state["save_run_preference"]:
                        save_run_to_history(
                            st.session_state["company_profile"],
                            consolidated_pnl,
                            consolidated_bs,
                            consolidated_kpis,
                            branch_summary,
                        )

                    st.session_state["anomaly_flags"] = detect_anomalies(
                        consolidated_kpis,
                        branch_outputs,
                        prior_kpis=st.session_state.get("prior_kpis"),
                        ar_summary=ar_summary,
                        ap_summary=ap_summary,
                        budget_summary=budget_summary,
                        forecast_summary=forecast_summary,
                    ) if consolidated_kpis is not None else []

                    if unmapped.empty:
                        st.success("Files validated and loaded successfully.")
                    else:
                        st.warning("Files loaded, but unmapped GL rows were found.")

            except Exception as e:
                st.error(f"Error: {e}")

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
                coa = st.session_state.get("coa")
                kpi_master = st.session_state.get("kpi_master")
                loaded_any = False

                if prior_gl_file is not None:
                    if coa is None:
                        st.error("Load current files first so COA mapping is available.")
                    else:
                        prior_pnl, prior_bs, prior_kpis = build_prior_period_from_gl(prior_gl_file, coa, kpi_master)
                        st.session_state["prior_pnl"] = prior_pnl
                        st.session_state["prior_bs"] = prior_bs
                        st.session_state["prior_kpis"] = prior_kpis
                        loaded_any = True
                else:
                    if prior_pnl_file is not None:
                        st.session_state["prior_pnl"] = clean_columns(pd.read_excel(prior_pnl_file))
                        loaded_any = True
                    if prior_bs_file is not None:
                        st.session_state["prior_bs"] = clean_columns(pd.read_excel(prior_bs_file))
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

        with g2:
            show_required_columns("Forecast Data", ["Month", "Branch", "Reporting Group", "Amount"], [])
            show_required_columns("AR Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("AP Ageing", ["Party Name", "Outstanding Amount"], ["Document Number", "Document Date", "Due Date", "Branch", "Age Bucket"])
            show_required_columns("Industry Benchmark File", ["Metric", "Benchmark Value"], [])
            show_required_columns("Prior Period GL Report", ["Account code", "Debit", "Credit", "Branch"], ["Net", "Date", "Description"])
            show_required_columns("Prior KPI Pack", ["KPI", "Value"], ["Display Value", "Output Type"])

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

        if st.session_state["forecast_summary"] is not None and not st.session_state["forecast_summary"].empty:
            st.markdown("**Forecast vs Actual**")
            st.bar_chart(st.session_state["forecast_summary"].set_index("Reporting Group")[["Actual", "Forecast"]])

        if st.session_state["py_compare"] is not None and not st.session_state["py_compare"].empty:
            st.markdown("**Actual vs Prior Year**")
            py_chart = st.session_state["py_compare"].copy().set_index("Metric")[["Current", "Prior Year"]]
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

    with sub_bs:
        if st.session_state["consolidated_bs"] is None or st.session_state["consolidated_bs"].empty:
            st.info("No Balance Sheet available yet.")
        else:
            if st.session_state["bs_disclaimer"]:
                st.warning(st.session_state["bs_disclaimer"])
            st.dataframe(style_dataframe(st.session_state["consolidated_bs"]), use_container_width=True)

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

        if st.session_state["forecast_compare"] is not None and not st.session_state["forecast_compare"].empty:
            st.markdown("### Forecast vs Actual")
            st.dataframe(style_dataframe(st.session_state["forecast_summary"]), use_container_width=True)
            st.dataframe(style_dataframe(st.session_state["forecast_compare"]), use_container_width=True)
        else:
            st.info("No forecast data uploaded.")

        if st.session_state["py_compare"] is not None and not st.session_state["py_compare"].empty:
            st.markdown("### Actual vs Prior Year")
            st.dataframe(style_dataframe(st.session_state["py_compare"]), use_container_width=True)

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
                        forecast_summary=st.session_state.get("forecast_summary"),
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
            forecast_compare=st.session_state["forecast_compare"],
            py_compare=st.session_state["py_compare"],
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
