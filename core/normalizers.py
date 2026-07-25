import pandas as pd
from core.common import clean_columns, validate_required_columns

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

