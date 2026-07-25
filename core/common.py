import re
import pandas as pd

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

