import streamlit as st
import pandas as pd
from io import BytesIO
from openai import OpenAI
import os
import re
from pathlib import Path

st.set_page_config(page_title="AI CFO Copilot", layout="wide")


# ----------------------------
# Config / Paths
# ----------------------------
HISTORY_ROOT = Path("history")
HISTORY_ROOT.mkdir(exist_ok=True)


# ----------------------------
# Helpers
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


def standardize_key_columns(
    gl: pd.DataFrame,
    coa: pd.DataFrame,
    kpi: pd.DataFrame | None = None,
    latest_bs: pd.DataFrame | None = None,
    prior_pnl: pd.DataFrame | None = None,
    prior_bs: pd.DataFrame | None = None,
    prior_kpi: pd.DataFrame | None = None,
):
    gl = clean_columns(gl)
    coa = clean_columns(coa)

    gl.rename(
        columns={
            "Account Code": "Account code",
            "account code": "Account code",
            "Account code ": "Account code",
            "branch": "Branch",
            "net": "Net",
            "Debit ": "Debit",
            "Credit ": "Credit",
            "Description ": "Description",
            "Date ": "Date",
        },
        inplace=True,
    )

    coa.rename(
        columns={
            "Account Code": "Account code",
            "account code": "Account code",
            "Reporting group": "Reporting Group",
            "Reporting subgroup": "Reporting Subgroup",
            "Sign convention": "Sign Convention",
            "Statement type": "Statement",
        },
        inplace=True,
    )

    if kpi is not None:
        kpi = clean_columns(kpi)
        kpi.rename(
            columns={
                "Formula type": "Formula Type",
                "Numerator group": "Numerator Group",
                "Denominator group": "Denominator Group",
                "Output type": "Output Type",
                "Display order": "Display Order",
                "Kpi name": "KPI Name",
            },
            inplace=True,
        )

    if latest_bs is not None:
        latest_bs = clean_columns(latest_bs)
        latest_bs.rename(
            columns={
                "Balance ": "Balance",
                "Reporting group": "Reporting Group",
                "Reporting subgroup": "Reporting Subgroup",
            },
            inplace=True,
        )

    if prior_pnl is not None:
        prior_pnl = clean_columns(prior_pnl)

    if prior_bs is not None:
        prior_bs = clean_columns(prior_bs)
        prior_bs.rename(
            columns={
                "Reporting group": "Reporting Group",
                "Reporting subgroup": "Reporting Subgroup",
                "Balance ": "Balance",
            },
            inplace=True,
        )

    if prior_kpi is not None:
        prior_kpi = clean_columns(prior_kpi)
        prior_kpi.rename(
            columns={
                "Kpi": "KPI",
                "Display value": "Display Value",
            },
            inplace=True,
        )

    return gl, coa, kpi, latest_bs, prior_pnl, prior_bs, prior_kpi


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], file_label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"{file_label} is missing required columns: {', '.join(missing)}")


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
    pnl = (
        report_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .sort_values(["Reporting Group", "Reporting Subgroup"])
    )
    return pnl


def build_balance_sheet_from_gl(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    out = (
        bs_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Balance"})
        .sort_values(["Reporting Group", "Reporting Subgroup"])
    )
    return out


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

    out = merged[["Reporting Group", "Reporting Subgroup", "Balance"]].copy()
    out = out.sort_values(["Reporting Group", "Reporting Subgroup"]).reset_index(drop=True)
    return out


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
            num_val = calculated.get(numerator, group_values.get(numerator, 0.0))
            den_val = calculated.get(denominator, group_values.get(denominator, 0.0))
            value = num_val - den_val

        elif formula_type == "ratio":
            num_val = calculated.get(numerator, group_values.get(numerator, 0.0))
            den_val = calculated.get(denominator, group_values.get(denominator, 0.0))
            value = (num_val / den_val * 100) if den_val != 0 else 0.0

        calculated[kpi_name] = value

        results.append(
            {
                "KPI": kpi_name,
                "Value": value,
                "Output Type": output_type,
            }
        )

    kpi_df = pd.DataFrame(results)
    kpi_df["Display Value"] = kpi_df.apply(
        lambda r: f"{r['Value']:.2f}%" if r["Output Type"] == "percent" else round(r["Value"], 2),
        axis=1,
    )

    return kpi_df[["KPI", "Value", "Output Type", "Display Value"]]


def dataframe_to_excel_bytes(df_dict: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in df_dict.items():
            safe_sheet = str(sheet_name)[:31]
            df.to_excel(writer, sheet_name=safe_sheet, index=False)
    return output.getvalue()


def create_excel_pack(consolidated_pnl, consolidated_bs, consolidated_kpis, branch_summary, branch_outputs, unmapped):
    df_dict = {"Consolidated P&L": consolidated_pnl}

    if consolidated_bs is not None and not consolidated_bs.empty:
        df_dict["Consolidated BS"] = consolidated_bs

    if consolidated_kpis is not None:
        df_dict["Consolidated KPIs"] = consolidated_kpis

    if branch_summary is not None and not branch_summary.empty:
        df_dict["Branch Summary KPIs"] = branch_summary

    for branch, reports in branch_outputs.items():
        df_dict[f"{branch} P&L"] = reports["pnl"]
        if reports["kpis"] is not None:
            df_dict[f"{branch} KPIs"] = reports["kpis"]

    if not unmapped.empty:
        df_dict["Unmapped Accounts"] = unmapped

    return dataframe_to_excel_bytes(df_dict)


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

    runs = []
    for item in company_folder.iterdir():
        if item.is_dir():
            runs.append(item.name)

    return sorted(runs, reverse=True)


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


def detect_anomalies(consolidated_kpis, branch_outputs, prior_kpis=None):
    flags = []

    current_kpi_map = {}
    if consolidated_kpis is not None and not consolidated_kpis.empty:
        for _, row in consolidated_kpis.iterrows():
            current_kpi_map[row["KPI"]] = row["Value"]

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

    branch_margins = []
    for branch, reports in branch_outputs.items():
        if reports["kpis"] is not None and not reports["kpis"].empty:
            branch_kpi_map = {}
            for _, row in reports["kpis"].iterrows():
                branch_kpi_map[row["KPI"]] = row["Value"]

            gm = branch_kpi_map.get("Gross Margin %", None)
            opm = branch_kpi_map.get("Operating Margin %", None)

            if gm is not None:
                branch_margins.append((branch, gm))

            if opm is not None and opm < 0:
                flags.append(f"{branch} has negative operating margin ({opm:.2f}%).")

    if len(branch_margins) >= 2:
        avg_margin = sum(x[1] for x in branch_margins) / len(branch_margins)
        for branch, gm in branch_margins:
            if gm < avg_margin - 10:
                flags.append(f"{branch} gross margin ({gm:.2f}%) is materially below branch average ({avg_margin:.2f}%).")

    if prior_kpis is not None and not prior_kpis.empty and "KPI" in prior_kpis.columns and "Value" in prior_kpis.columns:
        prior_kpi_map = {}
        for _, row in prior_kpis.iterrows():
            prior_kpi_map[row["KPI"]] = row["Value"]

        prior_revenue = prior_kpi_map.get("Revenue", None)
        prior_gm = prior_kpi_map.get("Gross Margin %", None)
        prior_opm = prior_kpi_map.get("Operating Margin %", None)

        if prior_revenue not in (None, 0):
            revenue_change_pct = ((revenue - prior_revenue) / prior_revenue) * 100
            if revenue_change_pct < -10:
                flags.append(f"Revenue declined {revenue_change_pct:.2f}% versus prior period.")

        if prior_gm is not None:
            gm_change = gross_margin - prior_gm
            if gm_change < -3:
                flags.append(f"Gross margin dropped by {gm_change:.2f} percentage points versus prior period.")

        if prior_opm is not None:
            opm_change = operating_margin - prior_opm
            if opm_change < -3:
                flags.append(f"Operating margin dropped by {opm_change:.2f} percentage points versus prior period.")

    return flags


def generate_ai_commentary(pnl_df, kpi_df, bs_df, profile, anomaly_flags=None):
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
        validate_required_columns(
            latest_bs,
            ["Reporting Group", "Reporting Subgroup", "Balance"],
            "Latest Balance Sheet",
        )

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
    prior_gl.rename(
        columns={
            "Account Code": "Account code",
            "account code": "Account code",
            "branch": "Branch",
            "net": "Net",
        },
        inplace=True,
    )

    validate_required_columns(prior_gl, ["Account code", "Debit", "Credit", "Branch"], "Prior Period GL")

    prior_gl["Account code"] = prior_gl["Account code"].astype(str).str.strip()
    prior_gl["Debit"] = pd.to_numeric(prior_gl["Debit"], errors="coerce").fillna(0)
    prior_gl["Credit"] = pd.to_numeric(prior_gl["Credit"], errors="coerce").fillna(0)

    if "Net" not in prior_gl.columns:
        prior_gl["Net"] = prior_gl["Debit"] - prior_gl["Credit"]
    else:
        prior_gl["Net"] = pd.to_numeric(prior_gl["Net"], errors="coerce")
        prior_gl["Net"] = prior_gl["Net"].fillna(prior_gl["Debit"] - prior_gl["Credit"])

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
    "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs",
    "branch_summary", "detected_branches", "validation_passed", "company_profile",
    "bs_disclaimer", "ai_commentary", "prior_pnl", "prior_bs", "prior_kpis",
    "save_run_preference", "anomaly_flags"
]:
    if key not in st.session_state:
        st.session_state[key] = None

if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}

if st.session_state["save_run_preference"] is None:
    st.session_state["save_run_preference"] = False


# ----------------------------
# Header
# ----------------------------
st.title("AI CFO Copilot")
st.caption("Automated branch-wise P&L, consolidated balance sheet, KPI packs, management reporting, memory, and AI commentary")

tab_profile, tab_upload, tab_history, tab_validation, tab_reports, tab_kpis, tab_ai, tab_anomalies, tab_issues, tab_download = st.tabs(
    ["Profile", "Upload", "History & Prior Period", "Validation", "Reports", "KPIs", "AI Insights", "Anomalies", "Issues", "Download"]
)


# ----------------------------
# Profile Tab
# ----------------------------
with tab_profile:
    st.subheader("Company Profile")
    st.caption("Company Name is mandatory and used for memory / restore.")

    c1, c2 = st.columns(2)

    with c1:
        company_name = st.text_input("Company Name *")
        industry = st.selectbox(
            "Industry",
            [
                "Select Industry",
                "Manufacturing",
                "Wholesale / Distribution",
                "Retail",
                "Professional Services",
                "Construction",
                "Logistics",
                "Hospitality",
                "Healthcare",
                "Technology",
                "Other",
            ],
        )
        country = st.selectbox(
            "Country",
            [
                "Select Country",
                "Australia",
                "India",
                "United States",
                "United Kingdom",
                "Canada",
                "New Zealand",
                "Other",
            ],
        )
        state_region = st.text_input("State / Region")
        financial_year = st.text_input("Financial Year", placeholder="Example: FY2025 or 2024-25")

    with c2:
        currency = st.selectbox(
            "Currency",
            ["Select Currency", "AUD", "INR", "USD", "GBP", "CAD", "NZD", "Other"],
        )
        tax_identifier = st.text_input("Tax Identifier / ABN / GSTIN (Optional)")
        reporting_period = st.selectbox(
            "Reporting Period",
            ["Monthly", "Quarterly", "Annual"],
        )
        benchmark_group = st.text_input("Benchmark Group (Optional)")

    business_notes = st.text_area(
        "Business Notes (Optional)",
        placeholder="Example: Multi-branch wholesale distributor with central procurement and branch-level sales reporting."
    )

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
        st.markdown("### Saved Profile")
        profile_df = pd.DataFrame(
            st.session_state["company_profile"].items(),
            columns=["Field", "Value"]
        )
        st.dataframe(profile_df, use_container_width=True)
        st.write(f"**Save future runs:** {'Yes' if st.session_state['save_run_preference'] else 'No'}")


# ----------------------------
# Upload Tab
# ----------------------------
with tab_upload:
    st.subheader("Upload Current Period Source Files")

    c1, c2 = st.columns(2)
    with c1:
        gl_file = st.file_uploader("Current GL Report", type=["xlsx"])
        mapping_file = st.file_uploader("COA Mapping", type=["xlsx"])
    with c2:
        kpi_file = st.file_uploader("KPI Master (Optional)", type=["xlsx"])
        latest_bs_file = st.file_uploader("Latest Previous Balance Sheet (Optional)", type=["xlsx"])

    st.info(
        "GL required columns: Account code, Debit, Credit, Branch. Optional: Net, Date, Description.\n\n"
        "COA mapping required columns: Account code, Reporting Group, Reporting Subgroup, Statement. Recommended: Sign Convention.\n\n"
        "KPI master is optional.\n\n"
        "Latest Previous Balance Sheet is optional. Required columns if uploaded: Reporting Group, Reporting Subgroup, Balance."
    )

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
                    bs_disclaimer = (
                        "Balance Sheet may not fully match because latest previous balance sheet / opening balances "
                        "were not provided. Current output is based only on mapped balance sheet movements available "
                        "in the uploaded GL."
                    )

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
                        summary_row = {"Branch": branch}
                        for _, r in branch_kpis.iterrows():
                            summary_row[r["KPI"]] = r["Display Value"]
                        branch_summary_rows.append(summary_row)

                branch_summary = pd.DataFrame(branch_summary_rows) if branch_summary_rows else pd.DataFrame()

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

                if st.session_state["save_run_preference"]:
                    save_run_to_history(
                        st.session_state["company_profile"],
                        consolidated_pnl,
                        consolidated_bs,
                        consolidated_kpis,
                        branch_summary,
                    )

                prior_kpis = st.session_state.get("prior_kpis", None)
                anomaly_flags = detect_anomalies(
                    consolidated_kpis,
                    branch_outputs,
                    prior_kpis=prior_kpis
                ) if consolidated_kpis is not None else []
                st.session_state["anomaly_flags"] = anomaly_flags

                if unmapped.empty:
                    st.success("Files validated and loaded successfully. No unmapped accounts found.")
                else:
                    st.warning("Files loaded, but unmapped GL rows were found. Fix them before generating reports.")

        except Exception as e:
            st.error(f"Error: {e}")


# ----------------------------
# History & Prior Period Tab
# ----------------------------
with tab_history:
    st.subheader("History / Prior Period Inputs")

    company_name_for_history = st.session_state["company_profile"].get("Company Name", "").strip()

    if not company_name_for_history:
        st.warning("Please save Company Profile first. Company Name is mandatory for history and restore.")
    else:
        st.markdown("### Restore Saved Run")
        saved_runs = list_saved_company_runs(company_name_for_history)

        if saved_runs:
            selected_run = st.selectbox("Select Saved Run", saved_runs)
            if st.button("Restore Selected Run", use_container_width=True):
                restored = restore_run_from_history(company_name_for_history, selected_run)

                st.session_state["prior_pnl"] = restored.get("prior_pnl", None)
                st.session_state["prior_bs"] = restored.get("prior_bs", None)
                st.session_state["prior_kpis"] = restored.get("prior_kpis", None)

                st.success(f"Restored saved run: {selected_run}")
        else:
            st.info("No saved history found for this company yet.")

        st.markdown("### Upload Prior Period Data (Optional)")
        c1, c2 = st.columns(2)

        with c1:
            prior_gl_file = st.file_uploader("Prior Period GL Report (Optional)", type=["xlsx"])
            prior_pnl_file = st.file_uploader("Prior Period P&L (Optional)", type=["xlsx"])

        with c2:
            prior_bs_file = st.file_uploader("Prior Period Balance Sheet (Optional)", type=["xlsx"])
            prior_kpi_file = st.file_uploader("Prior Period KPI Pack (Optional)", type=["xlsx"])

        if st.button("Load Prior Period Inputs", use_container_width=True):
            try:
                coa = st.session_state.get("coa", None)
                kpi_master = st.session_state.get("kpi_master", None)

                loaded_any = False

                if prior_gl_file is not None:
                    if coa is None:
                        st.error("Load current files first so COA mapping is available before using Prior GL.")
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

        if st.session_state.get("prior_pnl") is not None:
            with st.expander("Preview Prior P&L"):
                st.dataframe(st.session_state["prior_pnl"], use_container_width=True)

        if st.session_state.get("prior_bs") is not None:
            with st.expander("Preview Prior Balance Sheet"):
                st.dataframe(st.session_state["prior_bs"], use_container_width=True)

        if st.session_state.get("prior_kpis") is not None:
            with st.expander("Preview Prior KPIs"):
                st.dataframe(st.session_state["prior_kpis"], use_container_width=True)


# ----------------------------
# Validation Tab
# ----------------------------
with tab_validation:
    st.subheader("Validation Summary")

    if st.session_state["gl"] is None:
        st.warning("No validated files loaded yet. Go to Upload tab first.")
    else:
        gl = st.session_state["gl"]
        mapped = st.session_state["mapped"]
        unmapped = st.session_state["unmapped"]
        detected_branches = st.session_state["detected_branches"]

        if st.session_state["company_profile"]:
            st.markdown("### Company Context")
            profile_df = pd.DataFrame(
                st.session_state["company_profile"].items(),
                columns=["Field", "Value"]
            )
            st.dataframe(profile_df, use_container_width=True)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("GL Rows", len(gl))
        m2.metric("Mapped Rows", len(mapped))
        m3.metric("Unmapped Rows", len(unmapped))
        m4.metric("Branches Found", len(detected_branches))

        if st.session_state["latest_bs"] is not None:
            st.success("Latest previous balance sheet uploaded. Consolidated BS will include carried-forward balances.")
        else:
            st.info("Latest previous balance sheet not uploaded. Consolidated BS will be built from mapped BS movements only.")

        st.write("**Detected Branches:**")
        st.write(", ".join(detected_branches) if detected_branches else "No branches detected")

        with st.expander("Preview GL Columns"):
            st.write(list(gl.columns))

        with st.expander("Preview Mapped Data"):
            st.dataframe(mapped.head(20), use_container_width=True)

        st.subheader("Unmapped GL Preview")
        if unmapped.empty:
            st.success("All GL rows mapped correctly.")
        else:
            st.error(f"{len(unmapped)} unmapped rows found. Resolve these before generating reports.")
            cols_to_show = [c for c in ["Account code", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
            st.dataframe(unmapped[cols_to_show], use_container_width=True)

            csv_data = unmapped.to_csv(index=False).encode("utf-8")
            st.download_button(
                "Download Unmapped GL",
                data=csv_data,
                file_name="unmapped_gl.csv",
                mime="text/csv",
                use_container_width=True,
            )


# ----------------------------
# Reports Tab
# ----------------------------
with tab_reports:
    st.subheader("Report Generation")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Unmapped GL rows exist. Resolve them before generating reports.")
    else:
        b1, b2, b3 = st.columns(3)

        if b1.button("Show Consolidated P&L", use_container_width=True):
            st.markdown("### Consolidated P&L")
            st.dataframe(st.session_state["consolidated_pnl"], use_container_width=True)

        if b2.button("Show Consolidated Balance Sheet", use_container_width=True):
            st.markdown("### Consolidated Balance Sheet")
            if st.session_state["consolidated_bs"] is not None and not st.session_state["consolidated_bs"].empty:
                if st.session_state["bs_disclaimer"]:
                    st.warning(st.session_state["bs_disclaimer"])
                st.dataframe(st.session_state["consolidated_bs"], use_container_width=True)
            else:
                st.info("No consolidated balance sheet available.")

        if b3.button("Show Branch P&L", use_container_width=True):
            st.markdown("### Branch-wise P&L")
            for branch in st.session_state["detected_branches"]:
                with st.expander(f"{branch} P&L"):
                    st.dataframe(st.session_state["branch_outputs"][branch]["pnl"], use_container_width=True)


# ----------------------------
# KPIs Tab
# ----------------------------
with tab_kpis:
    st.subheader("KPI Generation")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Unmapped GL rows exist. Resolve them before generating KPI outputs.")
    elif st.session_state["kpi_master"] is None:
        st.info("No KPI master uploaded. KPI generation is skipped.")
    else:
        b1, b2 = st.columns(2)

        if b1.button("Show Consolidated KPIs", use_container_width=True):
            st.markdown("### Consolidated KPIs")
            st.dataframe(
                st.session_state["consolidated_kpis"][["KPI", "Display Value"]],
                use_container_width=True,
            )

        if b2.button("Show Branch KPI Pack", use_container_width=True):
            st.markdown("### Branch Summary KPIs")
            st.dataframe(st.session_state["branch_summary"], use_container_width=True)

            for branch in st.session_state["detected_branches"]:
                with st.expander(f"{branch} KPIs"):
                    st.dataframe(
                        st.session_state["branch_outputs"][branch]["kpis"][["KPI", "Display Value"]],
                        use_container_width=True,
                    )


# ----------------------------
# AI Insights Tab
# ----------------------------
with tab_ai:
    st.subheader("AI Financial Insights")

    if st.session_state["mapped"] is None:
        st.warning("Please upload and validate data first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped accounts before generating AI insights.")
    else:
        st.info("Generate CFO-style commentary based on the current outputs, company profile, and anomaly flags.")

        if st.button("Generate AI Insights", use_container_width=True):
            with st.spinner("Analyzing financials..."):
                commentary = generate_ai_commentary(
                    st.session_state["consolidated_pnl"],
                    st.session_state["consolidated_kpis"],
                    st.session_state["consolidated_bs"],
                    st.session_state["company_profile"],
                    anomaly_flags=st.session_state.get("anomaly_flags", []),
                )
                st.session_state["ai_commentary"] = commentary

        if st.session_state["ai_commentary"]:
            st.markdown("### AI Commentary")
            st.write(st.session_state["ai_commentary"])


# ----------------------------
# Anomalies Tab
# ----------------------------
with tab_anomalies:
    st.subheader("Anomaly Detection")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Resolve unmapped GL rows before anomaly detection.")
    else:
        flags = st.session_state.get("anomaly_flags", [])
        if flags:
            for flag in flags:
                st.warning(flag)
        else:
            st.success("No major anomalies detected based on current rules.")


# ----------------------------
# Issues Tab
# ----------------------------
with tab_issues:
    st.subheader("Issues & Exceptions")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    else:
        unmapped = st.session_state["unmapped"]

        if unmapped.empty:
            st.success("No unmapped accounts found.")
        else:
            st.error(f"{len(unmapped)} unmapped rows found. Review before relying on outputs.")
            cols_to_show = [c for c in ["Account code", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
            st.dataframe(unmapped[cols_to_show], use_container_width=True)


# ----------------------------
# Download Tab
# ----------------------------
with tab_download:
    st.subheader("Download Outputs")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    elif not st.session_state["validation_passed"]:
        st.error("Unmapped GL rows exist. Resolve them before downloading reports.")
    else:
        st.markdown("### Core Reports")

        col1, col2 = st.columns(2)

        with col1:
            pnl_bytes = dataframe_to_excel_bytes({
                "Consolidated P&L": st.session_state["consolidated_pnl"]
            })
            st.download_button(
                label="Download Consolidated P&L",
                data=pnl_bytes,
                file_name="consolidated_pnl.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            if st.session_state["consolidated_kpis"] is not None:
                kpi_bytes = dataframe_to_excel_bytes({
                    "Consolidated KPIs": st.session_state["consolidated_kpis"]
                })
                st.download_button(
                    label="Download Consolidated KPIs",
                    data=kpi_bytes,
                    file_name="consolidated_kpis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with col2:
            if (
                st.session_state["consolidated_bs"] is not None
                and not st.session_state["consolidated_bs"].empty
            ):
                bs_bytes = dataframe_to_excel_bytes({
                    "Consolidated Balance Sheet": st.session_state["consolidated_bs"]
                })
                st.download_button(
                    label="Download Consolidated Balance Sheet",
                    data=bs_bytes,
                    file_name="consolidated_balance_sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                if st.session_state.get("bs_disclaimer"):
                    st.warning(st.session_state["bs_disclaimer"])
            else:
                st.info("No balance sheet available for download.")

            if (
                st.session_state["branch_summary"] is not None
                and not st.session_state["branch_summary"].empty
            ):
                summary_bytes = dataframe_to_excel_bytes({
                    "Branch Summary KPIs": st.session_state["branch_summary"]
                })
                st.download_button(
                    label="Download Branch Summary KPIs",
                    data=summary_bytes,
                    file_name="branch_summary_kpis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        st.markdown("### Data Issues")

        if not st.session_state["unmapped"].empty:
            unmapped_csv = st.session_state["unmapped"].to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download Unmapped GL",
                data=unmapped_csv,
                file_name="unmapped_gl.csv",
                mime="text/csv",
                use_container_width=True,
            )
        else:
            st.success("No unmapped GL entries")

        st.markdown("### Branch-wise Downloads")

        for branch in st.session_state["detected_branches"]:
            with st.expander(f"{branch} Reports"):
                branch_pnl_bytes = dataframe_to_excel_bytes({
                    f"{branch} P&L": st.session_state["branch_outputs"][branch]["pnl"]
                })
                st.download_button(
                    label=f"Download {branch} P&L",
                    data=branch_pnl_bytes,
                    file_name=f"{branch.lower().replace(' ', '_')}_pnl.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"pnl_{branch}",
                )

                if st.session_state["branch_outputs"][branch]["kpis"] is not None:
                    branch_kpi_bytes = dataframe_to_excel_bytes({
                        f"{branch} KPIs": st.session_state["branch_outputs"][branch]["kpis"]
                    })
                    st.download_button(
                        label=f"Download {branch} KPIs",
                        data=branch_kpi_bytes,
                        file_name=f"{branch.lower().replace(' ', '_')}_kpis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"kpi_{branch}",
                    )

        st.markdown("### Full Management Pack")

        full_pack_bytes = create_excel_pack(
            consolidated_pnl=st.session_state["consolidated_pnl"],
            consolidated_bs=st.session_state["consolidated_bs"],
            consolidated_kpis=st.session_state["consolidated_kpis"],
            branch_summary=st.session_state["branch_summary"],
            branch_outputs=st.session_state["branch_outputs"],
            unmapped=st.session_state["unmapped"],
        )

        st.download_button(
            label="Download Full Management Pack",
            data=full_pack_bytes,
            file_name="full_management_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
