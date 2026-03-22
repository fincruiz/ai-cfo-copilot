import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="AI CFO Copilot", layout="wide")


# ----------------------------
# Helpers
# ----------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = df.columns.astype(str).str.strip()
    return df


def standardize_key_columns(gl: pd.DataFrame, coa: pd.DataFrame, kpi: pd.DataFrame):
    gl = clean_columns(gl)
    coa = clean_columns(coa)
    kpi = clean_columns(kpi)

    # Standardize common key columns
    gl.rename(
        columns={
            "Account Code": "Account code",
            "Account code ": "Account code",
            "account code": "Account code",
            "branch": "Branch",
            "net": "Net",
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

    return gl, coa, kpi


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], file_label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"{file_label} is missing required columns: {', '.join(missing)}")


def apply_sign_convention(row) -> float:
    """
    Converts Net into management-reporting sign using COA mapping sign convention.
    Supported sign values:
    - positive
    - negative
    If blank/unknown, leaves abs(Net) for safety in MVP.
    """
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

    # Pretty formatting for display/export consistency
    kpi_df["Display Value"] = kpi_df.apply(
        lambda r: f"{r['Value']:.2f}%" if r["Output Type"] == "percent" else round(r["Value"], 2),
        axis=1,
    )

    return kpi_df[["KPI", "Value", "Output Type", "Display Value"]]


def create_excel_pack(
    consolidated_pnl: pd.DataFrame,
    consolidated_kpis: pd.DataFrame,
    branch_summary: pd.DataFrame,
    branch_outputs: dict,
    unmapped: pd.DataFrame,
) -> bytes:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        consolidated_pnl.to_excel(writer, sheet_name="Consolidated P&L", index=False)
        consolidated_kpis.to_excel(writer, sheet_name="Consolidated KPIs", index=False)
        branch_summary.to_excel(writer, sheet_name="Branch Summary KPIs", index=False)

        for branch, reports in branch_outputs.items():
            safe_branch = str(branch)[:20]
            reports["pnl"].to_excel(writer, sheet_name=f"{safe_branch} P&L", index=False)
            reports["kpis"].to_excel(writer, sheet_name=f"{safe_branch} KPIs", index=False)

        if not unmapped.empty:
            unmapped.to_excel(writer, sheet_name="Unmapped Accounts", index=False)

    return output.getvalue()


# ----------------------------
# UI
# ----------------------------
st.title("AI CFO Copilot")
st.header("Upload GL and COA Mapping")

gl_file = st.file_uploader("Upload GL Report", type=["xlsx"])
mapping_file = st.file_uploader("Upload COA Mapping", type=["xlsx"])
kpi_file = st.file_uploader("Upload KPI Master", type=["xlsx"])


if st.button("Generate Branch Packs"):
    try:
        if not (gl_file and mapping_file and kpi_file):
            st.error("Please upload the GL report, COA mapping, and KPI master.")
            st.stop()

        # Read files
        gl = pd.read_excel(gl_file)
        coa = pd.read_excel(mapping_file)
        kpi_master = pd.read_excel(kpi_file)

        gl, coa, kpi_master = standardize_key_columns(gl, coa, kpi_master)

        # Validate required columns
        validate_required_columns(
            gl,
            ["Account code", "Debit", "Credit", "Branch"],
            "GL report",
        )
        validate_required_columns(
            coa,
            ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"],
            "COA mapping",
        )
        validate_required_columns(
            kpi_master,
            [
                "KPI Name",
                "Formula Type",
                "Numerator Group",
                "Denominator Group",
                "Output Type",
                "Display Order",
            ],
            "KPI master",
        )

        # Ensure Account code type alignment
        gl["Account code"] = gl["Account code"].astype(str).str.strip()
        coa["Account code"] = coa["Account code"].astype(str).str.strip()
        gl["Branch"] = gl["Branch"].astype(str).str.strip()

        # Build Net if missing / incomplete
        if "Net" not in gl.columns:
            gl["Net"] = gl["Debit"].fillna(0) - gl["Credit"].fillna(0)
        else:
            gl["Net"] = gl["Net"].fillna(gl["Debit"].fillna(0) - gl["Credit"].fillna(0))

        # Use only P&L mappings for this version
        coa = coa[coa["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()

        # Merge GL with mapping
        data = gl.merge(coa, on="Account code", how="left")

        # Unmapped rows
        unmapped = data[data["Reporting Group"].isna()].copy()

        # Mapped rows only
        mapped = data[data["Reporting Group"].notna()].copy()

        # Sign handling
        if "Sign Convention" not in mapped.columns:
            mapped["Sign Convention"] = "positive"

        mapped["Report Value"] = mapped.apply(apply_sign_convention, axis=1)

        # Consolidated P&L + KPIs
        consolidated_pnl = build_pnl(mapped)
        consolidated_kpis = build_kpis(mapped, kpi_master)

        # Branch outputs
        branch_list = sorted(mapped["Branch"].dropna().unique().tolist())
        branch_outputs = {}
        branch_summary_rows = []

        for branch in branch_list:
            branch_df = mapped[mapped["Branch"] == branch].copy()
            branch_pnl = build_pnl(branch_df)
            branch_kpis = build_kpis(branch_df, kpi_master)

            branch_outputs[branch] = {"pnl": branch_pnl, "kpis": branch_kpis}

            summary_row = {"Branch": branch}
            for _, r in branch_kpis.iterrows():
                summary_row[r["KPI"]] = r["Display Value"]
            branch_summary_rows.append(summary_row)

        branch_summary = pd.DataFrame(branch_summary_rows)

        # Display results
        st.subheader("Consolidated P&L")
        st.dataframe(consolidated_pnl, use_container_width=True)

        st.subheader("Consolidated KPIs")
        st.dataframe(consolidated_kpis[["KPI", "Display Value"]], use_container_width=True)

        st.subheader("Branch Summary KPIs")
        st.dataframe(branch_summary, use_container_width=True)

        for branch in branch_list:
            with st.expander(f"View {branch} Pack"):
                st.markdown(f"**{branch} P&L**")
                st.dataframe(branch_outputs[branch]["pnl"], use_container_width=True)

                st.markdown(f"**{branch} KPIs**")
                st.dataframe(
                    branch_outputs[branch]["kpis"][["KPI", "Display Value"]],
                    use_container_width=True,
                )

        if not unmapped.empty:
            st.subheader("⚠️ Unmapped Accounts")
            cols_to_show = [c for c in ["Account code", "Description", "Branch", "Debit", "Credit", "Net"] if c in unmapped.columns]
            st.dataframe(unmapped[cols_to_show], use_container_width=True)

        # Download pack
        excel_bytes = create_excel_pack(
            consolidated_pnl=consolidated_pnl,
            consolidated_kpis=consolidated_kpis,
            branch_summary=branch_summary,
            branch_outputs=branch_outputs,
            unmapped=unmapped,
        )

        st.download_button(
            label="Download Branch Management Pack",
            data=excel_bytes,
            file_name="branch_management_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success("Branch packs generated successfully.")

    except Exception as e:
        st.error(f"Error: {e}")
