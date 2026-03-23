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

    gl.rename(
        columns={
            "Account Code": "Account code",
            "account code": "Account code",
            "Account code ": "Account code",
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
    kpi_df["Display Value"] = kpi_df.apply(
        lambda r: f"{r['Value']:.2f}%" if r["Output Type"] == "percent" else round(r["Value"], 2),
        axis=1,
    )

    return kpi_df[["KPI", "Value", "Output Type", "Display Value"]]


def create_excel_pack(consolidated_pnl, consolidated_kpis, branch_summary, branch_outputs, unmapped):
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


def prepare_data(gl_file, mapping_file, kpi_file):
    gl = pd.read_excel(gl_file)
    coa = pd.read_excel(mapping_file)
    kpi_master = pd.read_excel(kpi_file)

    gl, coa, kpi_master = standardize_key_columns(gl, coa, kpi_master)

    validate_required_columns(gl, ["Account code", "Debit", "Credit", "Branch"], "GL report")
    validate_required_columns(coa, ["Account code", "Reporting Group", "Reporting Subgroup", "Statement"], "COA mapping")
    validate_required_columns(
        kpi_master,
        ["KPI Name", "Formula Type", "Numerator Group", "Denominator Group", "Output Type", "Display Order"],
        "KPI master",
    )

    gl["Account code"] = gl["Account code"].astype(str).str.strip()
    coa["Account code"] = coa["Account code"].astype(str).str.strip()
    gl["Branch"] = gl["Branch"].astype(str).str.strip()

    if "Net" not in gl.columns:
        gl["Net"] = gl["Debit"].fillna(0) - gl["Credit"].fillna(0)
    else:
        gl["Net"] = gl["Net"].fillna(gl["Debit"].fillna(0) - gl["Credit"].fillna(0))

    coa = coa[coa["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()

    data = gl.merge(coa, on="Account code", how="left")
    unmapped = data[data["Reporting Group"].isna()].copy()
    mapped = data[data["Reporting Group"].notna()].copy()

    if "Sign Convention" not in mapped.columns:
        mapped["Sign Convention"] = "positive"

    mapped["Report Value"] = mapped.apply(apply_sign_convention, axis=1)

    return gl, coa, kpi_master, mapped, unmapped


# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "mapped", "unmapped",
    "consolidated_pnl", "consolidated_kpis", "branch_outputs",
    "branch_summary", "detected_branches"
]:
    if key not in st.session_state:
        st.session_state[key] = None


# ----------------------------
# Header
# ----------------------------
st.title("AI CFO Copilot")
st.caption("Automated branch-wise P&L, KPI packs, and management reporting from GL data")

tab_upload, tab_validation, tab_reports, tab_kpis, tab_issues, tab_download = st.tabs(
    ["Upload", "Validation", "Reports", "KPIs", "Issues", "Download"]
)

# ----------------------------
# Upload Tab
# ----------------------------
with tab_upload:
    st.subheader("Upload Source Files")

    c1, c2, c3 = st.columns(3)
    with c1:
        gl_file = st.file_uploader("GL Report", type=["xlsx"])
    with c2:
        mapping_file = st.file_uploader("COA Mapping", type=["xlsx"])
    with c3:
        kpi_file = st.file_uploader("KPI Master", type=["xlsx"])

    st.info(
        "GL required columns: Account code, Debit, Credit, Branch. Optional: Net, Date, Description.\n\n"
        "COA mapping required columns: Account code, Reporting Group, Reporting Subgroup, Statement. "
        "Recommended: Sign Convention.\n\n"
        "KPI master required columns: KPI Name, Formula Type, Numerator Group, Denominator Group, Output Type, Display Order."
    )

    if st.button("Validate & Load Files", use_container_width=True):
        try:
            if not (gl_file and mapping_file and kpi_file):
                st.error("Please upload GL Report, COA Mapping, and KPI Master.")
            else:
                gl, coa, kpi_master, mapped, unmapped = prepare_data(gl_file, mapping_file, kpi_file)

                consolidated_pnl = build_pnl(mapped)
                consolidated_kpis = build_kpis(mapped, kpi_master)

                detected_branches = sorted(mapped["Branch"].dropna().unique().tolist())

                branch_outputs = {}
                branch_summary_rows = []

                for branch in detected_branches:
                    branch_df = mapped[mapped["Branch"] == branch].copy()
                    branch_pnl = build_pnl(branch_df)
                    branch_kpis = build_kpis(branch_df, kpi_master)

                    branch_outputs[branch] = {"pnl": branch_pnl, "kpis": branch_kpis}

                    summary_row = {"Branch": branch}
                    for _, r in branch_kpis.iterrows():
                        summary_row[r["KPI"]] = r["Display Value"]
                    branch_summary_rows.append(summary_row)

                branch_summary = pd.DataFrame(branch_summary_rows)

                st.session_state["gl"] = gl
                st.session_state["coa"] = coa
                st.session_state["kpi_master"] = kpi_master
                st.session_state["mapped"] = mapped
                st.session_state["unmapped"] = unmapped
                st.session_state["consolidated_pnl"] = consolidated_pnl
                st.session_state["consolidated_kpis"] = consolidated_kpis
                st.session_state["branch_outputs"] = branch_outputs
                st.session_state["branch_summary"] = branch_summary
                st.session_state["detected_branches"] = detected_branches

                st.success("Files validated and loaded successfully. Move to the other tabs to generate outputs.")

        except Exception as e:
            st.error(f"Error: {e}")


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

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("GL Rows", len(gl))
        m2.metric("Mapped Rows", len(mapped))
        m3.metric("Unmapped Rows", len(unmapped))
        m4.metric("Branches Found", len(detected_branches))

        st.write("**Detected Branches:**")
        st.write(", ".join(detected_branches) if detected_branches else "No branches detected")

        with st.expander("Preview GL Columns"):
            st.write(list(gl.columns))

        with st.expander("Preview Mapped Data"):
            st.dataframe(mapped.head(20), use_container_width=True)


# ----------------------------
# Reports Tab
# ----------------------------
with tab_reports:
    st.subheader("Report Generation")

    if st.session_state["mapped"] is None:
        st.warning("Please validate and load files first.")
    else:
        b1, b2, b3 = st.columns(3)

        if b1.button("Show Consolidated P&L", use_container_width=True):
            st.markdown("### Consolidated P&L")
            st.dataframe(st.session_state["consolidated_pnl"], use_container_width=True)

        if b2.button("Show Branch P&L", use_container_width=True):
            st.markdown("### Branch-wise P&L")
            for branch in st.session_state["detected_branches"]:
                with st.expander(f"{branch} P&L"):
                    st.dataframe(st.session_state["branch_outputs"][branch]["pnl"], use_container_width=True)

        if b3.button("Show Full Management Pack", use_container_width=True):
            st.markdown("### Consolidated P&L")
            st.dataframe(st.session_state["consolidated_pnl"], use_container_width=True)

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
    else:
        excel_bytes = create_excel_pack(
            consolidated_pnl=st.session_state["consolidated_pnl"],
            consolidated_kpis=st.session_state["consolidated_kpis"],
            branch_summary=st.session_state["branch_summary"],
            branch_outputs=st.session_state["branch_outputs"],
            unmapped=st.session_state["unmapped"],
        )

        st.download_button(
            label="Download Full Management Pack",
            data=excel_bytes,
            file_name="branch_management_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        st.info("This download includes consolidated P&L, consolidated KPIs, branch-wise P&L, branch-wise KPIs, summary KPIs, and unmapped accounts.")
