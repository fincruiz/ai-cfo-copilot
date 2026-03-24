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


def standardize_key_columns(gl: pd.DataFrame, coa: pd.DataFrame, kpi: pd.DataFrame | None = None, latest_bs: pd.DataFrame | None = None):
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

    return gl, coa, kpi, latest_bs


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


def build_balance_sheet(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    if "Balance" in bs_df.columns and "Reporting Group" in bs_df.columns and "Reporting Subgroup" in bs_df.columns:
        out = (
            bs_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Balance"]
            .sum()
            .reset_index()
            .sort_values(["Reporting Group", "Reporting Subgroup"])
        )
        return out

    out = (
        bs_df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Balance"})
        .sort_values(["Reporting Group", "Reporting Subgroup"])
    )
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

    mapped["Report Value"] = mapped.apply(apply_sign_convention, axis=1)

    pnl_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "income statement"].copy()
    bs_mapped = mapped[mapped["Statement"].astype(str).str.strip().str.lower() == "balance sheet"].copy()

    if latest_bs is not None:
        latest_bs["Balance"] = pd.to_numeric(latest_bs["Balance"], errors="coerce").fillna(0)

    return gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped


# ----------------------------
# Session defaults
# ----------------------------
for key in [
    "gl", "coa", "kpi_master", "latest_bs", "mapped", "pnl_mapped", "bs_mapped", "unmapped",
    "consolidated_pnl", "consolidated_bs", "consolidated_kpis", "branch_outputs",
    "branch_summary", "detected_branches", "validation_passed", "company_profile"
]:
    if key not in st.session_state:
        st.session_state[key] = None

if st.session_state["company_profile"] is None:
    st.session_state["company_profile"] = {}


# ----------------------------
# Header
# ----------------------------
st.title("AI CFO Copilot")
st.caption("Automated branch-wise P&L, consolidated balance sheet, KPI packs, and management reporting from GL data")

tab_profile, tab_upload, tab_validation, tab_reports, tab_kpis, tab_issues, tab_download = st.tabs(
    ["Profile", "Upload", "Validation", "Reports", "KPIs", "Issues", "Download"]
)


# ----------------------------
# Profile Tab
# ----------------------------
with tab_profile:
    st.subheader("Company Profile")
    st.caption("Optional business context to support better insights, benchmarking, and commentary.")

    c1, c2 = st.columns(2)

    with c1:
        company_name = st.text_input("Company Name")
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

    if st.button("Save Company Profile", use_container_width=True):
        if industry == "Select Industry" or country == "Select Country":
            st.error("Please select at least Industry and Country.")
        else:
            st.session_state["company_profile"] = {
                "Company Name": company_name,
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
            st.success("Company profile saved successfully.")

    if st.session_state["company_profile"]:
        st.markdown("### Saved Profile")
        profile_df = pd.DataFrame(
            st.session_state["company_profile"].items(),
            columns=["Field", "Value"]
        )
        st.dataframe(profile_df, use_container_width=True)


# ----------------------------
# Upload Tab
# ----------------------------
with tab_upload:
    st.subheader("Upload Source Files")

    c1, c2 = st.columns(2)
    with c1:
        gl_file = st.file_uploader("GL Report", type=["xlsx"])
        mapping_file = st.file_uploader("COA Mapping", type=["xlsx"])
    with c2:
        kpi_file = st.file_uploader("KPI Master (Optional)", type=["xlsx"])
        latest_bs_file = st.file_uploader("Latest Balance Sheet (Optional)", type=["xlsx"])

    st.info(
        "GL required columns: Account code, Debit, Credit, Branch. Optional: Net, Date, Description.\n\n"
        "COA mapping required columns: Account code, Reporting Group, Reporting Subgroup, Statement. Recommended: Sign Convention.\n\n"
        "KPI master is optional.\n\n"
        "Latest Balance Sheet is optional. If uploaded, required columns are: Reporting Group, Reporting Subgroup, Balance."
    )

    if st.button("Validate & Load Files", use_container_width=True):
        try:
            if not (gl_file and mapping_file):
                st.error("Please upload GL Report and COA Mapping.")
            else:
                gl, coa, kpi_master, latest_bs, mapped, pnl_mapped, bs_mapped, unmapped = prepare_data(
                    gl_file, mapping_file, kpi_file, latest_bs_file
                )

                consolidated_pnl = build_pnl(pnl_mapped)

                if latest_bs is not None:
                    consolidated_bs = build_balance_sheet(latest_bs)
                elif not bs_mapped.empty:
                    consolidated_bs = build_balance_sheet(bs_mapped)
                else:
                    consolidated_bs = pd.DataFrame()

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

                if unmapped.empty:
                    st.success("Files validated and loaded successfully. No unmapped accounts found.")
                else:
                    st.warning("Files loaded, but unmapped GL rows were found. Fix them before generating reports.")

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

        if st.session_state["consolidated_bs"] is not None and not st.session_state["consolidated_bs"].empty:
            st.success("Consolidated Balance Sheet source is available.")
        else:
            st.info("No Balance Sheet source detected. Upload latest BS or include BS accounts in GL + mapping.")

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
        st.markdown("### Individual Downloads")

        col1, col2 = st.columns(2)

        with col1:
            pnl_bytes = dataframe_to_excel_bytes({"Consolidated P&L": st.session_state["consolidated_pnl"]})
            st.download_button(
                label="Download Consolidated P&L",
                data=pnl_bytes,
                file_name="consolidated_pnl.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            if st.session_state["consolidated_bs"] is not None and not st.session_state["consolidated_bs"].empty:
                bs_bytes = dataframe_to_excel_bytes({"Consolidated BS": st.session_state["consolidated_bs"]})
                st.download_button(
                    label="Download Consolidated Balance Sheet",
                    data=bs_bytes,
                    file_name="consolidated_balance_sheet.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if st.session_state["consolidated_kpis"] is not None:
                kpi_bytes = dataframe_to_excel_bytes({"Consolidated KPIs": st.session_state["consolidated_kpis"]})
                st.download_button(
                    label="Download Consolidated KPIs",
                    data=kpi_bytes,
                    file_name="consolidated_kpis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        with col2:
            if st.session_state["branch_summary"] is not None and not st.session_state["branch_summary"].empty:
                summary_bytes = dataframe_to_excel_bytes({"Branch Summary KPIs": st.session_state["branch_summary"]})
                st.download_button(
                    label="Download Branch Summary KPIs",
                    data=summary_bytes,
                    file_name="branch_summary_kpis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            if not st.session_state["unmapped"].empty:
                unmapped_csv = st.session_state["unmapped"].to_csv(index=False).encode("utf-8")
                st.download_button(
                    label="Download Unmapped GL",
                    data=unmapped_csv,
                    file_name="unmapped_gl.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

        st.markdown("### Branch-wise Individual Downloads")
        for branch in st.session_state["detected_branches"]:
            with st.expander(f"{branch} Downloads"):
                branch_pnl_bytes = dataframe_to_excel_bytes({f"{branch} P&L": st.session_state["branch_outputs"][branch]["pnl"]})
                st.download_button(
                    label=f"Download {branch} P&L",
                    data=branch_pnl_bytes,
                    file_name=f"{branch.lower().replace(' ', '_')}_pnl.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key=f"dl_pnl_{branch}",
                )

                if st.session_state["branch_outputs"][branch]["kpis"] is not None:
                    branch_kpi_bytes = dataframe_to_excel_bytes({f"{branch} KPIs": st.session_state["branch_outputs"][branch]["kpis"]})
                    st.download_button(
                        label=f"Download {branch} KPIs",
                        data=branch_kpi_bytes,
                        file_name=f"{branch.lower().replace(' ', '_')}_kpis.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                        key=f"dl_kpi_{branch}",
                    )

        st.markdown("### Full Pack")
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
