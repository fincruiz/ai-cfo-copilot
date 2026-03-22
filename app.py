import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="AI CFO Copilot", layout="wide")

st.title("AI CFO Copilot")
st.header("Upload GL and COA Mapping")

gl_file = st.file_uploader("Upload GL Report", type=["xlsx"])
mapping_file = st.file_uploader("Upload COA Mapping", type=["xlsx"])
kpi_file = st.file_uploader("Upload KPI Master", type=["xlsx"])


def build_pnl(df):
    pnl = (
        df.groupby(["Reporting Group", "Reporting Subgroup"], dropna=False)["Net"]
        .sum()
        .reset_index()
    )
    return pnl


def build_kpis(df, kpi_master):
    grouped_values = df.groupby("Reporting Group")["Net"].sum().to_dict()

    normalized = {k: abs(v) for k, v in grouped_values.items()}

    results = []
    calculated = {}

    kpi_master = kpi_master.sort_values("Display Order").copy()

    for _, row in kpi_master.iterrows():
        kpi_name = row["KPI Name"]
        formula_type = str(row["Formula Type"]).strip().lower()
        numerator = str(row["Numerator Group"]).strip() if pd.notna(row["Numerator Group"]) else ""
        denominator = str(row["Denominator Group"]).strip() if pd.notna(row["Denominator Group"]) else ""
        output_type = str(row["Output Type"]).strip().lower()

        value = 0

        if formula_type == "direct":
            value = normalized.get(numerator, 0)

        elif formula_type == "derived":
            if kpi_name == "Gross Profit":
                value = normalized.get("Revenue", 0) - normalized.get("Cost of Sales", 0)
            elif kpi_name == "Operating Profit":
                value = calculated.get("Gross Profit", 0) - normalized.get("Operating Expense", 0)

        elif formula_type == "ratio":
            if kpi_name == "Gross Margin %":
                gp = calculated.get("Gross Profit", 0)
                rev = normalized.get("Revenue", 0)
                value = (gp / rev * 100) if rev != 0 else 0
            elif kpi_name == "Operating Margin %":
                op = calculated.get("Operating Profit", 0)
                rev = normalized.get("Revenue", 0)
                value = (op / rev * 100) if rev != 0 else 0
            elif kpi_name == "Opex as % of Revenue":
                opex = normalized.get("Operating Expense", 0)
                rev = normalized.get("Revenue", 0)
                value = (opex / rev * 100) if rev != 0 else 0
            else:
                num = normalized.get(numerator, calculated.get(numerator, 0))
                den = normalized.get(denominator, calculated.get(denominator, 0))
                value = (num / den * 100) if den != 0 else 0

        calculated[kpi_name] = value

        results.append({
            "KPI": kpi_name,
            "Value": value,
            "Output Type": output_type
        })

    return pd.DataFrame(results)


if st.button("Generate Branch Packs"):
    if gl_file and mapping_file and kpi_file:
        gl = pd.read_excel(gl_file)
        kpi_master = pd.read_excel(kpi_file)
        kpi_master.columns = kpi_master.columns.str.strip()
        mapping = pd.read_excel(mapping_file)

        # Basic cleanup
        gl.columns = gl.columns.str.strip()
        mapping.columns = mapping.columns.str.strip()

        # Keep only P&L mappings for this MVP
        mapping = mapping[mapping["Statement"] == "Income Statement"].copy()

        # If Net is missing or blank, calculate it
        if "Net" not in gl.columns:
            gl["Net"] = gl["Debit"].fillna(0) - gl["Credit"].fillna(0)
        else:
            gl["Net"] = gl["Net"].fillna(gl["Debit"].fillna(0) - gl["Credit"].fillna(0))

        # Merge mapping
        data = gl.merge(mapping, on="Account code", how="left")

        # Unmapped rows
        unmapped = data[data["Reporting Group"].isna()].copy()

        # Only mapped rows for reporting
        mapped_data = data[data["Reporting Group"].notna()].copy()

        # Consolidated P&L + KPIs
        consolidated_pnl = build_pnl(mapped_data)
        consolidated_kpis = build_kpis(mapped_data, kpi_master)

        # Branch-wise packs
        branch_list = sorted(mapped_data["Branch"].dropna().astype(str).unique().tolist())

        branch_summary_rows = []
        branch_outputs = {}

        for branch in branch_list:
            branch_df = mapped_data[mapped_data["Branch"].astype(str) == branch].copy()
            branch_pnl = build_pnl(branch_df)
            branch_kpis = build_kpis(branch_df, kpi_master)

            branch_outputs[branch] = {
                "pnl": branch_pnl,
                "kpis": branch_kpis
            }

            # Pull summary KPIs for comparison sheet
            summary = {"Branch": branch}
            for _, row in branch_kpis.iterrows():
                summary[row["KPI"]] = row["Value"]
            branch_summary_rows.append(summary)

        branch_summary = pd.DataFrame(branch_summary_rows)

        # Display in app
        st.subheader("Consolidated P&L")
        st.dataframe(consolidated_pnl, use_container_width=True)

        st.subheader("Consolidated KPIs")
        st.dataframe(consolidated_kpis, use_container_width=True)

        st.subheader("Branch Summary KPIs")
        st.dataframe(branch_summary, use_container_width=True)

        if not unmapped.empty:
            st.subheader("⚠️ Unmapped Accounts")
            st.dataframe(
                unmapped[["Account code", "Description", "Branch", "Net"]].copy(),
                use_container_width=True
            )

        # Create downloadable Excel workbook
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

        st.download_button(
            label="Download Branch Management Pack",
            data=output.getvalue(),
            file_name="branch_management_pack.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.error("Please upload both the GL report and the COA mapping file.")