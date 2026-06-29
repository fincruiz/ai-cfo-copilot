import pandas as pd


REPORTING_GROUP_ORDER = {
    "Revenue": 1,
    "Sales Revenue": 1,
    "Income": 1,
    "Cost of Sales": 2,
    "COGS": 2,
    "Cost of Goods Sold": 2,
    "Gross Profit": 3,
    "Operating Expenses": 4,
    "Overheads": 4,
    "Operating Profit": 5,
    "Other Income": 6,
    "Other Expenses": 7,
    "Interest": 8,
    "Tax": 9,
    "Net Profit": 10,
    "Assets": 20,
    "Current Assets": 21,
    "Non Current Assets": 22,
    "Liabilities": 30,
    "Current Liabilities": 31,
    "Non Current Liabilities": 32,
    "Equity": 40,
}


def apply_reporting_order(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df

    df = df.copy()

    if "Display Order" in df.columns:
        df["__order"] = pd.to_numeric(df["Display Order"], errors="coerce").fillna(999)
    elif "Reporting Group" in df.columns:
        df["__order"] = df["Reporting Group"].map(REPORTING_GROUP_ORDER).fillna(999)
    else:
        df["__order"] = 999

    sort_cols = ["__order"]

    if "Reporting Group" in df.columns:
        sort_cols.append("Reporting Group")
    if "Reporting Subgroup" in df.columns:
        sort_cols.append("Reporting Subgroup")
    if "Account code" in df.columns:
        sort_cols.append("Account code")

    df = df.sort_values(sort_cols).drop(columns=["__order"], errors="ignore")
    return df.reset_index(drop=True)


def build_pnl(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value"])

    group_cols = ["Reporting Group", "Reporting Subgroup"]

    if "Display Order" in report_df.columns:
        group_cols.append("Display Order")

    pnl = (
        report_df.groupby(group_cols, dropna=False)["Report Value"]
        .sum()
        .reset_index()
    )

    return apply_reporting_order(pnl)


def is_revenue_group(group_name: str) -> bool:
    text = str(group_name).lower()
    return "revenue" in text or text in ["income", "sales"]


def is_cogs_group(group_name: str) -> bool:
    text = str(group_name).lower()
    return (
        "cogs" in text
        or "cost of sales" in text
        or "cost of goods sold" in text
        or "direct cost" in text
    )


def is_overhead_group(group_name: str) -> bool:
    text = str(group_name).lower()
    return (
        "operating expense" in text
        or "overhead" in text
        or "opex" in text
    )


def build_pnl_with_subtotals(pnl_df: pd.DataFrame) -> pd.DataFrame:
    if pnl_df is None or pnl_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value"])

    df = apply_reporting_order(pnl_df.copy())

    rows = []

    revenue_total = 0.0
    cogs_total = 0.0
    overhead_total = 0.0
    other_total = 0.0

    revenue_rows = df[df["Reporting Group"].apply(is_revenue_group)]
    cogs_rows = df[df["Reporting Group"].apply(is_cogs_group)]
    overhead_rows = df[df["Reporting Group"].apply(is_overhead_group)]

    other_rows = df[
        ~df.index.isin(revenue_rows.index)
        & ~df.index.isin(cogs_rows.index)
        & ~df.index.isin(overhead_rows.index)
    ]

    def append_section(section_df):
        out = []
        for _, r in section_df.iterrows():
            out.append(r.to_dict())
        return out

    if not revenue_rows.empty:
        rows.extend(append_section(revenue_rows))
        revenue_total = float(revenue_rows["Report Value"].sum())
        rows.append({
            "Reporting Group": "Total Revenue",
            "Reporting Subgroup": "",
            "Display Order": 199,
            "Report Value": revenue_total,
        })

    if not cogs_rows.empty:
        rows.extend(append_section(cogs_rows))
        cogs_total = float(cogs_rows["Report Value"].sum())
        rows.append({
            "Reporting Group": "Total COGS",
            "Reporting Subgroup": "",
            "Display Order": 299,
            "Report Value": cogs_total,
        })

    gross_profit = revenue_total - cogs_total
    rows.append({
        "Reporting Group": "Gross Profit",
        "Reporting Subgroup": "",
        "Display Order": 399,
        "Report Value": gross_profit,
    })

    if not overhead_rows.empty:
        rows.extend(append_section(overhead_rows))
        overhead_total = float(overhead_rows["Report Value"].sum())
        rows.append({
            "Reporting Group": "Total Overheads",
            "Reporting Subgroup": "",
            "Display Order": 499,
            "Report Value": overhead_total,
        })

    if not other_rows.empty:
        rows.extend(append_section(other_rows))
        other_total = float(other_rows["Report Value"].sum())

    net_profit = gross_profit - overhead_total + other_total
    rows.append({
        "Reporting Group": "Net Profit",
        "Reporting Subgroup": "",
        "Display Order": 999,
        "Report Value": net_profit,
    })

    result = pd.DataFrame(rows)

    if "Display Order" not in result.columns:
        result["Display Order"] = range(1, len(result) + 1)

    return result[["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value"]]


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
        return current_bs.copy() if current_bs is not None else pd.DataFrame()

    if current_bs is None or current_bs.empty:
        return opening_bs.copy()

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

    return merged[["Reporting Group", "Reporting Subgroup", "Balance"]]


def build_kpis(report_df: pd.DataFrame, kpi_master: pd.DataFrame) -> pd.DataFrame:
    if report_df is None or report_df.empty or kpi_master is None or kpi_master.empty:
        return pd.DataFrame(columns=["KPI", "Value", "Output Type", "Display Value"])

    group_values = report_df.groupby("Reporting Group")["Report Value"].sum().to_dict()

    results = []
    calculated = {}

    kpi_master = kpi_master.copy()

    if "Display Order" in kpi_master.columns:
        kpi_master["Display Order"] = pd.to_numeric(kpi_master["Display Order"], errors="coerce").fillna(999)
        kpi_master = kpi_master.sort_values("Display Order")

    for _, row in kpi_master.iterrows():
        kpi_name = str(row.get("KPI Name", "")).strip()
        formula_type = str(row.get("Formula Type", "")).strip().lower()
        numerator = str(row.get("Numerator Group", "")).strip()
        denominator = str(row.get("Denominator Group", "")).strip()
        output_type = str(row.get("Output Type", "")).strip().lower()

        value = 0.0

        if formula_type == "direct":
            value = group_values.get(numerator, 0.0)
        elif formula_type == "derived":
            value = calculated.get(numerator, group_values.get(numerator, 0.0)) - calculated.get(
                denominator, group_values.get(denominator, 0.0)
            )
        elif formula_type == "ratio":
            num_val = calculated.get(numerator, group_values.get(numerator, 0.0))
            den_val = calculated.get(denominator, group_values.get(denominator, 0.0))
            value = (num_val / den_val * 100) if den_val != 0 else 0.0

        calculated[kpi_name] = value

        results.append({
            "KPI": kpi_name,
            "Value": round(value, 2),
            "Output Type": output_type,
            "Display Value": f"{value:.2f}%" if output_type == "percent" else round(value, 2),
        })

    return pd.DataFrame(results)
