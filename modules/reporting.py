import pandas as pd
from core.common import safe_float

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

    out = account_level_report_values(report_df)
    out = out.drop(columns=["Sign Convention"], errors="ignore")
    return apply_reporting_order(out)


def build_balance_sheet_detail(bs_df: pd.DataFrame) -> pd.DataFrame:
    """Account-level balance sheet detail."""
    cols = ["Reporting Group", "Reporting Subgroup", "Account code"]
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=cols + ["Balance"])

    out = account_level_report_values(bs_df)
    out = out.drop(columns=["Sign Convention"], errors="ignore")
    out = out.rename(columns={"Report Value": "Balance"})
    return apply_reporting_order(out)


def apply_sign_convention_to_gl(row) -> float:
    """
    Keep transaction-level value as raw Net. Do not use abs() at transaction level.

    The display sign is applied after grouping by Account code, so debit and credit
    movements inside the same GL account are netted first.
    """
    net = row.get("Net", 0)
    if pd.isna(net):
        return 0.0
    return float(net)


def apply_sign_after_account_group(df: pd.DataFrame, value_col: str = "Report Value") -> pd.DataFrame:
    """Apply Sign Convention after account-level netting."""
    if df is None or df.empty:
        return df

    out = df.copy()
    if "Sign Convention" not in out.columns:
        out["Sign Convention"] = "positive"

    def signed_value(row):
        value = safe_float(row.get(value_col, 0))
        sign = str(row.get("Sign Convention", "positive")).strip().lower()
        display_value = abs(value)
        return -display_value if sign == "negative" else display_value

    out[value_col] = out.apply(signed_value, axis=1)
    return out


def account_level_report_values(report_df: pd.DataFrame, extra_cols=None) -> pd.DataFrame:
    """
    Net transactions by Account code first, then apply display sign convention.
    This fixes accounts with both debit and credit movements.
    """
    if report_df is None or report_df.empty:
        return pd.DataFrame()

    extra_cols = extra_cols or []
    df = report_df.copy()

    for col in ["Account Name", "Display Order", "Sign Convention"]:
        if col not in df.columns:
            df[col] = "" if col != "Display Order" else pd.NA

    group_cols = [
        "Reporting Group",
        "Reporting Subgroup",
        "Account code",
        "Account Name",
        "Display Order",
        "Sign Convention",
    ]

    for col in extra_cols:
        if col in df.columns and col not in group_cols:
            group_cols.append(col)

    grouped = df.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    grouped = apply_sign_after_account_group(grouped, "Report Value")
    return grouped


def infer_pnl_section_from_row(row) -> str:
    """Infer whether a P&L row belongs to Revenue, COGS, Overheads, etc.

    This deliberately checks both Reporting Group and Reporting Subgroup because
    many COA files use Reporting Group as the GL/report line name and use
    Reporting Subgroup as the real financial section, for example:
    - Reporting Group = Sales Revenue Labour...
    - Reporting Subgroup = Income
    """
    group = str(row.get("Reporting Group", "") or "").strip()
    subgroup = str(row.get("Reporting Subgroup", "") or "").strip()
    text = f"{group} {subgroup}".lower()

    # Most specific checks first.
    if any(k in text for k in ["cost of goods", "cost of sales", "cogs", "direct cost", "direct costs"]):
        return "COGS"
    if any(k in text for k in ["other income", "sundry income", "non operating income"]):
        return "Other Income"
    if any(k in text for k in ["other expense", "other expenses", "non operating expense"]):
        return "Other Expenses"
    if any(k in text for k in ["interest", "finance cost", "finance costs", "borrowing cost"]):
        return "Interest"
    if "tax" in text:
        return "Tax"
    if any(k in text for k in ["sales", "revenue", "income"]):
        return "Revenue"
    if any(k in text for k in ["operating expense", "operating expenses", "overhead", "overheads", "opex", "expense", "expenses"]):
        return "Overheads"
    if "gross profit" in text:
        return "Calculated"
    if any(k in text for k in ["net profit", "profit after tax", "profit for the period"]):
        return "Calculated"
    if any(k in text for k in ["operating profit", "ebit", "ebitda"]):
        return "Calculated"
    return "Other"


def _sum_section(pnl_df: pd.DataFrame, section: str) -> float:
    if pnl_df is None or pnl_df.empty or "__Section" not in pnl_df.columns:
        return 0.0
    return float(pd.to_numeric(pnl_df.loc[pnl_df["__Section"] == section, "Report Value"], errors="coerce").fillna(0).sum())


def _make_pnl_total_row(label: str, value: float, order: float, line_type: str = "Total") -> dict:
    return {
        "Reporting Group": label,
        "Reporting Subgroup": "",
        "Display Order": order,
        "Report Value": round(float(value), 2),
        "Line Type": line_type,
    }


def add_pnl_subtotals(base_pnl: pd.DataFrame) -> pd.DataFrame:
    """Insert management-report totals into P&L.

    Output order:
    Revenue lines -> Total Revenue -> COGS lines -> Total COGS -> Gross Profit
    -> Overheads lines -> Total Overheads -> other sections -> Net Profit.
    """
    if base_pnl is None or base_pnl.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"])

    pnl = base_pnl.copy()
    if "Display Order" not in pnl.columns:
        pnl["Display Order"] = pd.NA
    pnl["Display Order"] = pd.to_numeric(pnl["Display Order"], errors="coerce")
    pnl["Report Value"] = pd.to_numeric(pnl["Report Value"], errors="coerce").fillna(0).round(2)
    pnl["__Section"] = pnl.apply(infer_pnl_section_from_row, axis=1)
    pnl["Line Type"] = "Detail"

    section_sort = {
        "Revenue": 1,
        "COGS": 2,
        "Overheads": 4,
        "Other Income": 6,
        "Other Expenses": 7,
        "Interest": 8,
        "Tax": 9,
        "Other": 10,
        "Calculated": 99,
    }
    pnl["__Section Order"] = pnl["__Section"].map(section_sort).fillna(99)
    pnl = pnl.sort_values(["__Section Order", "Display Order", "Reporting Group", "Reporting Subgroup"], na_position="last")

    revenue_total = _sum_section(pnl, "Revenue")
    cogs_raw = _sum_section(pnl, "COGS")
    overheads_raw = _sum_section(pnl, "Overheads")
    other_income = _sum_section(pnl, "Other Income")
    other_expenses_raw = _sum_section(pnl, "Other Expenses")
    interest_raw = _sum_section(pnl, "Interest")
    tax_raw = _sum_section(pnl, "Tax")

    # Costs can be uploaded/displayed as either positive or negative depending on Sign Convention.
    # For management P&L totals, we treat cost sections as deductions using absolute values.
    cogs_total = abs(cogs_raw)
    overheads_total = abs(overheads_raw)
    other_expenses_total = abs(other_expenses_raw)
    interest_total = abs(interest_raw)
    tax_total = abs(tax_raw)

    gross_profit = revenue_total - cogs_total
    net_profit = gross_profit - overheads_total + other_income - other_expenses_total - interest_total - tax_total

    output_rows = []

    def append_section(section: str):
        details = pnl[pnl["__Section"] == section].drop(columns=["__Section", "__Section Order"], errors="ignore")
        if not details.empty:
            output_rows.extend(details.to_dict("records"))

    append_section("Revenue")
    if revenue_total != 0:
        output_rows.append(_make_pnl_total_row("Total Revenue", revenue_total, 1.90))

    append_section("COGS")
    if cogs_total != 0:
        output_rows.append(_make_pnl_total_row("Total COGS", cogs_total, 2.90))

    if revenue_total != 0 or cogs_total != 0:
        output_rows.append(_make_pnl_total_row("Gross Profit", gross_profit, 3.00, "Subtotal"))

    append_section("Overheads")
    if overheads_total != 0:
        output_rows.append(_make_pnl_total_row("Total Overheads", overheads_total, 4.90))

    append_section("Other Income")
    if other_income != 0:
        output_rows.append(_make_pnl_total_row("Total Other Income", other_income, 6.90))

    append_section("Other Expenses")
    if other_expenses_total != 0:
        output_rows.append(_make_pnl_total_row("Total Other Expenses", other_expenses_total, 7.90))

    append_section("Interest")
    if interest_total != 0:
        output_rows.append(_make_pnl_total_row("Total Interest / Finance Costs", interest_total, 8.90))

    append_section("Tax")
    if tax_total != 0:
        output_rows.append(_make_pnl_total_row("Total Tax", tax_total, 9.90))

    append_section("Other")

    output_rows.append(_make_pnl_total_row("Net Profit", net_profit, 99.00, "Final Profit"))

    out = pd.DataFrame(output_rows)
    preferred_cols = ["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"]
    for col in preferred_cols:
        if col not in out.columns:
            out[col] = "" if col != "Report Value" else 0.0
    out["Report Value"] = pd.to_numeric(out["Report Value"], errors="coerce").fillna(0).round(2)
    return out[preferred_cols].reset_index(drop=True)


def build_pnl(report_df: pd.DataFrame) -> pd.DataFrame:
    if report_df is None or report_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Display Order", "Report Value", "Line Type"])

    account_values = account_level_report_values(report_df)

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in account_values.columns:
        group_cols.append("Display Order")

    base_pnl = account_values.groupby(group_cols, dropna=False)["Report Value"].sum().reset_index()
    base_pnl = apply_reporting_order(base_pnl)
    return add_pnl_subtotals(base_pnl)


def build_balance_sheet_from_gl(bs_df: pd.DataFrame) -> pd.DataFrame:
    if bs_df is None or bs_df.empty:
        return pd.DataFrame(columns=["Reporting Group", "Reporting Subgroup", "Balance"])

    account_values = account_level_report_values(bs_df)

    group_cols = ["Reporting Group", "Reporting Subgroup"]
    if "Display Order" in account_values.columns:
        group_cols.append("Display Order")

    bs = (
        account_values.groupby(group_cols, dropna=False)["Report Value"]
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
    if report_df is not None and not report_df.empty:
        account_values = account_level_report_values(report_df)
        group_values = account_values.groupby("Reporting Group")["Report Value"].sum().to_dict()
    else:
        group_values = {}
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
    account_values = account_level_report_values(pnl_mapped, extra_cols=["Branch"])
    return (
        account_values.groupby(["Branch", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Actual"})
    )


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
    account_values = account_level_report_values(df, extra_cols=["Month"])
    return (
        account_values.groupby(["Month", "Reporting Group"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Reporting Group"])
    )


def build_monthly_branch_actuals(pnl_mapped: pd.DataFrame) -> pd.DataFrame:
    if pnl_mapped is None or pnl_mapped.empty or "Date" not in pnl_mapped.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    df = pnl_mapped.copy()
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df[df["Date"].notna()].copy()
    if df.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    if "Branch" not in df.columns:
        df["Branch"] = "Consolidated"
    df["Branch"] = df["Branch"].fillna("Consolidated").astype(str).str.strip().replace("", "Consolidated")
    df["Month"] = df["Date"].dt.to_period("M").astype(str)

    # Revenue names vary by client, e.g. "Sales Revenue Labour..." rather than exactly "Revenue".
    group_text = df["Reporting Group"].astype(str).str.strip().str.lower()
    subgroup_text = df["Reporting Subgroup"].astype(str).str.strip().str.lower() if "Reporting Subgroup" in df.columns else ""
    rev = df[
        group_text.str.contains("revenue|sales|income", na=False)
        | (subgroup_text.str.contains("income|revenue|sales", na=False) if hasattr(subgroup_text, "str") else False)
    ].copy()

    if rev.empty:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    account_values = account_level_report_values(rev, extra_cols=["Month", "Branch"])
    if account_values.empty or "Month" not in account_values.columns or "Branch" not in account_values.columns:
        return pd.DataFrame(columns=["Month", "Branch", "Amount"])

    return (
        account_values.groupby(["Month", "Branch"], dropna=False)["Report Value"]
        .sum()
        .reset_index()
        .rename(columns={"Report Value": "Amount"})
        .sort_values(["Month", "Branch"])
    )


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

