import pandas as pd


def validate_required_columns(df: pd.DataFrame, required_cols: list[str], file_label: str):
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(
            f"{file_label} → Missing columns: {missing}. Found columns: {list(df.columns)}"
        )


def find_coa_duplicate_rows(coa: pd.DataFrame) -> pd.DataFrame:
    if coa is None or coa.empty or "Account code" not in coa.columns:
        return pd.DataFrame()

    duplicated = coa[coa["Account code"].duplicated(keep=False)].copy()
    return duplicated.sort_values("Account code") if not duplicated.empty else pd.DataFrame()


def build_validation_issue(area: str, issue: str, severity: str = "Warning", recommendation: str = ""):
    return {
        "Severity": severity,
        "Area": area,
        "Issue": issue,
        "Recommendation": recommendation,
    }


def calculate_readiness_score(critical_errors, warnings, recommendations):
    score = 100
    score -= len(critical_errors) * 35
    score -= len(warnings) * 8
    score -= len(recommendations) * 3
    return max(score, 0)


KEYWORD_MAPPING_RULES = {
    "freight": ["Cost of Sales", "COGS", "Operating Expenses", "Overheads"],
    "shipping": ["Cost of Sales", "COGS", "Operating Expenses", "Overheads"],
    "delivery": ["Cost of Sales", "COGS", "Operating Expenses", "Overheads"],
    "sales": ["Revenue"],
    "income": ["Revenue", "Other Income"],
    "rent": ["Operating Expenses", "Overheads"],
    "salary": ["Operating Expenses", "Overheads"],
    "wages": ["Cost of Sales", "COGS", "Operating Expenses", "Overheads"],
    "interest": ["Interest"],
    "tax": ["Tax"],
    "depreciation": ["Operating Expenses", "Overheads"],
    "cash": ["Assets", "Current Assets", "Cash and Cash Equivalent"],
    "bank": ["Assets", "Current Assets", "Cash and Cash Equivalent"],
    "receivable": ["Assets", "Current Assets"],
    "payable": ["Liabilities", "Current Liabilities"],
    "loan": ["Liabilities", "Non Current Liabilities"],
}


def review_coa_mapping(coa: pd.DataFrame) -> pd.DataFrame:
    if coa is None or coa.empty:
        return pd.DataFrame(columns=[
            "Account code", "Account Name", "Current Mapping", "Suggested Mapping", "Issue", "Severity"
        ])

    required = ["Account code", "Reporting Group"]
    for col in required:
        if col not in coa.columns:
            return pd.DataFrame(columns=[
                "Account code", "Account Name", "Current Mapping", "Suggested Mapping", "Issue", "Severity"
            ])

    review_rows = []
    coa = coa.copy()

    if "Account Name" not in coa.columns:
        coa["Account Name"] = ""

    for _, row in coa.iterrows():
        account_code = str(row.get("Account code", "")).strip()
        account_name = str(row.get("Account Name", "")).strip()
        current_group = str(row.get("Reporting Group", "")).strip()

        text_to_check = f"{account_code} {account_name}".lower()

        for keyword, suggested_groups in KEYWORD_MAPPING_RULES.items():
            if keyword in text_to_check:
                if current_group and current_group not in suggested_groups:
                    review_rows.append({
                        "Account code": account_code,
                        "Account Name": account_name,
                        "Current Mapping": current_group,
                        "Suggested Mapping": " / ".join(suggested_groups),
                        "Issue": f"Account contains keyword '{keyword}' but is mapped to '{current_group}'.",
                        "Severity": "Warning",
                    })

    return pd.DataFrame(review_rows)


def build_validation_report(critical_errors=None, warnings=None, recommendations=None):
    critical_errors = critical_errors or []
    warnings = warnings or []
    recommendations = recommendations or []

    rows = []

    for item in critical_errors:
        rows.append({
            "Severity": "Critical",
            "Area": item.get("Area", ""),
            "Issue": item.get("Issue", ""),
            "Recommendation": item.get("Recommendation", ""),
        })

    for item in warnings:
        rows.append({
            "Severity": "Warning",
            "Area": item.get("Area", ""),
            "Issue": item.get("Issue", ""),
            "Recommendation": item.get("Recommendation", ""),
        })

    for item in recommendations:
        rows.append({
            "Severity": "Recommendation",
            "Area": item.get("Area", ""),
            "Issue": item.get("Issue", ""),
            "Recommendation": item.get("Recommendation", ""),
        })

    return pd.DataFrame(rows)
