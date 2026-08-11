from dataclasses import dataclass


@dataclass(frozen=True)
class MappingSuggestion:
    statement: str
    reporting_group: str
    reporting_subgroup: str | None
    sign_convention: str
    confidence: float
    reason: str


RULES = [
    (("sales", "revenue", "turnover"), "income_statement", "Revenue", "Sales", "credit", 0.95),
    (("cost of sales", "cogs", "cost of goods", "raw material", "purchases", "direct labour", "direct labor"), "income_statement", "Cost of Sales", None, "debit", 0.92),
    (("rent", "salary", "wages", "admin", "marketing", "advertising", "insurance", "utilities", "travel", "software", "professional fee"), "income_statement", "Operating Expenses", None, "debit", 0.88),
    (("depreciation", "amortisation", "amortization"), "income_statement", "Depreciation", None, "debit", 0.95),
    (("interest income",), "income_statement", "Other Income", None, "credit", 0.92),
    (("interest", "finance charge", "bank charge"), "income_statement", "Finance Costs", None, "debit", 0.90),
    (("tax", "income tax"), "income_statement", "Tax", None, "debit", 0.90),
    (("cash", "bank"), "balance_sheet", "Current Assets", "Cash and Cash Equivalents", "debit", 0.95),
    (("receivable", "debtor", "accounts receivable"), "balance_sheet", "Current Assets", "Trade Receivables", "debit", 0.95),
    (("inventory", "stock"), "balance_sheet", "Current Assets", "Inventory", "debit", 0.93),
    (("prepayment",), "balance_sheet", "Current Assets", "Prepayments", "debit", 0.90),
    (("fixed asset", "plant", "equipment", "vehicle", "property"), "balance_sheet", "Non Current Assets", "Property Plant and Equipment", "debit", 0.88),
    (("payable", "creditor", "accounts payable"), "balance_sheet", "Current Liabilities", "Trade Payables", "credit", 0.95),
    (("gst", "vat", "sales tax payable", "payroll payable"), "balance_sheet", "Current Liabilities", None, "credit", 0.88),
    (("loan", "borrow", "debt", "finance lease"), "balance_sheet", "Non Current Liabilities", "Borrowings", "credit", 0.90),
    (("capital", "equity", "retained earning", "shareholder"), "balance_sheet", "Equity", None, "credit", 0.92),
]


def suggest_mapping(account_code: str, account_name: str | None) -> MappingSuggestion:
    haystack = f"{account_code} {account_name or ''}".lower()
    for keywords, statement, group, subgroup, sign, confidence in RULES:
        matched = next((k for k in keywords if k in haystack), None)
        if matched:
            return MappingSuggestion(statement, group, subgroup, sign, confidence, f"Matched keyword '{matched}'")
    return MappingSuggestion("income_statement", "Operating Expenses", "Unclassified", "debit", 0.30, "No deterministic keyword match; review required")
