from types import MappingProxyType


REPORTING_GROUP_ORDER = MappingProxyType(
    {
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
)


REPORTING_GROUP_ALIASES = MappingProxyType(
    {
        "sales": "Revenue",
        "revenue": "Revenue",
        "cost of sales": "Cost of Sales",
        "cogs": "Cost of Sales",
        "cost of goods sold": "Cost of Sales",
        "operating expense": "Operating Expenses",
        "operating expenses": "Operating Expenses",
        "overheads": "Operating Expenses",
        "opex": "Operating Expenses",
        "finance costs": "Finance Costs",
        "interest": "Finance Costs",
        "current asset": "Current Assets",
        "current assets": "Current Assets",
        "non current asset": "Non Current Assets",
        "non current assets": "Non Current Assets",
        "non-current assets": "Non Current Assets",
        "current liability": "Current Liabilities",
        "current liabilities": "Current Liabilities",
        "non current liability": "Non Current Liabilities",
        "non current liabilities": "Non Current Liabilities",
        "non-current liabilities": "Non Current Liabilities",
        "equity": "Equity",
    }
)


P_AND_L_GROUPS = frozenset(
    {
        "Revenue",
        "Cost of Sales",
        "Operating Expenses",
        "Depreciation",
        "Other Income",
        "Other Expenses",
        "Finance Costs",
        "Tax",
    }
)


BALANCE_SHEET_GROUPS = frozenset(
    {
        "Current Assets",
        "Non Current Assets",
        "Current Liabilities",
        "Non Current Liabilities",
        "Equity",
    }
)


def canonical_reporting_group(
    value: str | None,
) -> str | None:
    if value is None:
        return None

    cleaned = " ".join(
        str(value).strip().split()
    )

    if not cleaned:
        return None

    return REPORTING_GROUP_ALIASES.get(
        cleaned.lower(),
        cleaned,
    )


def reporting_group_order(
    value: str | None,
) -> int:
    canonical = canonical_reporting_group(value)

    if canonical is None:
        return 999

    return REPORTING_GROUP_ORDER.get(
        canonical,
        999,
    )