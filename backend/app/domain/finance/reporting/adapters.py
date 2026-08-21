from decimal import Decimal

from app.domain.finance.reporting.models import AccountBalance
from app.domain.finance.reporting.rules import canonical_reporting_group


_CREDIT_NORMAL_GROUPS = {
    "Revenue",
    "Other Income",
    "Current Liabilities",
    "Non Current Liabilities",
    "Equity",
}


def _normal_sign(row) -> str:
    sign = str(row.sign_convention or "positive").strip().lower()
    if sign in {"credit", "negative", "invert", "reverse"}:
        return "credit"
    if sign == "debit":
        return "debit"

    # Legacy mappings can contain the generic value ``positive``.  Do not use
    # abs(debit-credit): that masks reversals and contra balances.  Infer the
    # normal balance from the financial classification instead.
    subgroup = str(row.reporting_subgroup or "").strip().lower()
    if "accumulated depreciation" in subgroup:
        return "credit"
    group = canonical_reporting_group(row.reporting_group)
    return "credit" if group in _CREDIT_NORMAL_GROUPS else "debit"


def rows_to_account_balances(rows, *, include_unmapped=False):
    result = []
    for row in rows:
        if not row.reporting_group and not include_unmapped:
            continue
        debit = Decimal(row.debit or 0)
        credit = Decimal(row.credit or 0)
        signed = credit - debit if _normal_sign(row) == "credit" else debit - credit
        result.append(
            AccountBalance(
                str(row.source_account_code),
                row.account_name,
                row.reporting_group or "Unmapped",
                row.reporting_subgroup,
                debit,
                credit,
                signed,
            )
        )
    return result
