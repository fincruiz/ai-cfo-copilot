from decimal import Decimal
from app.domain.finance.reporting.models import AccountBalance


def rows_to_account_balances(rows, *, include_unmapped=False):
    result=[]
    for row in rows:
        if not row.reporting_group and not include_unmapped: continue
        debit=Decimal(row.debit or 0); credit=Decimal(row.credit or 0); net=debit-credit
        sign=(row.sign_convention or "positive").lower()
        if sign in {"credit","negative","invert","reverse"}: signed=-net
        elif sign=="debit": signed=net
        else: signed=abs(net)
        result.append(AccountBalance(str(row.source_account_code),row.account_name,row.reporting_group or "Unmapped",row.reporting_subgroup,debit,credit,signed))
    return result
