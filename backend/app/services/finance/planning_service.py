from __future__ import annotations
import csv, io, re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from uuid import UUID
from sqlalchemy import select, text
from sqlalchemy.ext.asyncio import AsyncSession
from app.database.models.core.branch import Branch
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService


def normalise(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(value or "").strip().lower()).strip("_")


def parse_decimal(value) -> Decimal:
    raw = str(value or "").replace(",", "").replace("$", "").replace("₹", "").strip()
    if not raw:
        return Decimal("0")
    try:
        return Decimal(raw)
    except InvalidOperation as exc:
        raise ValueError(f"Invalid amount: {value}") from exc


def parse_period(value) -> date:
    raw = str(value or "").strip()
    for fmt in ("%Y-%m-%d", "%Y-%m", "%d/%m/%Y", "%m/%Y"):
        try:
            parsed = datetime.strptime(raw, fmt)
            return parsed.date().replace(day=1)
        except ValueError:
            continue
    raise ValueError(f"Invalid period: {value}")


class PlanningService:
    def __init__(self, session: AsyncSession) -> None:
        self.session = session
        self.reporting = ReportingService(GLTransactionRepository(session))

    async def branch_lookup(self, company_id: UUID):
        rows = (await self.session.execute(
            select(Branch).where(Branch.company_id == company_id, Branch.is_active.is_(True))
        )).scalars().all()
        result = {}
        for branch in rows:
            result[branch.branch_code.strip().lower()] = branch.id
            result[branch.branch_name.strip().lower()] = branch.id
        return result

    async def import_plan(
        self,
        *,
        company_id: UUID,
        plan_type: str,
        version_name: str,
        content: bytes,
        replace_existing: bool,
    ):
        decoded = content.decode("utf-8-sig")
        reader = csv.DictReader(io.StringIO(decoded))
        if not reader.fieldnames:
            raise ValueError("The planning file needs a header row.")

        headers = {normalise(name): name for name in reader.fieldnames}
        def column(*aliases):
            for alias in aliases:
                match = headers.get(normalise(alias))
                if match:
                    return match
            return None

        cols = {
            "period": column("period", "month", "date"),
            "reporting_group": column("reporting_group", "reporting group", "group"),
            "reporting_subgroup": column("reporting_subgroup", "reporting subgroup", "subgroup"),
            "account_code": column("account_code", "account code", "source_account_code"),
            "amount": column("amount", "budget", "forecast", "value"),
            "branch": column("branch", "branch_code", "branch code", "location"),
            "notes": column("notes", "comment", "comments"),
        }
        missing = [key for key in ("period", "reporting_group", "amount") if not cols[key]]
        if missing:
            raise ValueError("Missing required planning columns: " + ", ".join(missing))

        branches = await self.branch_lookup(company_id)
        rows = []
        issues = []
        total = 0

        for row_number, row in enumerate(reader, start=2):
            total += 1
            try:
                group = str(row.get(cols["reporting_group"], "")).strip()
                if not group:
                    raise ValueError("Reporting group is required.")
                branch_value = str(row.get(cols["branch"], "") or "").strip() if cols["branch"] else ""
                rows.append({
                    "company_id": company_id,
                    "plan_type": plan_type,
                    "version_name": version_name.strip() or "Default",
                    "period": parse_period(row.get(cols["period"])),
                    "source_account_code": str(row.get(cols["account_code"], "") or "").strip() or None if cols["account_code"] else None,
                    "reporting_group": group,
                    "reporting_subgroup": str(row.get(cols["reporting_subgroup"], "") or "").strip() or None if cols["reporting_subgroup"] else None,
                    "branch_id": branches.get(branch_value.lower()) if branch_value else None,
                    "branch_source_value": branch_value or None,
                    "amount": parse_decimal(row.get(cols["amount"])),
                    "notes": str(row.get(cols["notes"], "") or "").strip() or None if cols["notes"] else None,
                })
            except Exception as exc:
                issues.append({"row_number": row_number, "message": str(exc)})

        if replace_existing:
            await self.session.execute(text(
                "DELETE FROM public.finance_plan_lines "
                "WHERE company_id=:company_id AND plan_type=:plan_type AND version_name=:version_name"
            ), {"company_id": company_id, "plan_type": plan_type, "version_name": version_name})

        if rows:
            await self.session.execute(text("""
                INSERT INTO public.finance_plan_lines
                (company_id, plan_type, version_name, period, source_account_code,
                 reporting_group, reporting_subgroup, branch_id, branch_source_value,
                 amount, notes)
                VALUES
                (:company_id, :plan_type, :version_name, :period, :source_account_code,
                 :reporting_group, :reporting_subgroup, :branch_id, :branch_source_value,
                 :amount, :notes)
            """), rows)

        await self.session.commit()
        return {
            "plan_type": plan_type,
            "version_name": version_name,
            "total_rows": total,
            "inserted_rows": len(rows),
            "invalid_rows": len(issues),
            "issues": issues,
        }

    async def variance(self, company_id: UUID, branch_id: UUID | None = None):
        actual_rows = await self.reporting.monthly_actuals(company_id, branch_id=branch_id)
        actual = {}
        for row in actual_rows:
            period = row["month"]
            for key, group in (
                ("revenue", "Revenue"),
                ("cost_of_sales", "Cost of Sales"),
                ("operating_expenses", "Operating Expenses"),
                ("depreciation", "Depreciation"),
                ("finance_costs", "Finance Costs"),
                ("net_profit", "Net Profit"),
            ):
                actual[(period, group)] = Decimal(row[key])

        plan_rows = (await self.session.execute(text("""
            SELECT period, plan_type, reporting_group, SUM(amount) AS amount
            FROM public.finance_plan_lines
            WHERE company_id=CAST(:company_id AS uuid)
              AND (
                CAST(:branch_id AS uuid) IS NULL
                OR branch_id=CAST(:branch_id AS uuid)
              )
            GROUP BY period, plan_type, reporting_group
            ORDER BY period, reporting_group
        """), {"company_id": company_id, "branch_id": branch_id})).mappings().all()

        budget = {}
        forecast = {}
        for row in plan_rows:
            target = budget if row["plan_type"] == "budget" else forecast
            target[(row["period"], row["reporting_group"])] = Decimal(row["amount"] or 0)

        keys = sorted(set(actual) | set(budget) | set(forecast))
        result = []
        for period, group in keys:
            a = actual.get((period, group), Decimal("0"))
            b = budget.get((period, group), Decimal("0"))
            f = forecast.get((period, group), Decimal("0"))
            bv = a - b
            fv = a - f
            result.append({
                "period": period,
                "reporting_group": group,
                "actual": a,
                "budget": b,
                "forecast": f,
                "budget_variance": bv,
                "budget_variance_percent": (bv / abs(b) * Decimal("100")) if b else None,
                "forecast_variance": fv,
                "forecast_variance_percent": (fv / abs(f) * Decimal("100")) if f else None,
            })
        return result
