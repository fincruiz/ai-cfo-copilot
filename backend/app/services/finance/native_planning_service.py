from __future__ import annotations

from decimal import Decimal

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService


class NativePlanningService:
    def __init__(self, session: AsyncSession):
        self.session = session
        self.reporting = ReportingService(GLTransactionRepository(session))

    async def _ensure_schema(self) -> None:
        """Keep upgraded local databases from failing with a raw 500.

        The P4 SQL migration remains the recommended deployment path. This guard is
        intentionally limited to the two tables owned by Native Planning.
        """
        try:
            await self.session.execute(
                text(
                    """
                    CREATE TABLE IF NOT EXISTS public.planning_versions (
                        id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
                        company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
                        plan_type text NOT NULL CHECK (plan_type IN ('budget','forecast')),
                        version_name text NOT NULL,
                        financial_year_start date NOT NULL,
                        financial_year_end date NOT NULL,
                        status text NOT NULL DEFAULT 'draft' CHECK (status IN ('draft','submitted','approved','locked')),
                        source_type text NOT NULL DEFAULT 'native',
                        assumptions jsonb NOT NULL DEFAULT '{}'::jsonb,
                        created_by uuid NULL,
                        created_at timestamptz NOT NULL DEFAULT now(),
                        updated_at timestamptz NOT NULL DEFAULT now(),
                        CONSTRAINT uq_planning_version UNIQUE(company_id, plan_type, version_name)
                    )
                    """
                )
            )
            await self.session.execute(
                text(
                    """
                    CREATE TABLE IF NOT EXISTS public.native_plan_lines (
                        id uuid PRIMARY KEY DEFAULT gen_random_uuid(),
                        version_id uuid NOT NULL REFERENCES public.planning_versions(id) ON DELETE CASCADE,
                        company_id uuid NOT NULL REFERENCES public.companies(id) ON DELETE CASCADE,
                        period date NOT NULL,
                        branch_id uuid NULL REFERENCES public.branches(id) ON DELETE SET NULL,
                        reporting_group text NOT NULL,
                        reporting_subgroup text NULL,
                        source_account_code text NULL,
                        amount numeric NOT NULL DEFAULT 0,
                        driver_type text NOT NULL DEFAULT 'manual',
                        driver_value numeric NULL,
                        notes text NULL,
                        created_at timestamptz NOT NULL DEFAULT now(),
                        updated_at timestamptz NOT NULL DEFAULT now()
                    )
                    """
                )
            )
            await self.session.execute(
                text(
                    """
                    CREATE UNIQUE INDEX IF NOT EXISTS uq_native_plan_line
                    ON public.native_plan_lines(
                        version_id, period, COALESCE(branch_id::text,''), reporting_group,
                        COALESCE(reporting_subgroup,''), COALESCE(source_account_code,'')
                    )
                    """
                )
            )
            await self.session.commit()
        except Exception as exc:
            await self.session.rollback()
            raise ApplicationError(
                message=(
                    "Planning storage is not ready. Run "
                    "backend/migrations/20260814_p4_customer_beta.sql once in your database, "
                    "then refresh this page."
                ),
                error_code="PLANNING_SCHEMA_NOT_READY",
                status_code=503,
            ) from exc

    async def create_version(self, company_id, request):
        await self._ensure_schema()
        version_id = (
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.planning_versions(
                        company_id,plan_type,version_name,financial_year_start,
                        financial_year_end,assumptions
                    )
                    VALUES (
                        :company_id,:plan_type,:version_name,:start,:end,
                        CAST(:assumptions AS jsonb)
                    )
                    RETURNING id
                    """
                ),
                {
                    "company_id": company_id,
                    "plan_type": request.plan_type,
                    "version_name": request.version_name,
                    "start": request.financial_year_start,
                    "end": request.financial_year_end,
                    "assumptions": '{"seed_growth_percent": '
                    + str(float(request.seed_growth_percent))
                    + "}",
                },
            )
        ).scalar_one()

        if request.seed_from_actuals:
            monthly = await self.reporting.monthly_actuals(company_id)
            growth = Decimal("1") + request.seed_growth_percent / 100
            rows = []
            for row in monthly:
                if request.financial_year_start <= row["month"] <= request.financial_year_end:
                    for key, group in [
                        ("revenue", "Revenue"),
                        ("cost_of_sales", "Cost of Sales"),
                        ("operating_expenses", "Operating Expenses"),
                        ("depreciation", "Depreciation"),
                        ("finance_costs", "Finance Costs"),
                    ]:
                        rows.append(
                            {
                                "version_id": version_id,
                                "company_id": company_id,
                                "period": row["month"],
                                "reporting_group": group,
                                "amount": Decimal(row[key]) * growth,
                            }
                        )
            if rows:
                await self.session.execute(
                    text(
                        """
                        INSERT INTO public.native_plan_lines(
                            version_id,company_id,period,reporting_group,amount,driver_type
                        )
                        VALUES (
                            :version_id,:company_id,:period,:reporting_group,:amount,'actual_growth'
                        )
                        """
                    ),
                    rows,
                )
        await self.session.commit()
        return await self.get_version(company_id, version_id)

    async def list_versions(self, company_id):
        await self._ensure_schema()
        result = await self.session.execute(
            text(
                """
                SELECT id,plan_type,version_name,financial_year_start,financial_year_end,
                       status,assumptions
                FROM public.planning_versions
                WHERE company_id=:company_id
                ORDER BY updated_at DESC
                """
            ),
            {"company_id": company_id},
        )
        return [dict(row) for row in result.mappings().all()]

    async def get_version(self, company_id, version_id):
        await self._ensure_schema()
        version = (
            await self.session.execute(
                text(
                    """
                    SELECT id,plan_type,version_name,financial_year_start,financial_year_end,
                           status,assumptions
                    FROM public.planning_versions
                    WHERE company_id=:company_id AND id=:id
                    """
                ),
                {"company_id": company_id, "id": version_id},
            )
        ).mappings().one()
        lines = [
            dict(row)
            for row in (
                await self.session.execute(
                    text(
                        """
                        SELECT id,period,branch_id,reporting_group,reporting_subgroup,
                               source_account_code,amount,driver_type,driver_value,notes
                        FROM public.native_plan_lines
                        WHERE company_id=:company_id AND version_id=:id
                        ORDER BY period,reporting_group
                        """
                    ),
                    {"company_id": company_id, "id": version_id},
                )
            ).mappings().all()
        ]
        return {**dict(version), "lines": lines}

    async def save_lines(self, company_id, version_id, lines):
        await self._ensure_schema()
        await self.session.execute(
            text(
                "DELETE FROM public.native_plan_lines "
                "WHERE company_id=:company_id AND version_id=:id"
            ),
            {"company_id": company_id, "id": version_id},
        )
        rows = [
            {"version_id": version_id, "company_id": company_id, **line.model_dump()}
            for line in lines
        ]
        if rows:
            await self.session.execute(
                text(
                    """
                    INSERT INTO public.native_plan_lines(
                        version_id,company_id,period,branch_id,reporting_group,
                        reporting_subgroup,source_account_code,amount,driver_type,
                        driver_value,notes
                    )
                    VALUES (
                        :version_id,:company_id,:period,:branch_id,:reporting_group,
                        :reporting_subgroup,:source_account_code,:amount,:driver_type,
                        :driver_value,:notes
                    )
                    """
                ),
                rows,
            )
        await self.session.commit()
        return await self.get_version(company_id, version_id)
