from __future__ import annotations

import json
from collections import defaultdict
from datetime import date
from decimal import Decimal
from uuid import UUID

from sqlalchemy import text
from sqlalchemy.ext.asyncio import AsyncSession

from app.core.exceptions import ApplicationError
from app.domain.finance.planning_seed_engine import allocate_total, annualize_history, month_range, monthly_weights_from_history
from app.repositories.finance.gl_transaction_repository import GLTransactionRepository
from app.services.finance.reporting_service import ReportingService

GROUPS = [
    'Revenue', 'Cost of Sales', 'Payroll', 'Operating Expenses',
    'Depreciation', 'Finance Costs', 'Tax', 'Other Income', 'Other Expenses'
]


class NativePlanningService:
    def __init__(self, session: AsyncSession):
        self.session = session
        self.reporting = ReportingService(GLTransactionRepository(session))

    async def _ensure_schema(self) -> None:
        try:
            await self.session.execute(text("""
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
            """))
            await self.session.execute(text("""
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
            """))
            await self.session.execute(text("""
                CREATE UNIQUE INDEX IF NOT EXISTS uq_native_plan_line
                ON public.native_plan_lines(
                    version_id, period, COALESCE(branch_id::text,''), reporting_group,
                    COALESCE(reporting_subgroup,''), COALESCE(source_account_code,'')
                )
            """))
            await self.session.commit()
        except Exception as exc:
            await self.session.rollback()
            raise ApplicationError(
                message='Planning storage is not ready. Run the planning database migration once, then refresh this page.',
                error_code='PLANNING_SCHEMA_NOT_READY', status_code=503,
            ) from exc

    async def planning_context(self, company_id: UUID) -> dict:
        await self._ensure_schema()
        monthly = await self.reporting.monthly_actuals(company_id)
        health = await self.reporting.data_health(company_id)
        versions = await self.list_versions(company_id)
        imported = (await self.session.execute(text("""
            SELECT plan_type, version_name, MIN(period) first_period, MAX(period) last_period,
                   COUNT(*)::int line_count
            FROM public.finance_plan_lines
            WHERE company_id=:company_id
            GROUP BY plan_type, version_name
            ORDER BY MAX(period) DESC
        """), {'company_id': company_id})).mappings().all()
        return {
            'actual_months': len(monthly),
            'first_actual_month': monthly[0]['month'] if monthly else None,
            'latest_actual_month': monthly[-1]['month'] if monthly else None,
            'mapped_accounts': int(health.get('mapped_account_count') or 0),
            'native_versions': versions,
            'imported_versions': [dict(r) for r in imported],
            'recommended_seed': 'actuals' if monthly else ('previous_budget' if versions or imported else 'blank'),
        }

    async def _account_history(self, company_id: UUID, branch_id: UUID | None = None) -> list[dict]:
        rows = (await self.session.execute(text("""
            SELECT date_trunc('month', gt.transaction_date)::date AS month,
                   gt.source_account_code,
                   COALESCE(MAX(gt.source_account_name), MAX(fam.source_account_name), gt.source_account_code) account_name,
                   fam.reporting_group,
                   fam.reporting_subgroup,
                   SUM(CASE WHEN fam.sign_convention='credit' THEN gt.credit-gt.debit ELSE gt.debit-gt.credit END) amount
            FROM public.gl_transactions gt
            JOIN public.file_uploads fu ON fu.id=gt.file_upload_id
            JOIN public.finance_account_mappings fam
              ON fam.company_id=gt.company_id AND fam.source_account_code=gt.source_account_code
            WHERE gt.company_id=:company_id AND gt.validation_status='valid'
              AND (:branch_id IS NULL OR gt.branch_id=:branch_id)
              AND gt.is_elimination=false AND fu.is_active=true AND fu.processing_status='validated'
              AND fam.statement='income_statement'
            GROUP BY date_trunc('month', gt.transaction_date)::date,
                     gt.source_account_code, fam.reporting_group, fam.reporting_subgroup
            ORDER BY month, fam.reporting_group, gt.source_account_code
        """), {'company_id': company_id, 'branch_id': branch_id})).mappings().all()
        return [dict(r) for r in rows]

    async def _actual_seed(self, company_id: UUID, start: date, end: date, growth: Decimal, detail_level: str) -> list[dict]:
        target_months = month_range(start, end)
        if not target_months:
            return []
        history = await self._account_history(company_id)
        if not history:
            return []
        source_months = sorted({r['month'] for r in history})[-12:]
        history = [r for r in history if r['month'] in source_months]
        by_group_month: dict[str, dict[date, Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal('0')))
        by_group_account: dict[str, dict[tuple[str, str | None, str | None], Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal('0')))
        for r in history:
            group = r['reporting_group']
            amount = Decimal(r['amount'] or 0)
            by_group_month[group][r['month']] += amount
            by_group_account[group][(r['source_account_code'], r['account_name'], r['reporting_subgroup'])] += amount
        factor = Decimal('1') + Decimal(growth) / 100
        result: list[dict] = []
        for group, monthly in by_group_month.items():
            values = [monthly.get(m, Decimal('0')) for m in source_months]
            annual_target = annualize_history(values) * factor * (Decimal(len(target_months)) / Decimal('12'))
            month_amounts = allocate_total(annual_target, monthly_weights_from_history(values, len(target_months)))
            if detail_level == 'detailed':
                accounts = list(by_group_account[group].items())
                account_weights = [abs(v) for _, v in accounts]
                for target_month, group_amount in zip(target_months, month_amounts):
                    splits = allocate_total(group_amount, account_weights)
                    for ((code, name, subgroup), _), amount in zip(accounts, splits):
                        result.append({
                            'period': target_month, 'branch_id': None, 'reporting_group': group,
                            'reporting_subgroup': subgroup or name, 'source_account_code': code,
                            'amount': amount, 'driver_type': 'actuals_ratio', 'driver_value': growth,
                            'notes': 'Seeded from mapped historical actuals and historical account mix.'
                        })
            else:
                for target_month, amount in zip(target_months, month_amounts):
                    result.append({
                        'period': target_month, 'branch_id': None, 'reporting_group': group,
                        'reporting_subgroup': None, 'source_account_code': None,
                        'amount': amount, 'driver_type': 'actuals_ratio', 'driver_value': growth,
                        'notes': 'High-level seed from historical actuals and monthly seasonality.'
                    })
        return result

    async def _previous_seed(self, company_id: UUID, request) -> list[dict]:
        target_months = month_range(request.financial_year_start, request.financial_year_end)
        source_rows: list[dict] = []
        if request.seed_version_id:
            source_rows = [dict(r) for r in (await self.session.execute(text("""
                SELECT period,branch_id,reporting_group,reporting_subgroup,source_account_code,amount
                FROM public.native_plan_lines
                WHERE company_id=:company_id AND version_id=:version_id ORDER BY period,reporting_group
            """), {'company_id': company_id, 'version_id': request.seed_version_id})).mappings().all()]
        elif request.seed_imported_version:
            source_rows = [dict(r) for r in (await self.session.execute(text("""
                SELECT period,branch_id,reporting_group,reporting_subgroup,source_account_code,amount
                FROM public.finance_plan_lines
                WHERE company_id=:company_id AND plan_type=:plan_type AND version_name=:version_name
                ORDER BY period,reporting_group
            """), {'company_id': company_id, 'plan_type': request.plan_type, 'version_name': request.seed_imported_version})).mappings().all()]
        if not source_rows:
            return []
        source_months = sorted({r['period'] for r in source_rows})
        month_map = {m: target_months[i % len(target_months)] for i, m in enumerate(source_months[:len(target_months)])}
        factor = Decimal('1') + request.seed_growth_percent / 100
        out = []
        for r in source_rows:
            if r['period'] not in month_map:
                continue
            detailed = request.detail_level == 'detailed'
            out.append({
                'period': month_map[r['period']], 'branch_id': r.get('branch_id'),
                'reporting_group': r['reporting_group'],
                'reporting_subgroup': r.get('reporting_subgroup') if detailed else None,
                'source_account_code': r.get('source_account_code') if detailed else None,
                'amount': Decimal(r['amount'] or 0) * factor,
                'driver_type': 'previous_budget', 'driver_value': request.seed_growth_percent,
                'notes': 'Seeded from a previous plan version.'
            })
        if request.detail_level != 'detailed':
            agg: dict[tuple, Decimal] = defaultdict(lambda: Decimal('0'))
            for r in out:
                agg[(r['period'], r['branch_id'], r['reporting_group'])] += r['amount']
            out = [
                {'period': k[0], 'branch_id': k[1], 'reporting_group': k[2], 'reporting_subgroup': None,
                 'source_account_code': None, 'amount': v, 'driver_type': 'previous_budget',
                 'driver_value': request.seed_growth_percent, 'notes': 'High-level seed from previous plan.'}
                for k, v in agg.items()
            ]
        return out

    async def _build_seed(self, company_id: UUID, request) -> list[dict]:
        if request.seed_mode == 'blank':
            return []
        if request.seed_mode == 'previous_budget':
            return await self._previous_seed(company_id, request)
        return await self._actual_seed(company_id, request.financial_year_start, request.financial_year_end, request.seed_growth_percent, request.detail_level)

    async def _insert_lines(self, company_id: UUID, version_id: UUID, rows: list[dict]) -> None:
        if not rows:
            return
        payload = [{'version_id': version_id, 'company_id': company_id, **r} for r in rows]
        await self.session.execute(text("""
            INSERT INTO public.native_plan_lines(
                version_id,company_id,period,branch_id,reporting_group,reporting_subgroup,
                source_account_code,amount,driver_type,driver_value,notes)
            VALUES (:version_id,:company_id,:period,:branch_id,:reporting_group,:reporting_subgroup,
                    :source_account_code,:amount,:driver_type,:driver_value,:notes)
            ON CONFLICT DO NOTHING
        """), payload)

    async def create_version(self, company_id, request):
        await self._ensure_schema()
        assumptions = {
            'seed_mode': request.seed_mode,
            'detail_level': request.detail_level,
            'allocation_method': request.allocation_method,
            'seed_growth_percent': float(request.seed_growth_percent),
            'seed_version_id': str(request.seed_version_id) if request.seed_version_id else None,
            'seed_imported_version': request.seed_imported_version,
        }
        version_id = (await self.session.execute(text("""
            INSERT INTO public.planning_versions(company_id,plan_type,version_name,financial_year_start,financial_year_end,assumptions)
            VALUES (:company_id,:plan_type,:version_name,:start,:end,CAST(:assumptions AS jsonb)) RETURNING id
        """), {'company_id': company_id, 'plan_type': request.plan_type, 'version_name': request.version_name,
                 'start': request.financial_year_start, 'end': request.financial_year_end,
                 'assumptions': json.dumps(assumptions)})).scalar_one()
        rows = await self._build_seed(company_id, request)
        await self._insert_lines(company_id, version_id, rows)
        await self.session.commit()
        return await self.get_version(company_id, version_id)

    async def list_versions(self, company_id):
        await self._ensure_schema()
        result = await self.session.execute(text("""
            SELECT id,plan_type,version_name,financial_year_start,financial_year_end,status,assumptions
            FROM public.planning_versions WHERE company_id=:company_id ORDER BY updated_at DESC
        """), {'company_id': company_id})
        return [dict(row) for row in result.mappings().all()]

    async def get_version(self, company_id, version_id):
        await self._ensure_schema()
        version = (await self.session.execute(text("""
            SELECT id,plan_type,version_name,financial_year_start,financial_year_end,status,assumptions
            FROM public.planning_versions WHERE company_id=:company_id AND id=:id
        """), {'company_id': company_id, 'id': version_id})).mappings().one()
        lines = [dict(row) for row in (await self.session.execute(text("""
            SELECT id,period,branch_id,reporting_group,reporting_subgroup,source_account_code,
                   amount,driver_type,driver_value,notes
            FROM public.native_plan_lines WHERE company_id=:company_id AND version_id=:id
            ORDER BY period,reporting_group,source_account_code NULLS FIRST
        """), {'company_id': company_id, 'id': version_id})).mappings().all()]
        return {**dict(version), 'lines': lines}

    async def save_lines(self, company_id, version_id, lines):
        await self._ensure_schema()
        await self.session.execute(text('DELETE FROM public.native_plan_lines WHERE company_id=:company_id AND version_id=:id'), {'company_id': company_id, 'id': version_id})
        rows = [{'version_id': version_id, 'company_id': company_id, **line.model_dump()} for line in lines]
        await self._insert_lines(company_id, version_id, rows)
        await self.session.execute(text('UPDATE public.planning_versions SET updated_at=now() WHERE id=:id AND company_id=:company_id'), {'id': version_id, 'company_id': company_id})
        await self.session.commit()
        return await self.get_version(company_id, version_id)

    async def reseed(self, company_id: UUID, version_id: UUID):
        version = await self.get_version(company_id, version_id)
        from types import SimpleNamespace
        assumptions = version.get('assumptions') or {}
        request = SimpleNamespace(
            plan_type=version['plan_type'], version_name=version['version_name'],
            financial_year_start=version['financial_year_start'], financial_year_end=version['financial_year_end'],
            seed_mode=assumptions.get('seed_mode', 'actuals'), detail_level=assumptions.get('detail_level', 'high_level'),
            allocation_method=assumptions.get('allocation_method', 'actuals_ratio'),
            seed_growth_percent=Decimal(str(assumptions.get('seed_growth_percent', 0))),
            seed_version_id=UUID(assumptions['seed_version_id']) if assumptions.get('seed_version_id') else None,
            seed_imported_version=assumptions.get('seed_imported_version'),
        )
        rows = await self._build_seed(company_id, request)
        await self.session.execute(text('DELETE FROM public.native_plan_lines WHERE company_id=:company_id AND version_id=:id'), {'company_id': company_id, 'id': version_id})
        await self._insert_lines(company_id, version_id, rows)
        await self.session.commit()
        return await self.get_version(company_id, version_id)

    async def allocate_high_level(self, company_id: UUID, version_id: UUID, request):
        version = await self.get_version(company_id, version_id)
        months = month_range(version['financial_year_start'], version['financial_year_end'])
        history = await self._account_history(company_id, request.branch_id)
        source_months = sorted({r['month'] for r in history})[-12:]
        history = [r for r in history if r['month'] in source_months]
        by_group_month: dict[str, dict[date, Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal('0')))
        by_group_account: dict[str, dict[tuple[str, str | None, str | None], Decimal]] = defaultdict(lambda: defaultdict(lambda: Decimal('0')))
        historical_totals: dict[str, Decimal] = defaultdict(lambda: Decimal('0'))
        for r in history:
            amt = Decimal(r['amount'] or 0); group = r['reporting_group']
            by_group_month[group][r['month']] += amt
            by_group_account[group][(r['source_account_code'], r['account_name'], r['reporting_subgroup'])] += amt
            historical_totals[group] += amt

        targets = {k: Decimal(v) for k, v in (request.annual_targets or {}).items()}
        if request.revenue_target is not None:
            targets['Revenue'] = Decimal(request.revenue_target)
        if request.gross_margin_percent is not None and 'Revenue' in targets:
            gp = Decimal(request.gross_margin_percent) / Decimal('100')
            targets['Cost of Sales'] = targets['Revenue'] * (Decimal('1') - gp)
        if request.net_profit_target is not None and 'Revenue' in targets:
            gp_value = targets['Revenue'] - targets.get('Cost of Sales', Decimal('0'))
            non_operating = sum((targets.get(g, Decimal('0')) for g in ('Depreciation','Finance Costs','Tax','Other Expenses')), Decimal('0'))
            other_income = targets.get('Other Income', Decimal('0'))
            opex_envelope = max(Decimal('0'), gp_value + other_income - non_operating - Decimal(request.net_profit_target))
            payroll_hist = abs(historical_totals.get('Payroll', Decimal('0')))
            opex_hist = abs(historical_totals.get('Operating Expenses', Decimal('0')))
            denom = payroll_hist + opex_hist
            payroll_share = payroll_hist / denom if denom else Decimal('0.5')
            targets.setdefault('Payroll', opex_envelope * payroll_share)
            targets.setdefault('Operating Expenses', opex_envelope * (Decimal('1') - payroll_share))

        new_rows = []
        for group, target in targets.items():
            values = [by_group_month[group].get(m, Decimal('0')) for m in source_months]
            month_weights = monthly_weights_from_history(values, len(months)) if request.seasonality == 'historical' else [Decimal('1')] * len(months)
            month_amounts = allocate_total(Decimal(target), month_weights)
            if request.detail_level == 'detailed' and by_group_account.get(group):
                accounts = list(by_group_account[group].items())
                acc_weights = [abs(v) for _, v in accounts] if request.allocation_method == 'historical_actuals' else [Decimal('1')] * len(accounts)
                for period, month_amount in zip(months, month_amounts):
                    splits = allocate_total(month_amount, acc_weights)
                    for ((code, name, subgroup), _), amount in zip(accounts, splits):
                        new_rows.append({'period': period, 'branch_id': request.branch_id, 'reporting_group': group,
                                         'reporting_subgroup': subgroup or name, 'source_account_code': code,
                                         'amount': amount, 'driver_type': 'executive_target_allocation', 'driver_value': None,
                                         'notes': 'Derived from management targets and allocated using the selected historical method.'})
            else:
                for period, amount in zip(months, month_amounts):
                    new_rows.append({'period': period, 'branch_id': request.branch_id, 'reporting_group': group,
                                     'reporting_subgroup': None, 'source_account_code': None, 'amount': amount,
                                     'driver_type': 'executive_target_allocation', 'driver_value': None,
                                     'notes': 'Derived from high-level management targets.'})
        for group in targets:
            await self.session.execute(text("""DELETE FROM public.native_plan_lines
                WHERE company_id=:company_id AND version_id=:id AND reporting_group=:group
                  AND ((:branch_id IS NULL AND branch_id IS NULL) OR branch_id=:branch_id)"""),
                {'company_id': company_id, 'id': version_id, 'group': group, 'branch_id': request.branch_id})
        await self._insert_lines(company_id, version_id, new_rows)
        await self.session.execute(text("""UPDATE public.planning_versions
            SET assumptions = assumptions || CAST(:patch AS jsonb), updated_at=now()
            WHERE id=:id AND company_id=:company_id"""), {
                'id': version_id, 'company_id': company_id,
                'patch': json.dumps({'last_target_model': {
                    'branch_id': str(request.branch_id) if request.branch_id else None,
                    'revenue_target': float(request.revenue_target) if request.revenue_target is not None else None,
                    'gross_margin_percent': float(request.gross_margin_percent) if request.gross_margin_percent is not None else None,
                    'net_profit_target': float(request.net_profit_target) if request.net_profit_target is not None else None,
                    'seasonality': request.seasonality, 'allocation_method': request.allocation_method,
                    'detail_level': request.detail_level,
                }})
            })
        await self.session.commit()
        return await self.get_version(company_id, version_id)
