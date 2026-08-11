# Finance engine adoption

This integration preserves the useful legacy finance rules while replacing Streamlit/pandas orchestration with tenant-safe PostgreSQL/FastAPI modules.

## Adopted
- GL validation, normalisation and transaction ingestion
- COA mapping review through deterministic suggestions and confirmed company mappings
- Trial balance, P&L, balance sheet
- Management ratios/KPIs and health statuses
- Basic run-rate/trend scenarios as the foundation for the legacy three-way forecasting migration

## Deliberately not imported
- Streamlit UI, session state, local SQLite, direct Excel reads
- pandas dependencies (blocked by local Windows policy and unsuitable as the core SaaS persistence layer)
- Word/PowerPoint export adapters and AI gateways; these remain later adapters around the new report objects

## Setup
1. Run `migrations/20260801_finance_engine.sql` in Supabase.
2. Restart FastAPI.
3. Upload a valid GL. Valid uploads now insert `gl_transactions`.
4. Call `/api/v1/account-mappings/suggestions`, confirm mappings with `PUT /api/v1/account-mappings`.
5. Use `/api/v1/reports/*`.

## Legacy equivalence
- `prepare_data`: validator + parser + database + mapping service
- `build_pnl`: `domain/finance/reporting/pnl.py`
- `build_balance_sheet_from_gl`: `balance_sheet.py`
- `calculate_management_ratios`: `kpis/ratio_engine.py`
- forecast trend/run-rate/scenarios: `forecasting/engine.py`
