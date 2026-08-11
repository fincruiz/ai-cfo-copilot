# FinCruiz finance adoption summary

The legacy finance module has been adopted into the FastAPI backend as database-native, tenant-aware code.

## Production modules added

- Validated CSV -> `gl_transactions` ingestion
- Company-specific account mappings with deterministic legacy-style suggestions
- Trial balance
- Profit and loss with legacy reporting groups/subtotals
- Balance sheet and balance check
- Management KPI/ratio engine with legacy status thresholds
- Monthly reporting-group aggregation
- Run-rate and trend forecasts with base/upside/downside scenarios
- FastAPI routes for mappings, reports, KPIs and forecasts
- SQL migration and pure-domain regression tests

## API routes

- `POST /api/v1/uploads/general-ledger`
- `GET /api/v1/account-mappings`
- `GET /api/v1/account-mappings/suggestions`
- `PUT /api/v1/account-mappings`
- `GET /api/v1/reports/trial-balance`
- `GET /api/v1/reports/profit-and-loss`
- `GET /api/v1/reports/balance-sheet`
- `GET /api/v1/reports/kpis`
- `POST /api/v1/forecasts`

## Required database step

Run `migrations/20260801_finance_engine.sql` in Supabase before using mapping/report endpoints.

## Verification performed

- All Python source compiled successfully.
- Pure finance smoke tests produced gross profit 6000, balanced balance sheet difference 0, correct bank mapping, and a run-rate forecast.
- No pandas dependency was added to the production finance engine.
