# Branch and Forecast Update

1. Run `backend/migrations/20260802_branches_monthly_forecasting.sql` in Supabase.
2. Install backend requirements.
3. Start backend and frontend.
4. Create branches at `/dashboard/branches`.
5. Upload a CSV containing `branch_code`, `branch`, `location`, or `business_unit`.
6. Review `/dashboard/reports` for consolidated/branch views and monthly actuals.
7. Generate forecasts at `/dashboard/forecasting`.

Unknown branch values cause upload validation to fail intentionally. Create branch records first.
