export interface PlanImportResult {
  plan_type: string;
  version_name: string;
  total_rows: number;
  inserted_rows: number;
  invalid_rows: number;
  issues: Array<{ row_number?: number; message: string }>;
}
export interface VarianceLine {
  period: string;
  reporting_group: string;
  actual: string | number;
  budget: string | number;
  forecast: string | number;
  budget_variance: string | number;
  budget_variance_percent?: string | number | null;
  forecast_variance: string | number;
  forecast_variance_percent?: string | number | null;
}
