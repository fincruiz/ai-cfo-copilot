export interface ForecastRun { run_id:string; run_name:string; summary:Record<string,number|boolean>; profit_and_loss:Array<Record<string,unknown>>; balance_sheet:Array<Record<string,unknown>>; cash_flow:Array<Record<string,unknown>>; ratios:Array<Record<string,unknown>>; checks:Array<Record<string,unknown>>; scenarios:Array<Record<string,unknown>>; diagnostics:Array<Record<string,unknown>>; }
export interface PlanningVersion {id:string;plan_type:string;version_name:string;financial_year_start:string;financial_year_end:string;status:string;assumptions:Record<string,unknown>;lines?:Array<Record<string,any>>}
export interface Artifact {id:string;artifact_type:string;file_name:string;download_url:string;file_size_bytes:number}

export interface DecisionSimulation {
  scenario_name: string;
  assumptions: Record<string, number>;
  base_summary: Record<string, number | boolean | string | null>;
  scenario_summary: Record<string, number | boolean | string | null>;
  impact: Record<string, number>;
  assessment: { level: 'green' | 'amber' | 'red'; title: string; message: string; minimum_cash_target: number };
  comparison_series: Array<{ period: string; base_revenue: number; scenario_revenue: number; base_net_income: number; scenario_net_income: number; base_cash: number; scenario_cash: number }>;
  base_checks: Array<Record<string, unknown>>;
  scenario_checks: Array<Record<string, unknown>>;
}

export interface PlanningContext {
  actual_months:number; first_actual_month?:string|null; latest_actual_month?:string|null; mapped_accounts:number;
  native_versions:PlanningVersion[]; imported_versions:Array<{plan_type:string;version_name:string;first_period:string;last_period:string;line_count:number}>;
  recommended_seed:string;
}
export interface PlanningBaseline {
  history_months:number; period_start:string; period_end:string; suggested_forecast_start:string; trailing_revenue:number; trailing_cogs:number; trailing_payroll:number; trailing_opex:number; trailing_net_profit:number; gross_margin_percent:number; payroll_percent_revenue:number; other_opex_percent_revenue:number;
  suggested_drivers?:{gross_margin:number;payroll_pct_revenue:number;other_opex_pct_revenue:number;dso_days:number;dpo_days:number;inventory_days:number};
  opening_balance_sheet:Record<string,number|null>;
  monthly:Array<{month:string;revenue:number;cost_of_sales:number;payroll:number;operating_expenses:number;net_profit_proxy:number}>;
}
