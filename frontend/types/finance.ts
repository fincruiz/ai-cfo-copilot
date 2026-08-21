export interface ApiResponse<T> {
  success: boolean;
  message: string;
  data: T;
}

export interface ValidationIssue {
  row_number?: number | null;
  column?: string | null;
  message: string;
  severity: string;
}

export interface GLValidationSummary {
  required_columns: string[];
  detected_columns: string[];
  missing_columns: string[];
  total_rows: number;
  valid_rows: number;
  invalid_rows: number;
  issues: ValidationIssue[];
}

export interface FileUploadRecord {
  id: string;
  company_id: string;
  original_file_name?: string | null;
  processing_status: string;
  row_count?: number | null;
  valid_row_count?: number | null;
  invalid_row_count?: number | null;
  column_mapping: Record<string, string>;
  created_at: string;
}

export interface GLUploadResult {
  upload: FileUploadRecord;
  validation: GLValidationSummary;
  inserted_transaction_count: number;
}

export interface MappingSuggestion {
  source_account_code: string;
  source_account_name?: string | null;
  statement: string;
  reporting_group: string;
  reporting_subgroup?: string | null;
  sign_convention: string;
  confidence: number;
  reason: string;
}

export interface AccountMappingInput {
  source_account_code: string;
  source_account_name?: string | null;
  statement: string;
  reporting_group: string;
  reporting_subgroup?: string | null;
  sign_convention: string;
  display_order?: number | null;
  is_confirmed: boolean;
}

export interface AccountMapping extends AccountMappingInput {
  id: string;
  company_id: string;
  created_at: string;
  updated_at: string;
}

export interface Branch {
  id: string;
  company_id: string;
  branch_code: string;
  branch_name: string;
  region?: string | null;
  review_status: string;
  source_value?: string | null;
  discovered_from_upload_id?: string | null;
  is_active: boolean;
  created_at: string;
  updated_at: string;
}

export interface BranchInput {
  branch_code: string;
  branch_name: string;
  region?: string | null;
}

export interface ReportLine {
  code: string;
  label: string;
  amount: string | number;
  order: number;
  is_total: boolean;
}

export interface TrialBalance {
  total_debit: string | number;
  total_credit: string | number;
  difference: string | number;
  lines: ReportLine[];
}

export interface ProfitAndLoss {
  revenue: string | number;
  cost_of_sales: string | number;
  gross_profit: string | number;
  operating_expenses: string | number;
  operating_profit: string | number;
  depreciation: string | number;
  ebit: string | number;
  other_income: string | number;
  other_expenses: string | number;
  finance_costs: string | number;
  profit_before_tax: string | number;
  tax: string | number;
  net_profit: string | number;
  lines: ReportLine[];
}

export interface BalanceSheet {
  current_assets: string | number;
  non_current_assets: string | number;
  total_assets: string | number;
  current_liabilities: string | number;
  non_current_liabilities: string | number;
  total_liabilities: string | number;
  contributed_equity: string | number;
  current_period_earnings: string | number;
  equity: string | number;
  total_liabilities_and_equity: string | number;
  balance_difference: string | number;
  lines: ReportLine[];
}

export interface Ratio {
  name: string;
  category: string;
  value?: string | number | null;
  unit: string;
  status: string;
  tone: string;
  interpretation: string;
}

export interface MonthlyActual {
  month: string;
  revenue: string | number;
  cost_of_sales: string | number;
  gross_profit: string | number;
  operating_expenses: string | number;
  depreciation: string | number;
  ebit: string | number;
  finance_costs: string | number;
  tax: string | number;
  net_profit: string | number;
}

export interface BranchComparison {
  branch_id: string;
  branch_code: string;
  branch_name: string;
  revenue: string | number;
  gross_profit: string | number;
  operating_expenses: string | number;
  ebit: string | number;
  net_profit: string | number;
  gross_margin_percent?: string | number | null;
  net_margin_percent?: string | number | null;
}

export interface ForecastPoint {
  period: string;
  base: string | number;
  downside: string | number;
  upside: string | number;
}

export interface ForecastResult {
  reporting_group: string;
  method: string;
  branch_id?: string | null;
  history_periods: number;
  confidence: string;
  warning?: string | null;
  history?: Array<{ period: string; actual: string | number }>;
  points: ForecastPoint[];
}

export interface DataHealth {
  transaction_count: number;
  upload_count: number;
  active_upload_count: number;
  account_count: number;
  mapped_account_count: number;
  unmapped_account_count: number;
  invalid_transaction_count: number;
  duplicate_candidate_count: number;
  first_transaction_date?: string | null;
  last_transaction_date?: string | null;
  total_debit: string | number;
  total_credit: string | number;
  trial_balance_difference: string | number;
  balance_sheet_difference: string | number;
  is_trial_balance_balanced: boolean;
  is_balance_sheet_balanced: boolean;
  is_mapping_complete: boolean;
  overall_status: "empty" | "healthy" | "attention_required" | string;
}

export interface FinancialAssuranceCheck {
  key: string;
  label: string;
  status: "pass" | "warning" | "fail" | string;
  score: number;
  detail: string;
  action?: string | null;
}

export interface FinancialAssurance {
  score: number;
  grade: string;
  status: "ready" | "review" | "not_ready" | string;
  checks: FinancialAssuranceCheck[];
  caveat: string;
}


export interface IngestionJob {
  id: string; company_id: string; job_type: string; original_file_name: string; file_size_bytes: number;
  source_system?: string | null; status: string; progress_percent: number; phase: string; total_rows?: number | null;
  valid_rows?: number | null; invalid_rows?: number | null; inserted_rows: number; file_upload_id?: string | null;
  error_message?: string | null; attempts: number; created_at: string; started_at?: string | null; completed_at?: string | null; updated_at: string;
}


export interface FinanceReliabilityCheck {
  key: string;
  label: string;
  status: "pass" | "warning" | "fail" | string;
  detail: string;
  action?: string | null;
  category: string;
  blocking: boolean;
}

export interface FinanceReliability {
  status: "ready" | "attention" | "blocked" | string;
  score: number;
  pass_count: number;
  warning_count: number;
  fail_count: number;
  checks: FinanceReliabilityCheck[];
  active_upload_id?: string | null;
  first_transaction_date?: string | null;
  last_transaction_date?: string | null;
  assurance_score: number;
  assurance_grade: string;
  certified_at: string;
  caveat: string;
}


export interface ReportContext {
  period_start?: string | null;
  period_end?: string | null;
  data_as_of?: string | null;
  transaction_count: number;
  branch_id?: string | null;
}

export interface LedgerTransaction {
  id: string;
  transaction_date: string;
  source_account_code: string;
  source_account_name?: string | null;
  description?: string | null;
  document_number?: string | null;
  debit: string | number;
  credit: string | number;
  branch_id?: string | null;
  external_reference?: string | null;
}
