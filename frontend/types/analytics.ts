export interface ImportIssue {
  row_number?: number | null;
  column?: string | null;
  message: string;
  severity: string;
}

export interface FinanceImportResult {
  import_type: string;
  original_file_name: string;
  total_rows: number;
  valid_rows: number;
  invalid_rows: number;
  inserted_rows: number;
  issues: ImportIssue[];
  metadata: Record<string, unknown>;
}

export interface AgeingBucket {
  bucket: string;
  amount: string | number;
  document_count: number;
}

export interface PartyExposure {
  party_name: string;
  outstanding_amount: string | number;
  overdue_amount: string | number;
  document_count: number;
  oldest_due_date?: string | null;
  weighted_days_overdue?: string | number | null;
}

export interface WorkingCapitalSummary {
  ageing_type: string;
  total_outstanding: string | number;
  overdue_amount: string | number;
  overdue_percent: string | number;
  current_amount: string | number;
  document_count: number;
  party_count: number;
  weighted_average_days_overdue?: string | number | null;
  buckets: AgeingBucket[];
  top_parties: PartyExposure[];
}

export interface AnalyticsOverview {
  monthly_actuals: Array<Record<string, string | number>>;
  branch_comparison: Array<Record<string, string | number>>;
  ar_summary?: WorkingCapitalSummary | null;
  ap_summary?: WorkingCapitalSummary | null;
  insights: string[];
}

export interface AICFOAnswer {
  answer: string;
  mode: string;
  suggested_questions: string[];
}
