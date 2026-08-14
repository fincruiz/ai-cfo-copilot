export type Provider = "xero" | "zoho" | "tally";

export type IntegrationConnection = {
  provider: Provider;
  status: string;
  configured?: boolean;
  external_tenant_id?: string | null;
  external_tenant_name?: string | null;
  last_synced_at?: string | null;
  last_sync_status?: string | null;
  last_sync_message?: string | null;
  metadata?: Record<string, any>;
};

export type ManagementMemory = {
  id: string;
  title: string;
  content: string;
  memory_type: string;
  importance: string;
  created_at: string;
};

export type IntelligencePriority = {
  level: "critical" | "attention" | "positive" | "monitor" | string;
  title: string;
  evidence: string;
  action: string;
  source: string;
};

export type IntelligenceMetric = {
  key: string;
  label: string;
  value: number | null;
  format: "currency" | "percent" | "score" | string;
  change?: number | null;
  change_unit?: "percent" | "points" | "of_ar" | string | null;
  context: string;
};

export type MonthlyTrendPoint = {
  month: string;
  revenue: number;
  gross_profit: number;
  net_profit: number;
  gross_margin?: number | null;
};

export type SourceFreshness = {
  provider: string;
  name: string;
  status: string;
  last_synced_at?: string | null;
  last_sync_status?: string | null;
  last_sync_message?: string | null;
};

export type BrainOverview = {
  company: {
    name: string;
    currency: string;
    industry?: string | null;
    business_model?: string | null;
  };
  executive_summary: {
    headline: string;
    narrative: string;
    critical_count: number;
    attention_count: number;
    positive_count: number;
    generated_at: string;
  };
  financial_snapshot: IntelligenceMetric[];
  monthly_trends: MonthlyTrendPoint[];
  priorities: IntelligencePriority[];
  connections: IntegrationConnection[];
  source_counts: Array<{ provider: string; entity_type: string; count: number }>;
  source_freshness: SourceFreshness[];
  memories: ManagementMemory[];
  signals: Array<{
    severity: string;
    title: string;
    evidence: string;
    action: string;
  }>;
  financial_assurance: any;
  working_capital: {
    receivables?: { total: number; overdue: number; overdue_percent: number } | null;
    payables?: { total: number; overdue: number } | null;
  };
  suggested_questions: string[];
};
