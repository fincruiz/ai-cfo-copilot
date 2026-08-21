import { api } from "@/lib/api";
import type {
  AccountMapping,
  AccountMappingInput,
  ApiResponse,
  BalanceSheet,
  Branch,
  BranchComparison,
  BranchInput,
  DataHealth,
  ForecastResult,
  GLUploadResult,
  IngestionJob,
  MappingSuggestion,
  MonthlyActual,
  ProfitAndLoss,
  Ratio,
  TrialBalance,
  ReportContext,
  LedgerTransaction,
} from "@/types/finance";

function queryString(params?: Record<string, string | undefined>) {
  if (!params) return "";
  const search = new URLSearchParams();
  Object.entries(params).forEach(([key, value]) => {
    if (value) search.set(key, value);
  });
  const value = search.toString();
  return value ? `?${value}` : "";
}

export const financeService = {
  async stageGeneralLedger(file: File, sourceSystem?: string): Promise<IngestionJob> {
    const body = new FormData();
    body.append("file", file);
    if (sourceSystem?.trim()) body.append("source_system", sourceSystem.trim());
    return (await api.post<ApiResponse<IngestionJob>>("/uploads/general-ledger/jobs", body, {
      headers: { "Content-Type": "multipart/form-data" },
      timeout: 10 * 60 * 1000,
    })).data.data;
  },

  async getIngestionJobs(): Promise<IngestionJob[]> {
    return (await api.get<ApiResponse<IngestionJob[]>>("/uploads/general-ledger/jobs")).data.data;
  },

  async getIngestionJob(jobId: string): Promise<IngestionJob> {
    return (await api.get<ApiResponse<IngestionJob>>(`/uploads/general-ledger/jobs/${jobId}`)).data.data;
  },

  async retryIngestionJob(jobId: string): Promise<IngestionJob> {
    return (await api.post<ApiResponse<IngestionJob>>(`/uploads/general-ledger/jobs/${jobId}/retry`)).data.data;
  },

  async uploadGeneralLedger(file: File, sourceSystem?: string): Promise<GLUploadResult> {
    const body = new FormData();
    body.append("file", file);
    if (sourceSystem?.trim()) body.append("source_system", sourceSystem.trim());
    const response = await api.post<ApiResponse<GLUploadResult>>(
      "/uploads/general-ledger",
      body,
      { headers: { "Content-Type": "multipart/form-data" }, timeout: 120000 },
    );
    return response.data.data;
  },

  async getMappings(): Promise<AccountMapping[]> {
    return (await api.get<ApiResponse<AccountMapping[]>>("/account-mappings")).data.data;
  },

  async getMappingSuggestions(): Promise<MappingSuggestion[]> {
    return (await api.get<ApiResponse<MappingSuggestion[]>>(
      "/account-mappings/suggestions",
    )).data.data;
  },

  async saveMappings(items: AccountMappingInput[]): Promise<number> {
    return (await api.put<ApiResponse<{ saved: number }>>(
      "/account-mappings",
      { items },
    )).data.data.saved;
  },

  async getBranches(): Promise<Branch[]> {
    return (await api.get<ApiResponse<Branch[]>>("/branches")).data.data;
  },

  async createBranch(payload: BranchInput): Promise<Branch> {
    return (await api.post<ApiResponse<Branch>>("/branches", payload)).data.data;
  },

  async updateBranch(
    branchId: string,
    payload: Partial<Pick<Branch, "branch_code" | "branch_name" | "region" | "review_status" | "is_active">>,
  ): Promise<Branch> {
    return (await api.put<ApiResponse<Branch>>(`/branches/${branchId}`, payload)).data.data;
  },

  async getReportContext(params?: { branchId?: string }): Promise<ReportContext> {
    return (await api.get<ApiResponse<ReportContext>>(
      `/reports/context${queryString({ branch_id: params?.branchId })}`,
    )).data.data;
  },

  async getLedgerTransactions(params?: {
    startDate?: string;
    endDate?: string;
    accountCode?: string;
    branchId?: string;
    limit?: number;
  }): Promise<LedgerTransaction[]> {
    return (await api.get<ApiResponse<LedgerTransaction[]>>(
      `/reports/transactions${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
        account_code: params?.accountCode,
        branch_id: params?.branchId,
        limit: params?.limit ? String(params.limit) : undefined,
      })}`,
    )).data.data;
  },

  async getTrialBalance(params?: {
    startDate?: string;
    endDate?: string;
    branchId?: string;
  }): Promise<TrialBalance> {
    return (await api.get<ApiResponse<TrialBalance>>(
      `/reports/trial-balance${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
        branch_id: params?.branchId,
      })}`,
    )).data.data;
  },

  async getProfitAndLoss(params?: {
    startDate?: string;
    endDate?: string;
    branchId?: string;
  }): Promise<ProfitAndLoss> {
    return (await api.get<ApiResponse<ProfitAndLoss>>(
      `/reports/profit-and-loss${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
        branch_id: params?.branchId,
      })}`,
    )).data.data;
  },

  async getBalanceSheet(params?: {
    endDate?: string;
    branchId?: string;
  }): Promise<BalanceSheet> {
    return (await api.get<ApiResponse<BalanceSheet>>(
      `/reports/balance-sheet${queryString({
        end_date: params?.endDate,
        branch_id: params?.branchId,
      })}`,
    )).data.data;
  },

  async getDataHealth(): Promise<DataHealth> {
    return (await api.get<ApiResponse<DataHealth>>("/reports/data-health")).data.data;
  },

  async getKpis(params?: {
    startDate?: string;
    endDate?: string;
    periodDays?: number;
    branchId?: string;
  }): Promise<Ratio[]> {
    return (await api.get<ApiResponse<Ratio[]>>(
      `/reports/kpis${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
        period_days: params?.periodDays ? String(params.periodDays) : undefined,
        branch_id: params?.branchId,
      })}`,
    )).data.data;
  },

  async getMonthlyActuals(params?: {
    startDate?: string;
    endDate?: string;
    branchId?: string;
  }): Promise<MonthlyActual[]> {
    return (await api.get<ApiResponse<MonthlyActual[]>>(
      `/reports/monthly-actuals${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
        branch_id: params?.branchId,
      })}`,
    )).data.data;
  },

  async getBranchComparison(params?: {
    startDate?: string;
    endDate?: string;
  }): Promise<BranchComparison[]> {
    return (await api.get<ApiResponse<BranchComparison[]>>(
      `/reports/branch-comparison${queryString({
        start_date: params?.startDate,
        end_date: params?.endDate,
      })}`,
    )).data.data;
  },

  async createForecast(payload: {
    reporting_group: string;
    future_months: number;
    method: string;
    branch_id?: string | null;
    downside_factor: number;
    upside_factor: number;
    recent_months: number;
  }): Promise<ForecastResult> {
    return (await api.post<ApiResponse<ForecastResult>>("/forecasts", payload)).data.data;
  },

  async getFinanceReliability(): Promise<import("@/types/finance").FinanceReliability> {
    return (
      await api.get<ApiResponse<import("@/types/finance").FinanceReliability>>(
        "/reports/reliability",
      )
    ).data.data;
  },

  async getFinancialAssurance() {
    return (await api.get<ApiResponse<import("@/types/finance").FinancialAssurance>>("/reports/assurance")).data.data;
  },
};
