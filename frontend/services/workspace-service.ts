import { api } from "@/lib/api";
import type { ApiResponse } from "@/types/auth";

export interface WorkspaceStatus {
  has_financial_data: boolean;
  demo_data_active: boolean;
  upload_count: number;
  transaction_count: number;
  mapping_count: number;
}


export interface LaunchReadinessCheck { key: string; label: string; ready: boolean; detail: string; path: string; }
export interface LaunchReadiness { score: number; completed_steps: number; total_steps: number; checks: LaunchReadinessCheck[]; next_path: string; next_label: string; connected_sources: number; healthy_sources: number; ready_for_management_use: boolean; }


export interface CommercialOnboardingStep { key: string; label: string; complete: boolean; }
export interface CommercialOnboardingSummary {
  stage: string; ready_for_intelligence: boolean; progress_percent: number; completed_steps: number; total_steps: number; steps: CommercialOnboardingStep[];
  transaction_count: number; account_count: number; mapping_count: number; unmapped_account_count: number; branch_count: number; pending_branch_count: number; months_history: number;
  period_start?: string | null; period_end?: string | null; financial_confidence_score?: number | null; financial_confidence_grade?: string | null; financial_checks: Array<Record<string, unknown>>;
  latest_ingestion?: Record<string, unknown> | null;
  briefing?: { executive_summary?: { headline?: string; narrative?: string; critical_count?: number; attention_count?: number; positive_count?: number }; priorities?: Array<{ level:string; title:string; evidence?:string; action?:string; source?:string }>; financial_snapshot?: Array<{key:string;label:string;value:number|null;format:string;context?:string}>; monthly_trends?: Array<Record<string, unknown>>; suggested_questions?: string[] } | null;
  briefing_error?: string | null; next_path: string; next_label: string;
}

export interface DemoDataResult {
  upload_id: string;
  months: number;
  transactions_created: number;
  mappings_created: number;
}

export type ResetScope = "general_ledger" | "account_mappings" | "coa" | "ar_ageing" | "ap_ageing" | "planning" | "forecasts" | "board_packs" | "branches";

export interface ResetResult {
  deleted_rows: Record<string, number>;
}

export interface AccountDeletionResult {
  auth_user_deleted: boolean;
  companies_deleted: number;
  memberships_deleted: number;
  profile_deleted: boolean;
}

export const workspaceService = {
  async getStatus(): Promise<WorkspaceStatus> {
    return (await api.get<ApiResponse<WorkspaceStatus>>("/workspace/status")).data.data;
  },

  async getLaunchReadiness(): Promise<LaunchReadiness> {
    return (await api.get<ApiResponse<LaunchReadiness>>("/workspace/launch-readiness")).data.data;
  },

  async getCommercialOnboardingSummary(): Promise<CommercialOnboardingSummary> {
    return (await api.get<ApiResponse<CommercialOnboardingSummary>>("/workspace/commercial-onboarding")).data.data;
  },

  async loadDemo(replaceExisting = false): Promise<DemoDataResult> {
    return (
      await api.post<ApiResponse<DemoDataResult>>("/workspace/demo", {
        replace_existing: replaceExisting,
      })
    ).data.data;
  },

  async resetScope(scope: ResetScope): Promise<ResetResult> {
    return (await api.delete<ApiResponse<ResetResult>>(`/workspace/data/${scope}`, { data: { confirmed: true } })).data.data;
  },

  async resetData(): Promise<ResetResult> {
    return (
      await api.delete<ApiResponse<ResetResult>>("/workspace/data", {
        data: { confirmed: true },
      })
    ).data.data;
  },

  async deleteAccount(): Promise<AccountDeletionResult> {
    return (
      await api.delete<ApiResponse<AccountDeletionResult>>("/account/me", {
        data: { confirmed: true },
      })
    ).data.data;
  },
};
