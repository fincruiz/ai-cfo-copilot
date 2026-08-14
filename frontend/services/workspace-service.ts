import { api } from "@/lib/api";
import type { ApiResponse } from "@/types/auth";

export interface WorkspaceStatus {
  has_financial_data: boolean;
  demo_data_active: boolean;
  upload_count: number;
  transaction_count: number;
  mapping_count: number;
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
