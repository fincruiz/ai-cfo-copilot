import { api } from "@/lib/api";
import type { ApiResponse } from "@/types/auth";

export type SubscriptionStatus = {
  plan: "trial" | "founding" | "growth" | "enterprise";
  status: "trialing" | "active" | "past_due" | "cancelled" | "expired";
  trial_started_at?: string | null;
  trial_ends_at?: string | null;
  current_period_ends_at?: string | null;
  days_remaining?: number | null;
  entitlements: Record<string, string | number | boolean>;
  is_access_active: boolean;
  billing_managed_externally: boolean;
};

export type BetaReadiness = {
  score: number;
  status: "ready" | "attention" | "blocked";
  checks: Array<{
    key: string;
    label: string;
    status: "ready" | "attention" | "blocked";
    detail: string;
  }>;
};

export const subscriptionService = {
  async status(): Promise<SubscriptionStatus> {
    return (await api.get<ApiResponse<SubscriptionStatus>>("/subscription/status")).data.data;
  },
  async betaReadiness(): Promise<BetaReadiness> {
    return (await api.get<ApiResponse<BetaReadiness>>("/subscription/beta-readiness")).data.data;
  },
};
