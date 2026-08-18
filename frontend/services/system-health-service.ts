import { api } from "@/lib/api";

export type Readiness = {
  status: "healthy" | "degraded" | "unhealthy";
  version: string;
  environment: string;
  checks: {
    api: { status: string };
    database?: { status: string; latency_ms: number };
  };
};

export type OperationalReadiness = {
  status: "healthy" | "degraded" | "unhealthy";
  score: number;
  checks: Array<{
    key: string;
    label: string;
    status: "healthy" | "degraded" | "unhealthy" | string;
    detail: string;
    action?: string | null;
  }>;
  database_latency_ms: number;
  ingestion_open_jobs: number;
  ingestion_stale_jobs: number;
  ingestion_recent_failures: number;
  active_gl_datasets: number;
  latest_ingestion_update_at?: string | null;
  checked_at: string;
};

export const systemHealthService = {
  async readiness(): Promise<Readiness> {
    return (
      await api.get<Readiness>("/health/readiness", { timeout: 8000 })
    ).data;
  },

  async operations(): Promise<OperationalReadiness> {
    return (
      await api.get<{ success: boolean; message: string; data: OperationalReadiness }>(
        "/operations/readiness",
        { timeout: 10000 },
      )
    ).data.data;
  },
};
