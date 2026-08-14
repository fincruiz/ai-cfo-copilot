import { api } from "@/lib/api";
import type { IntegrationConnection } from "@/types/integrations";

type ApiResponse<T> = { success: boolean; message: string; data: T };

export const integrationService = {
  list: async () =>
    (await api.get<ApiResponse<IntegrationConnection[]>>("/integrations")).data.data,

  start: async (provider: "xero" | "zoho") =>
    (
      await api.post<ApiResponse<{ authorization_url: string }>>(
        `/integrations/${provider}/start`,
      )
    ).data.data,

  sync: async (provider: "xero" | "zoho") =>
    (
      await api.post<ApiResponse<Record<string, number>>>(
        `/integrations/${provider}/sync`,
        undefined,
        { timeout: 300000 },
      )
    ).data.data,

  selectTenant: async (provider: "xero" | "zoho", tenant_id: string) =>
    api.post(`/integrations/${provider}/select-tenant`, { tenant_id }),

  disconnect: async (provider: string) => api.delete(`/integrations/${provider}`),

  createTallyToken: async () =>
    (
      await api.post<ApiResponse<{ bridge_token: string; ingest_url: string }>>(
        "/integrations/tally/bridge-token",
      )
    ).data.data,
};
