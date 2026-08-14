import { api } from "@/lib/api";
import type { BrainOverview } from "@/types/integrations";
import type { AICFOAnswer } from "@/types/analytics";

type ApiResponse<T> = { success: boolean; message: string; data: T };

export const intelligenceService = {
  overview: async () =>
    (await api.get<ApiResponse<BrainOverview>>("/intelligence/overview")).data.data,

  addMemory: async (payload: {
    title: string;
    content: string;
    memory_type?: string;
    importance?: string;
  }) =>
    (await api.post<ApiResponse<any>>("/intelligence/memory", payload)).data.data,

  ask: async (question: string, includeExternalContext = true) =>
    (
      await api.post<ApiResponse<AICFOAnswer>>("/ai-cfo/ask", {
        question,
        include_external_context: includeExternalContext,
      })
    ).data.data,
};
