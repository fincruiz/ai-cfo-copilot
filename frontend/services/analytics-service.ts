import { api } from "@/lib/api";
import type { ApiResponse } from "@/types/finance";
import type {
  AICFOAnswer,
  AnalyticsOverview,
  FinanceImportResult,
  WorkingCapitalSummary,
  AICFOSignalsResponse,
} from "@/types/analytics";

async function upload(
  endpoint: string,
  file: File,
  sourceSystem: string,
): Promise<FinanceImportResult> {
  const body = new FormData();
  body.append("file", file);
  body.append("source_system", sourceSystem);
  body.append("replace_existing", "true");

  const response = await api.post<ApiResponse<FinanceImportResult>>(
    endpoint,
    body,
    {
      headers: { "Content-Type": "multipart/form-data" },
      timeout: 120000,
    },
  );

  return response.data.data;
}

export const analyticsService = {
  uploadCoa(file: File, sourceSystem = "Manual CSV") {
    return upload("/imports/coa", file, sourceSystem);
  },

  uploadArAgeing(file: File, sourceSystem = "Manual CSV") {
    return upload("/imports/ar-ageing", file, sourceSystem);
  },

  uploadApAgeing(file: File, sourceSystem = "Manual CSV") {
    return upload("/imports/ap-ageing", file, sourceSystem);
  },

  async getOverview(): Promise<AnalyticsOverview> {
    return (
      await api.get<ApiResponse<AnalyticsOverview>>("/analytics/overview")
    ).data.data;
  },

  async getWorkingCapital(
    type: "AR" | "AP",
  ): Promise<WorkingCapitalSummary | null> {
    return (
      await api.get<ApiResponse<WorkingCapitalSummary | null>>(
        `/analytics/working-capital/${type}`,
      )
    ).data.data;
  },

  async getExecutiveBrief(): Promise<AICFOAnswer> {
    return (await api.get<ApiResponse<AICFOAnswer>>("/ai-cfo/executive-brief")).data.data;
  },

  async getProactiveSignals(): Promise<AICFOSignalsResponse> {
    return (await api.get<ApiResponse<AICFOSignalsResponse>>("/ai-cfo/signals")).data.data;
  },

  async getIndustryBenchmark(): Promise<AICFOAnswer> {
    return (await api.get<ApiResponse<AICFOAnswer>>("/ai-cfo/industry-benchmark")).data.data;
  },

  async askAiCfo(question: string, includeExternalContext = true): Promise<AICFOAnswer> {
    return (
      await api.post<ApiResponse<AICFOAnswer>>("/ai-cfo/ask", {
        question,
        include_external_context: includeExternalContext,
      })
    ).data.data;
  },
};
