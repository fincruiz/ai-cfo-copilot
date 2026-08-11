import { api } from "@/lib/api";
import type { ApiResponse } from "@/types/finance";
import type { PlanImportResult, VarianceLine } from "@/types/planning";

async function upload(endpoint: string, file: File, versionName: string) {
  const body = new FormData();
  body.append("file", file);
  body.append("version_name", versionName);
  body.append("replace_existing", "true");
  return (await api.post<ApiResponse<PlanImportResult>>(endpoint, body, {
    headers: { "Content-Type": "multipart/form-data" },
    timeout: 120000,
  })).data.data;
}
export const planningService = {
  uploadBudget(file: File, versionName = "Default") {
    return upload("/planning/budget", file, versionName);
  },
  uploadForecast(file: File, versionName = "Default") {
    return upload("/planning/forecast", file, versionName);
  },
  async getVariance(): Promise<VarianceLine[]> {
    return (await api.get<ApiResponse<VarianceLine[]>>("/planning/variance")).data.data;
  },
};
