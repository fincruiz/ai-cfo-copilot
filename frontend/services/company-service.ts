import { api } from "@/lib/api";

import type {
  ApiResponse,
  Company,
  CreateCompanyRequest,
} from "@/types/auth";

export const companyService = {
  async createCompany(
    payload: CreateCompanyRequest,
  ): Promise<Company> {
    const response = await api.post<
      ApiResponse<Company>
    >("/companies", payload);

    return response.data.data;
  },

  async uploadLogo(file: File): Promise<Company> {
    const body = new FormData();
    body.append("file", file);

    const response = await api.post<ApiResponse<Company>>(
      "/companies/me/logo",
      body,
      {
        headers: { "Content-Type": "multipart/form-data" },
        timeout: 60000,
      },
    );

    return response.data.data;
  },

  async updateCompany(payload: Partial<CreateCompanyRequest>): Promise<Company> {
    return (await api.put<ApiResponse<Company>>("/companies/me", payload)).data.data;
  },

  async getPreferences() {
    return (await api.get<ApiResponse<Record<string, unknown>>>("/companies/me/preferences")).data.data;
  },

  async updatePreferences(payload: Record<string, unknown>) {
    return (await api.put<ApiResponse<Record<string, unknown>>>("/companies/me/preferences", payload)).data.data;
  },

  async getCurrentCompany(): Promise<Company> {
    const response = await api.get<
      ApiResponse<Company>
    >("/companies/me");

    return response.data.data;
  },
};