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

  async getAccess(): Promise<{ role: string; can_write_finance: boolean; can_reset_all: boolean; can_manage_members: boolean }> {
    return (await api.get<ApiResponse<{ role: string; can_write_finance: boolean; can_reset_all: boolean; can_manage_members: boolean }>>("/companies/me/access")).data.data;
  },

  async getMembers(): Promise<Array<{id:string;user_id:string;role:string;is_active:boolean;joined_at:string;full_name:string;job_title?:string|null}>> {
    return (await api.get<ApiResponse<Array<{id:string;user_id:string;role:string;is_active:boolean;joined_at:string;full_name:string;job_title?:string|null}>>>("/companies/me/members")).data.data;
  },

  async updateMemberRole(memberId: string, role: string): Promise<void> {
    await api.patch(`/companies/me/members/${memberId}/role`, { role });
  },
  async accessMe(){return (await api.get<ApiResponse<any>>("/access/me")).data.data;},
  async accessMembers(){return (await api.get<ApiResponse<any[]>>("/access/members")).data.data;},
  async accessInvitations(){return (await api.get<ApiResponse<any[]>>("/access/invitations")).data.data;},
  async createAccessInvitation(payload:{email:string;role:string;expires_in_days?:number}){return (await api.post<ApiResponse<any>>("/access/invitations",payload)).data.data;},
  async revokeAccessInvitation(id:string){await api.post(`/access/invitations/${id}/revoke`);},
  async acceptAccessInvitation(token:string){return (await api.post<ApiResponse<any>>("/access/invitations/accept",{token})).data.data;},
  async getPersonalProfile(){return (await api.get<ApiResponse<any>>("/access/profile")).data.data;},
  async updatePersonalProfile(payload:{full_name:string;job_title?:string|null}){return (await api.put<ApiResponse<any>>("/access/profile",payload)).data.data;},
  async accessRole(id:string,role:string){await api.patch(`/access/members/${id}/role`,{role});},
  async deactivateAccessMember(id:string){await api.post(`/access/members/${id}/deactivate`);},

};