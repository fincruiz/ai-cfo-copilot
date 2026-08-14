import { api } from "@/lib/api"; import type { ApiResponse } from "@/types/auth";
export interface AuditEvent { id:string; user_id:string|null; action:string; module:string; summary:string; metadata:Record<string,unknown>; created_at:string; }
export const auditService={ async list(limit=100):Promise<AuditEvent[]>{ return (await api.get<ApiResponse<AuditEvent[]>>(`/audit/events?limit=${limit}`)).data.data; } };
