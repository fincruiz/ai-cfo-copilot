import { api } from "@/lib/api";
export type Readiness = { status: "healthy"|"degraded"|"unhealthy"; version: string; environment: string; checks: { api: {status:string}; database?: {status:string; latency_ms:number} } };
export const systemHealthService = { async readiness(): Promise<Readiness> { return (await api.get<Readiness>("/health/readiness", { timeout: 8000 })).data; } };
