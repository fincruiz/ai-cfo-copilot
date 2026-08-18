import { api } from "@/lib/api";

type UsageProperties = Record<string, string | number | boolean | string[] | null | undefined>;

const sessionKey = "fincruiz_usage_session";

function sessionId(): string {
  if (typeof window === "undefined") return "server";
  let value = window.sessionStorage.getItem(sessionKey);
  if (!value) {
    value = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    window.sessionStorage.setItem(sessionKey, value);
  }
  return value;
}

export const usageService = {
  async summary(days = 30): Promise<Array<{ event_name: string; count: number; users: number }>> {
    return (await api.get<{ data: Array<{ event_name: string; count: number; users: number }> }>(`/usage/summary?days=${days}`)).data.data;
  },
  async funnel(days = 30): Promise<{active_users:number;page_views:number;ai_questions:number;upload_starts:number;upload_completions:number;frontend_errors:number}> {
    return (await api.get<{ data: {active_users:number;page_views:number;ai_questions:number;upload_starts:number;upload_completions:number;frontend_errors:number} }>(`/usage/funnel?days=${days}`)).data.data;
  },
  track(event_name: string, properties: UsageProperties = {}) {
    if (typeof window === "undefined") return;
    const payload = {
      event_name,
      path: window.location.pathname,
      session_id: sessionId(),
      properties,
    };
    // Product telemetry must never block the product experience.
    void api.post("/usage/events", payload, { timeout: 5000 }).catch(() => undefined);
  },
};
