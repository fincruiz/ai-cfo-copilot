import { api } from "@/lib/api";

const key = "fincruiz_marketing_session";

function sessionId() {
  if (typeof window === "undefined") return "server";
  let value = window.sessionStorage.getItem(key);
  if (!value) {
    value = globalThis.crypto?.randomUUID?.() ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    window.sessionStorage.setItem(key, value);
  }
  return value;
}

export type MarketingFunnel = {
  visitors: number;
  hero_demo: number;
  hero_signup: number;
  ai_questions: number;
  ai_signup: number;
  pricing: number;
  final_signup: number;
  demo_views?: number;
  demo_questions?: number;
  demo_signup?: number;
};

export const marketingService = {
  track(event_name: string, properties: Record<string, string | number | boolean> = {}) {
    if (typeof window === "undefined") return;
    void api
      .post(
        "/marketing/events",
        {
          event_name,
          session_id: sessionId(),
          path: window.location.pathname,
          referrer: document.referrer,
          properties,
        },
        { timeout: 4000 },
      )
      .catch(() => undefined);
  },

  async funnel(days = 30) {
    return (
      await api.get<{ data: MarketingFunnel }>(`/marketing/funnel?days=${days}`)
    ).data.data;
  },
};
