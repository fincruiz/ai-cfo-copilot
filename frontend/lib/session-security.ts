export const SESSION_STARTED_AT_KEY = "fincruiz_session_started_at";
export const SESSION_LAST_ACTIVITY_KEY = "fincruiz_session_last_activity";
export const SESSION_EVENT_KEY = "fincruiz_session_event";

export type SessionLogoutReason =
  | "signed-out"
  | "inactivity"
  | "session-limit"
  | "session-expired";

export type SessionSignal = {
  type: "logout";
  reason: SessionLogoutReason;
  at: number;
};

function positiveNumber(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

export const SESSION_IDLE_TIMEOUT_MS =
  positiveNumber(process.env.NEXT_PUBLIC_SESSION_IDLE_MINUTES, 30) * 60_000;

export const SESSION_MAX_LIFETIME_MS =
  positiveNumber(process.env.NEXT_PUBLIC_SESSION_MAX_HOURS, 12) * 60 * 60_000;

export const SESSION_WARNING_MS = Math.min(
  positiveNumber(process.env.NEXT_PUBLIC_SESSION_WARNING_MINUTES, 2) * 60_000,
  Math.max(30_000, SESSION_IDLE_TIMEOUT_MS / 2),
);

function setTimestamp(key: string, value: number): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(key, String(value));
}

export function beginSession(now = Date.now()): void {
  if (typeof window === "undefined") return;
  setTimestamp(SESSION_STARTED_AT_KEY, now);
  setTimestamp(SESSION_LAST_ACTIVITY_KEY, now);
}

export function markSessionActivity(now = Date.now()): void {
  setTimestamp(SESSION_LAST_ACTIVITY_KEY, now);
}

export function readSessionTimestamp(key: string): number | null {
  if (typeof window === "undefined") return null;
  const value = Number(window.localStorage.getItem(key));
  return Number.isFinite(value) && value > 0 ? value : null;
}

export function clearSessionMetadata(): void {
  if (typeof window === "undefined") return;
  window.localStorage.removeItem(SESSION_STARTED_AT_KEY);
  window.localStorage.removeItem(SESSION_LAST_ACTIVITY_KEY);
}

export function broadcastSessionLogout(reason: SessionLogoutReason): void {
  if (typeof window === "undefined") return;
  const signal: SessionSignal = { type: "logout", reason, at: Date.now() };
  window.localStorage.setItem(SESSION_EVENT_KEY, JSON.stringify(signal));
}

export function parseSessionSignal(value: string | null): SessionSignal | null {
  if (!value) return null;
  try {
    const parsed = JSON.parse(value) as Partial<SessionSignal>;
    if (
      parsed.type === "logout" &&
      typeof parsed.reason === "string" &&
      typeof parsed.at === "number"
    ) {
      return parsed as SessionSignal;
    }
  } catch {
    return null;
  }
  return null;
}
