"use client";

import { useCallback, useEffect, useRef, useState } from "react";
import { Clock3, LogOut, ShieldCheck } from "lucide-react";

import { Button } from "@/components/ui/button";
import { ACCESS_TOKEN_KEY } from "@/lib/api";
import {
  beginSession,
  markSessionActivity,
  parseSessionSignal,
  readSessionTimestamp,
  SESSION_EVENT_KEY,
  SESSION_IDLE_TIMEOUT_MS,
  SESSION_LAST_ACTIVITY_KEY,
  SESSION_MAX_LIFETIME_MS,
  SESSION_STARTED_AT_KEY,
  SESSION_WARNING_MS,
  type SessionLogoutReason,
} from "@/lib/session-security";
import { authService } from "@/services/auth-service";

const ACTIVITY_WRITE_THROTTLE_MS = 10_000;
const SESSION_CHECK_INTERVAL_MS = 10_000;

function loginReason(reason: SessionLogoutReason): string {
  if (reason === "inactivity") return "inactivity";
  if (reason === "session-limit") return "session-limit";
  if (reason === "session-expired") return "session-expired";
  return "signed-out";
}

export function SessionSecurityGuard() {
  const [warning, setWarning] = useState(false);
  const lastActivityWrite = useRef(0);
  const leaving = useRef(false);

  const leaveSession = useCallback((reason: SessionLogoutReason) => {
    if (leaving.current) return;
    leaving.current = true;

    if (reason === "signed-out") {
      void authService.logoutEverywhere(reason);
    } else {
      void authService.logoutCurrentSession(reason);
    }

    window.location.replace(`/login?reason=${loginReason(reason)}`);
  }, []);

  const evaluateSession = useCallback(() => {
    if (!window.localStorage.getItem(ACCESS_TOKEN_KEY)) return;

    const now = Date.now();
    let startedAt = readSessionTimestamp(SESSION_STARTED_AT_KEY);
    let lastActivity = readSessionTimestamp(SESSION_LAST_ACTIVITY_KEY);

    if (!startedAt || !lastActivity) {
      beginSession(now);
      startedAt = now;
      lastActivity = now;
    }

    const absoluteAge = now - startedAt;
    const idleAge = now - lastActivity;

    if (absoluteAge >= SESSION_MAX_LIFETIME_MS) {
      leaveSession("session-limit");
      return;
    }

    if (idleAge >= SESSION_IDLE_TIMEOUT_MS) {
      leaveSession("inactivity");
      return;
    }

    setWarning(idleAge >= SESSION_IDLE_TIMEOUT_MS - SESSION_WARNING_MS);
  }, [leaveSession]);

  useEffect(() => {
    evaluateSession();

    const recordActivity = () => {
      if (!window.localStorage.getItem(ACCESS_TOKEN_KEY)) return;
      const now = Date.now();
      if (now - lastActivityWrite.current < ACTIVITY_WRITE_THROTTLE_MS) return;
      lastActivityWrite.current = now;
      markSessionActivity(now);
      setWarning(false);
    };

    const onStorage = (event: StorageEvent) => {
      if (event.key === ACCESS_TOKEN_KEY && event.newValue === null) {
        leaveSession("signed-out");
        return;
      }

      if (event.key === SESSION_EVENT_KEY) {
        const signal = parseSessionSignal(event.newValue);
        if (signal?.type === "logout") leaveSession(signal.reason);
      }
    };

    const onVisibility = () => {
      if (document.visibilityState === "visible") evaluateSession();
    };

    const activityEvents: Array<keyof WindowEventMap> = [
      "pointerdown",
      "keydown",
      "touchstart",
      "focus",
    ];

    activityEvents.forEach((eventName) =>
      window.addEventListener(eventName, recordActivity, { passive: true }),
    );
    window.addEventListener("storage", onStorage);
    document.addEventListener("visibilitychange", onVisibility);

    const interval = window.setInterval(evaluateSession, SESSION_CHECK_INTERVAL_MS);

    return () => {
      activityEvents.forEach((eventName) =>
        window.removeEventListener(eventName, recordActivity),
      );
      window.removeEventListener("storage", onStorage);
      document.removeEventListener("visibilitychange", onVisibility);
      window.clearInterval(interval);
    };
  }, [evaluateSession, leaveSession]);

  if (!warning) return null;

  const remainingMinutes = Math.max(1, Math.ceil(SESSION_WARNING_MS / 60_000));

  return (
    <div className="fixed inset-0 z-[140] flex items-center justify-center bg-slate-950/45 px-4 backdrop-blur-sm">
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="session-warning-title"
        className="w-full max-w-md rounded-3xl border bg-background p-6 shadow-2xl"
      >
        <div className="flex size-11 items-center justify-center rounded-2xl bg-amber-500/10 text-amber-700 dark:text-amber-300">
          <Clock3 className="size-5" />
        </div>
        <h2 id="session-warning-title" className="mt-4 text-xl font-bold tracking-tight">
          Your session will expire soon
        </h2>
        <p className="mt-2 text-sm leading-6 text-muted-foreground">
          For the security of your financial data, FinCruiz signs you out after
          inactivity. Your session will close in about {remainingMinutes} minute{remainingMinutes === 1 ? "" : "s"} unless you continue.
        </p>
        <div className="mt-6 grid gap-2 sm:grid-cols-2">
          <Button
            type="button"
            onClick={() => {
              markSessionActivity();
              setWarning(false);
            }}
          >
            <ShieldCheck className="size-4" />
            Stay signed in
          </Button>
          <Button type="button" variant="outline" onClick={() => leaveSession("signed-out")}>
            <LogOut className="size-4" />
            Sign out
          </Button>
        </div>
      </div>
    </div>
  );
}
