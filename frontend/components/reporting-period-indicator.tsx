"use client";

import { useCallback, useEffect, useState } from "react";
import { CalendarRange } from "lucide-react";
import { financeService } from "@/services/finance-service";
import { readWorkspaceScope, WORKSPACE_SCOPE_EVENT, type WorkspaceScope } from "@/lib/workspace-scope";
import type { ReportContext } from "@/types/finance";

function formatDate(value?: string | null) {
  if (!value) return "—";
  const date = new Date(`${value.slice(0, 10)}T00:00:00`);
  return date.toLocaleDateString(undefined, { day: "numeric", month: "short", year: "numeric" });
}

export function ReportingPeriodIndicator() {
  const [context, setContext] = useState<ReportContext | null>(null);
  const [scope, setScope] = useState<WorkspaceScope>({ mode: "consolidated" });

  const load = useCallback(async (nextScope?: WorkspaceScope) => {
    const activeScope = nextScope ?? readWorkspaceScope();
    setScope(activeScope);
    try {
      setContext(await financeService.getReportContext({
        branchId: activeScope.mode === "branch" ? activeScope.branchId : undefined,
      }));
    } catch {
      setContext(null);
    }
  }, []);

  useEffect(() => {
    void load();
    const onScope = (event: Event) => {
      const next = (event as CustomEvent<WorkspaceScope>).detail ?? readWorkspaceScope();
      void load(next);
    };
    window.addEventListener(WORKSPACE_SCOPE_EVENT, onScope);
    return () => window.removeEventListener(WORKSPACE_SCOPE_EVENT, onScope);
  }, [load]);

  const label = context?.period_start && context?.period_end
    ? `${formatDate(context.period_start)} – ${formatDate(context.period_end)}`
    : "No active reporting period";
  const scopeLabel = scope.mode === "branch" ? scope.branchName : "Consolidated";

  return (
    <div
      className="hidden max-w-[280px] items-center gap-2 rounded-xl border bg-background px-3 py-2 text-xs shadow-sm lg:flex"
      title={`Reporting period: ${label}. Scope: ${scopeLabel}. Data as of: ${formatDate(context?.data_as_of)}. ${context?.transaction_count ?? 0} active ledger transactions.`}
    >
      <CalendarRange className="size-3.5 shrink-0 text-indigo-500" />
      <div className="min-w-0">
        <p className="truncate font-semibold text-foreground">{label}</p>
        <p className="truncate text-[10px] text-muted-foreground">Data as of {formatDate(context?.data_as_of)} · {scopeLabel}</p>
      </div>
    </div>
  );
}
