"use client";

import { useCallback, useEffect, useState } from "react";
import { Building2, ChevronDown, Layers3 } from "lucide-react";
import { financeService } from "@/services/finance-service";
import type { Branch } from "@/types/finance";
import { readWorkspaceScope, saveWorkspaceScope, WORKSPACE_SCOPE_EVENT, type WorkspaceScope } from "@/lib/workspace-scope";
import { BRANCHES_CHANGED_EVENT } from "@/lib/branch-events";
import { usageService } from "@/services/usage-service";

function selectableBranches(items: Branch[]) {
  return items
    .filter((item) => item.is_active && item.review_status === "accepted")
    .sort((a, b) => a.branch_name.localeCompare(b.branch_name));
}

export function WorkspaceScopeSelector() {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [scope, setScope] = useState<WorkspaceScope>({ mode: "consolidated" });

  const loadBranches = useCallback(async () => {
    try {
      const items = selectableBranches(await financeService.getBranches());
      setBranches(items);
      const current = readWorkspaceScope();
      if (current.mode === "branch") {
        const branch = items.find((item) => item.id === current.branchId);
        if (!branch) {
          const consolidated: WorkspaceScope = { mode: "consolidated" };
          setScope(consolidated);
          saveWorkspaceScope(consolidated);
        } else if (current.branchName !== branch.branch_name) {
          const refreshed: WorkspaceScope = { mode: "branch", branchId: branch.id, branchName: branch.branch_name };
          setScope(refreshed);
          saveWorkspaceScope(refreshed);
        }
      }
    } catch {
      // Keep the existing selector state during a transient API failure.
    }
  }, []);

  useEffect(() => {
    setScope(readWorkspaceScope());
    void loadBranches();

    const onBranchesChanged = () => void loadBranches();
    const onScopeChanged = (event: Event) => {
      const detail = (event as CustomEvent<WorkspaceScope>).detail;
      setScope(detail ?? readWorkspaceScope());
    };
    const onFocus = () => void loadBranches();
    const onVisible = () => { if (document.visibilityState === "visible") void loadBranches(); };

    window.addEventListener(BRANCHES_CHANGED_EVENT, onBranchesChanged);
    window.addEventListener(WORKSPACE_SCOPE_EVENT, onScopeChanged);
    window.addEventListener("focus", onFocus);
    document.addEventListener("visibilitychange", onVisible);
    return () => {
      window.removeEventListener(BRANCHES_CHANGED_EVENT, onBranchesChanged);
      window.removeEventListener(WORKSPACE_SCOPE_EVENT, onScopeChanged);
      window.removeEventListener("focus", onFocus);
      document.removeEventListener("visibilitychange", onVisible);
    };
  }, [loadBranches]);

  function choose(value: string) {
    if (value === "consolidated") {
      const next: WorkspaceScope = { mode: "consolidated" };
      setScope(next); saveWorkspaceScope(next);
      usageService.track("workspace_scope_changed", { scope: "consolidated" });
      return;
    }
    const branch = branches.find((item) => item.id === value);
    if (!branch) return;
    const next: WorkspaceScope = { mode: "branch", branchId: branch.id, branchName: branch.branch_name };
    setScope(next); saveWorkspaceScope(next);
    usageService.track("workspace_scope_changed", { scope: "branch" });
  }

  if (!branches.length) return null;

  return (
    <label className="relative hidden items-center gap-2 rounded-xl border bg-background px-2.5 py-2 text-xs font-semibold text-muted-foreground shadow-sm xl:flex" title="Choose whether FinCruiz should work at consolidated company or branch scope">
      {scope.mode === "branch" ? <Building2 className="size-3.5 text-indigo-500"/> : <Layers3 className="size-3.5 text-indigo-500"/>}
      <span className="max-w-28 truncate">{scope.mode === "branch" ? scope.branchName : "Consolidated"}</span>
      <ChevronDown className="size-3 opacity-60"/>
      <select value={scope.mode === "branch" ? scope.branchId : "consolidated"} onChange={(event) => choose(event.target.value)} className="absolute inset-0 cursor-pointer opacity-0" aria-label="Workspace analysis scope">
        <option value="consolidated">Consolidated company</option>
        {branches.map((branch) => <option key={branch.id} value={branch.id}>{branch.branch_name}</option>)}
      </select>
    </label>
  );
}
