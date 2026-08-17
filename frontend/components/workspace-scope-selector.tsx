"use client";

import { useEffect, useState } from "react";
import { Building2, ChevronDown, Layers3 } from "lucide-react";
import { financeService } from "@/services/finance-service";
import type { Branch } from "@/types/finance";
import { readWorkspaceScope, saveWorkspaceScope, type WorkspaceScope } from "@/lib/workspace-scope";
import { usageService } from "@/services/usage-service";

export function WorkspaceScopeSelector() {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [scope, setScope] = useState<WorkspaceScope>({ mode: "consolidated" });

  useEffect(() => {
    setScope(readWorkspaceScope());
    financeService.getBranches().then((items) => setBranches(items.filter((item) => item.is_active))).catch(() => setBranches([]));
  }, []);

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
