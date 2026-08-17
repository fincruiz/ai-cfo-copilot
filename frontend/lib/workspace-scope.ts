export type WorkspaceScope = {
  mode: "consolidated" | "branch";
  branchId?: string;
  branchName?: string;
};

export const WORKSPACE_SCOPE_STORAGE_KEY = "fincruiz_workspace_scope";
export const WORKSPACE_SCOPE_EVENT = "fincruiz:workspace-scope-changed";

export function readWorkspaceScope(): WorkspaceScope {
  if (typeof window === "undefined") return { mode: "consolidated" };
  try {
    const raw = window.localStorage.getItem(WORKSPACE_SCOPE_STORAGE_KEY);
    if (!raw) return { mode: "consolidated" };
    const parsed = JSON.parse(raw) as WorkspaceScope;
    if (parsed.mode === "branch" && parsed.branchId) return parsed;
  } catch {
    // Invalid old preference: fall back safely.
  }
  return { mode: "consolidated" };
}

export function saveWorkspaceScope(scope: WorkspaceScope) {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(WORKSPACE_SCOPE_STORAGE_KEY, JSON.stringify(scope));
  window.dispatchEvent(new CustomEvent(WORKSPACE_SCOPE_EVENT, { detail: scope }));
}
