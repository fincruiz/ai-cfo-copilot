export const BRANCHES_CHANGED_EVENT = "fincruiz:branches-changed";

export function notifyBranchesChanged() {
  if (typeof window === "undefined") return;
  window.dispatchEvent(new CustomEvent(BRANCHES_CHANGED_EVENT));
}
