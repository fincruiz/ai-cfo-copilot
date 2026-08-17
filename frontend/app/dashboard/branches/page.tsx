"use client";

import { FormEvent, useEffect, useMemo, useState } from "react";
import { Building2, Check, Loader2, Plus, RefreshCw, Save } from "lucide-react";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { getApiErrorMessage } from "@/lib/api";
import { notifyBranchesChanged } from "@/lib/branch-events";
import { ModuleResetButton } from "@/components/module-reset-button";
import { financeService } from "@/services/finance-service";
import type { Branch } from "@/types/finance";

export default function BranchesPage() {
  const [branches, setBranches] = useState<Branch[]>([]);
  const [drafts, setDrafts] = useState<Record<string, Branch>>({});
  const [code, setCode] = useState("");
  const [name, setName] = useState("");
  const [region, setRegion] = useState("");
  const [isLoading, setIsLoading] = useState(true);
  const [savingId, setSavingId] = useState("");
  const [error, setError] = useState("");

  const pending = useMemo(
    () => branches.filter((branch) => branch.review_status === "pending"),
    [branches],
  );

  async function load() {
    setIsLoading(true);
    setError("");
    try {
      const rows = await financeService.getBranches();
      setBranches(rows);
      setDrafts(Object.fromEntries(rows.map((row) => [row.id, { ...row }])));
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setIsLoading(false);
    }
  }

  useEffect(() => { void load(); }, []);

  async function create(event: FormEvent) {
    event.preventDefault();
    setSavingId("new");
    setError("");
    try {
      await financeService.createBranch({
        branch_code: code.trim().toUpperCase(),
        branch_name: name.trim(),
        region: region.trim() || null,
      });
      setCode(""); setName(""); setRegion("");
      await load();
      notifyBranchesChanged();
    } catch (saveError) {
      setError(getApiErrorMessage(saveError));
    } finally {
      setSavingId("");
    }
  }

  function updateDraft(id: string, patch: Partial<Branch>) {
    setDrafts((current) => ({ ...current, [id]: { ...current[id], ...patch } }));
  }

  async function saveBranch(id: string, accept = false) {
    const draft = drafts[id];
    if (!draft) return;
    setSavingId(id);
    setError("");
    try {
      await financeService.updateBranch(id, {
        branch_code: draft.branch_code.trim().toUpperCase(),
        branch_name: draft.branch_name.trim(),
        region: draft.region?.trim() || null,
        review_status: accept ? "accepted" : draft.review_status,
        is_active: draft.is_active,
      });
      await load();
      notifyBranchesChanged();
    } catch (saveError) {
      setError(getApiErrorMessage(saveError));
    } finally {
      setSavingId("");
    }
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
        <div>
          <p className="text-sm font-medium text-muted-foreground">Company structure</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Branches and business units</h1>
          <p className="mt-2 max-w-3xl text-muted-foreground">
            FinCruiz discovers unique branch values during upload. Review, rename and accept them here. Resetting branches keeps your ledger and clears branch links only.
          </p>
        </div>
        <ModuleResetButton scope="branches" label="Reset branches" description="Remove all saved branch records for this company. Your General Ledger and other financial data will remain; branch links on existing records will be cleared." onReset={() => void load()} />
      </div>

      {error ? <Alert variant="destructive"><AlertDescription>{error}</AlertDescription></Alert> : null}

      {pending.length ? (
        <Alert>
          <AlertDescription>
            {pending.length} discovered branch value{pending.length === 1 ? "" : "s"} require review. Transactions are already stored against these pending branches.
          </AlertDescription>
        </Alert>
      ) : null}

      <Card>
        <CardHeader>
          <CardTitle className="flex items-center gap-2"><Plus className="size-5" />Add branch manually</CardTitle>
          <CardDescription>Manual creation remains available, but it is no longer required before upload.</CardDescription>
        </CardHeader>
        <CardContent>
          <form className="grid gap-3 md:grid-cols-4" onSubmit={(event) => void create(event)}>
            <Input placeholder="Code, e.g. MEL" value={code} onChange={(event) => setCode(event.target.value)} required />
            <Input placeholder="Branch name" value={name} onChange={(event) => setName(event.target.value)} required />
            <Input placeholder="Region (optional)" value={region} onChange={(event) => setRegion(event.target.value)} />
            <Button type="submit" disabled={savingId === "new"}>
              {savingId === "new" ? <Loader2 className="size-4 animate-spin" /> : <Plus className="size-4" />}
              Create branch
            </Button>
          </form>
        </CardContent>
      </Card>

      <Card>
        <CardHeader className="flex-row items-center justify-between">
          <div>
            <CardTitle className="flex items-center gap-2"><Building2 className="size-5" />Branch review</CardTitle>
            <CardDescription>{branches.length} stored branch records · {pending.length} pending acceptance</CardDescription>
          </div>
          <Button variant="outline" onClick={() => void load()} disabled={isLoading}>
            <RefreshCw className="size-4" />Refresh
          </Button>
        </CardHeader>
        <CardContent>
          {isLoading ? (
            <div className="flex min-h-48 items-center justify-center"><Loader2 className="size-5 animate-spin" /></div>
          ) : branches.length === 0 ? (
            <div className="flex min-h-48 items-center justify-center text-center text-muted-foreground">
              Upload a branch-tagged ledger and unique values will appear here automatically.
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full min-w-[980px] text-sm">
                <thead>
                  <tr className="border-b text-left text-muted-foreground">
                    <th className="p-3">Detected value</th>
                    <th className="p-3">Branch code</th>
                    <th className="p-3">Branch name</th>
                    <th className="p-3">Region</th>
                    <th className="p-3">Status</th>
                    <th className="p-3 text-right">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {branches.map((branch) => {
                    const draft = drafts[branch.id] ?? branch;
                    const isPending = branch.review_status === "pending";
                    return (
                      <tr key={branch.id} className={isPending ? "border-b bg-amber-50/50 dark:bg-amber-950/10" : "border-b"}>
                        <td className="p-3 text-muted-foreground">{branch.source_value || "Manual"}</td>
                        <td className="p-3"><Input value={draft.branch_code} onChange={(e) => updateDraft(branch.id, { branch_code: e.target.value })} /></td>
                        <td className="p-3"><Input value={draft.branch_name} onChange={(e) => updateDraft(branch.id, { branch_name: e.target.value })} /></td>
                        <td className="p-3"><Input value={draft.region || ""} onChange={(e) => updateDraft(branch.id, { region: e.target.value })} /></td>
                        <td className="p-3">
                          <span className={`rounded-full px-2 py-1 text-xs font-medium ${isPending ? "bg-amber-100 text-amber-800" : "bg-emerald-100 text-emerald-800"}`}>
                            {isPending ? "Pending review" : "Accepted"}
                          </span>
                        </td>
                        <td className="p-3">
                          <div className="flex justify-end gap-2">
                            <Button size="sm" variant="outline" onClick={() => void saveBranch(branch.id)} disabled={savingId === branch.id}>
                              <Save className="size-4" />Save edits
                            </Button>
                            {isPending ? (
                              <Button size="sm" onClick={() => void saveBranch(branch.id, true)} disabled={savingId === branch.id}>
                                {savingId === branch.id ? <Loader2 className="size-4 animate-spin" /> : <Check className="size-4" />}
                                Accept
                              </Button>
                            ) : null}
                          </div>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </CardContent>
      </Card>
    </div>
  );
}
