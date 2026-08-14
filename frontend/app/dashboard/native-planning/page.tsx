"use client";

import { useEffect, useState } from "react";
import { AlertTriangle, Loader2, Plus, RefreshCw, Save } from "lucide-react";
import { ModuleResetButton } from "@/components/module-reset-button";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { advancedFinanceService } from "@/services/advanced-finance-service";
import { getApiErrorMessage } from "@/lib/api";
import type { PlanningVersion } from "@/types/advanced-forecasting";

const groups = ["Revenue", "Cost of Sales", "Payroll", "Operating Expenses", "Depreciation", "Finance Costs", "Tax"];

export default function NativePlanning() {
  const [versions, setVersions] = useState<PlanningVersion[]>([]);
  const [selected, setSelected] = useState<PlanningVersion | null>(null);
  const [loading, setLoading] = useState(false);
  const [initialLoading, setInitialLoading] = useState(true);
  const [error, setError] = useState("");
  const [name, setName] = useState("FY27 Budget");

  async function load() {
    setInitialLoading(true); setError("");
    try { setVersions(await advancedFinanceService.versions()); }
    catch (e) { setVersions([]); setError(getApiErrorMessage(e)); }
    finally { setInitialLoading(false); }
  }
  useEffect(() => { void load(); }, []);

  async function create() {
    setLoading(true); setError("");
    try {
      const v = await advancedFinanceService.createVersion({ plan_type: "budget", version_name: name, financial_year_start: "2027-01-01", financial_year_end: "2027-12-31", seed_from_actuals: true, seed_growth_percent: 5 });
      setSelected(v); await load();
    } catch (e) { setError(getApiErrorMessage(e)); } finally { setLoading(false); }
  }
  async function open(id: string) { setError(""); try { setSelected(await advancedFinanceService.getVersion(id)); } catch (e) { setError(getApiErrorMessage(e)); } }
  function patchLine(i: number, patch: Record<string, unknown>) { if (!selected) return; const lines = [...(selected.lines ?? [])]; lines[i] = { ...lines[i], ...patch }; setSelected({ ...selected, lines }); }
  function add() { if (!selected) return; setSelected({ ...selected, lines: [...(selected.lines ?? []), { period: selected.financial_year_start, reporting_group: "Revenue", amount: 0, driver_type: "manual", branch_id: null, reporting_subgroup: null, source_account_code: null, driver_value: null, notes: null }] }); }
  async function save() { if (!selected) return; setLoading(true); setError(""); try { setSelected(await advancedFinanceService.saveLines(selected.id, selected.lines ?? [])); } catch (e) { setError(getApiErrorMessage(e)); } finally { setLoading(false); } }

  return <div className="mx-auto max-w-7xl space-y-6">
    <div className="flex flex-col gap-4 sm:flex-row sm:items-start sm:justify-between">
      <div><p className="text-sm text-muted-foreground">Planning & intelligence</p><h1 className="mt-1 text-3xl font-semibold">Native Budget Builder</h1><p className="mt-2 max-w-3xl text-muted-foreground">Create, seed, edit and save monthly budgets. Backend failures are now handled inside the page instead of becoming a full-screen Next.js runtime error.</p></div>
      <ModuleResetButton scope="planning" label="Reset planning data" description="Remove saved budgets and planning lines only. Actual finance data remains." onReset={() => { setSelected(null); void load(); }} />
    </div>
    {error ? <Alert variant="destructive"><AlertTriangle className="size-4" /><AlertDescription><div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between"><span>{error}</span><Button size="sm" variant="outline" onClick={() => void load()}><RefreshCw className="size-4" />Retry</Button></div></AlertDescription></Alert> : null}
    <div className="grid gap-5 lg:grid-cols-[320px_1fr]">
      <Card><CardHeader><CardTitle>Versions</CardTitle><CardDescription>Start from actuals, then refine assumptions.</CardDescription></CardHeader><CardContent className="space-y-3"><Input value={name} onChange={e => setName(e.target.value)} /><Button className="w-full" onClick={() => void create()} disabled={loading || !name.trim()}>{loading ? <Loader2 className="size-4 animate-spin" /> : <Plus className="size-4" />}New seeded budget</Button>{initialLoading ? <div className="flex min-h-24 items-center justify-center"><Loader2 className="size-5 animate-spin" /></div> : versions.length === 0 ? <div className="rounded-xl border border-dashed p-4 text-sm text-muted-foreground">No planning versions yet.</div> : versions.map(v => <button key={v.id} onClick={() => void open(v.id)} className="w-full rounded-xl border p-3 text-left transition hover:bg-muted"><b>{v.version_name}</b><p className="text-xs text-muted-foreground">{v.status} · {v.plan_type}</p></button>)}</CardContent></Card>
      <Card><CardHeader className="flex-row items-center justify-between"><div><CardTitle>{selected?.version_name ?? "Planning workspace"}</CardTitle><CardDescription>{selected ? "Edit monthly values and save the version." : "Select a version to begin."}</CardDescription></div>{selected ? <div className="flex gap-2"><Button variant="outline" onClick={add}><Plus className="size-4" />Line</Button><Button onClick={() => void save()} disabled={loading}>{loading ? <Loader2 className="size-4 animate-spin" /> : <Save className="size-4" />}Save</Button></div> : null}</CardHeader><CardContent>{!selected ? <div className="flex min-h-72 items-center justify-center rounded-2xl border border-dashed text-muted-foreground">Your selected budget will appear here.</div> : <div className="overflow-x-auto"><table className="w-full min-w-[850px] text-sm"><thead><tr className="border-b">{["Period", "Reporting group", "Subgroup", "Amount", "Notes"].map(h => <th className="p-3 text-left" key={h}>{h}</th>)}</tr></thead><tbody>{(selected.lines ?? []).map((l: any, i) => <tr key={`${l.period}-${l.reporting_group}-${i}`} className="border-b"><td className="p-2"><Input type="date" value={String(l.period).slice(0, 10)} onChange={e => patchLine(i, { period: e.target.value })} /></td><td className="p-2"><select className="h-10 rounded-md border bg-background px-2" value={l.reporting_group} onChange={e => patchLine(i, { reporting_group: e.target.value })}>{groups.map(g => <option key={g}>{g}</option>)}</select></td><td className="p-2"><Input value={l.reporting_subgroup ?? ""} onChange={e => patchLine(i, { reporting_subgroup: e.target.value || null })} /></td><td className="p-2"><Input type="number" value={l.amount} onChange={e => patchLine(i, { amount: Number(e.target.value) })} /></td><td className="p-2"><Input value={l.notes ?? ""} onChange={e => patchLine(i, { notes: e.target.value || null })} /></td></tr>)}</tbody></table></div>}</CardContent></Card>
    </div>
  </div>;
}
