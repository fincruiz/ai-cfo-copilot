"use client";

import Link from "next/link";
import { useEffect, useMemo, useState } from "react";
import { AlertTriangle, BarChart3, ChevronRight, History, Loader2, Plus, RefreshCw, RotateCcw, Save, Sparkles, TrendingUp, WandSparkles } from "lucide-react";
import { ModuleResetButton } from "@/components/module-reset-button";
import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { advancedFinanceService } from "@/services/advanced-finance-service";
import { getApiErrorMessage } from "@/lib/api";
import { formatMoney, toNumber } from "@/lib/finance-format";
import { readWorkspaceScope } from "@/lib/workspace-scope";
import type { PlanningContext, PlanningVersion } from "@/types/advanced-forecasting";

const groups = ["Revenue", "Cost of Sales", "Payroll", "Operating Expenses", "Depreciation", "Finance Costs", "Tax", "Other Income", "Other Expenses"];
const snapshot = (v: PlanningVersion) => JSON.parse(JSON.stringify(v)) as PlanningVersion;

export default function NativePlanning() {
  const [versions, setVersions] = useState<PlanningVersion[]>([]);
  const [context, setContext] = useState<PlanningContext | null>(null);
  const [selected, setSelected] = useState<PlanningVersion | null>(null);
  const [history, setHistory] = useState<PlanningVersion[]>([]);
  const [loading, setLoading] = useState(false);
  const [initialLoading, setInitialLoading] = useState(true);
  const [error, setError] = useState("");
  const [name, setName] = useState("FY27 Budget");
  const [start, setStart] = useState("2027-01-01");
  const [end, setEnd] = useState("2027-12-31");
  const [seedMode, setSeedMode] = useState("actuals");
  const [detailLevel, setDetailLevel] = useState("high_level");
  const [growth, setGrowth] = useState(5);
  const [sourceVersion, setSourceVersion] = useState("");
  const [allocationDetail, setAllocationDetail] = useState("detailed");
  const [seasonality, setSeasonality] = useState("historical");
  const [targets, setTargets] = useState<Record<string, number>>({});
  const [targetMode, setTargetMode] = useState<"executive"|"detailed">("executive");
  const [revenueTarget, setRevenueTarget] = useState(0);
  const [grossMarginTarget, setGrossMarginTarget] = useState(40);
  const [netProfitTarget, setNetProfitTarget] = useState(0);
  const [allocationMethod, setAllocationMethod] = useState("historical_actuals");

  async function load() {
    setInitialLoading(true); setError("");
    try {
      const [v, c] = await Promise.all([advancedFinanceService.versions(), advancedFinanceService.planningContext()]);
      setVersions(v); setContext(c);
      if (c.recommended_seed) setSeedMode(c.recommended_seed);
    } catch (e) { setVersions([]); setError(getApiErrorMessage(e)); }
    finally { setInitialLoading(false); }
  }
  useEffect(() => { void load(); }, []);

  const annualTotals = useMemo(() => {
    const out: Record<string, number> = {};
    for (const line of selected?.lines ?? []) out[line.reporting_group] = (out[line.reporting_group] ?? 0) + toNumber(line.amount);
    return out;
  }, [selected]);

  useEffect(() => {
    if (selected) setTargets(Object.fromEntries(groups.map(g => [g, Number((annualTotals[g] ?? 0).toFixed(2))])));
  }, [selected?.id]); // eslint-disable-line react-hooks/exhaustive-deps

  async function create() {
    setLoading(true); setError("");
    try {
      const native = sourceVersion.startsWith("native:") ? sourceVersion.slice(7) : null;
      const imported = sourceVersion.startsWith("imported:") ? sourceVersion.slice(9) : null;
      const v = await advancedFinanceService.createVersion({
        plan_type: "budget", version_name: name, financial_year_start: start, financial_year_end: end,
        seed_mode: seedMode, detail_level: detailLevel, allocation_method: "actuals_ratio",
        seed_growth_percent: growth, seed_version_id: native, seed_imported_version: imported,
      });
      setSelected(v); setHistory([]); await load();
    } catch (e) { setError(getApiErrorMessage(e)); } finally { setLoading(false); }
  }

  async function open(id: string) { setError(""); try { const v = await advancedFinanceService.getVersion(id); setSelected(v); setHistory([]); } catch (e) { setError(getApiErrorMessage(e)); } }
  function remember() { if (selected) setHistory(h => [...h.slice(-19), snapshot(selected)]); }
  function patchLine(i: number, patch: Record<string, unknown>) { if (!selected) return; remember(); const lines = [...(selected.lines ?? [])]; lines[i] = { ...lines[i], ...patch }; setSelected({ ...selected, lines }); }
  function add() { if (!selected) return; remember(); setSelected({ ...selected, lines: [...(selected.lines ?? []), { period: selected.financial_year_start, reporting_group: "Revenue", amount: 0, driver_type: "manual", branch_id: null, reporting_subgroup: null, source_account_code: null, driver_value: null, notes: null }] }); }
  function undo() { const prev = history.at(-1); if (!prev) return; setSelected(prev); setHistory(h => h.slice(0, -1)); }
  async function save() { if (!selected) return; setLoading(true); setError(""); try { setSelected(await advancedFinanceService.saveLines(selected.id, selected.lines ?? [])); setHistory([]); } catch (e) { setError(getApiErrorMessage(e)); } finally { setLoading(false); } }
  async function reseed() { if (!selected || !confirm("Remove all manual edits and restore this budget to its original seed?")) return; setLoading(true); try { setSelected(await advancedFinanceService.reseedVersion(selected.id)); setHistory([]); } catch(e){setError(getApiErrorMessage(e))} finally{setLoading(false)} }
  async function allocate() { if (!selected) return; setLoading(true); setError(""); try {
    const scope = readWorkspaceScope();
    const annual_targets = targetMode === "detailed" ? Object.fromEntries(Object.entries(targets).filter(([,v]) => Number.isFinite(v))) : {};
    setSelected(await advancedFinanceService.allocateBudget(selected.id,{
      annual_targets,
      revenue_target: targetMode === "executive" ? revenueTarget : null,
      gross_margin_percent: targetMode === "executive" ? grossMarginTarget : null,
      net_profit_target: targetMode === "executive" ? netProfitTarget : null,
      branch_id: scope.mode === "branch" ? scope.branchId : null,
      detail_level: allocationDetail, seasonality, allocation_method: allocationMethod,
    })); setHistory([]);
  } catch(e){setError(getApiErrorMessage(e))} finally{setLoading(false)} }

  return <div className="mx-auto max-w-[1500px] space-y-6 pb-12">
    <div className="flex flex-col gap-4 xl:flex-row xl:items-start xl:justify-between">
      <div><p className="text-sm text-muted-foreground">Planning & intelligence</p><h1 className="mt-1 text-3xl font-semibold">Assisted Budget Builder</h1><p className="mt-2 max-w-3xl text-muted-foreground">Start with FinCruiz&apos;s understanding of your actuals, COA or a previous budget. Set management targets at a high level and let the system allocate them down when you want detail.</p></div>
      <div className="flex flex-wrap gap-2"><Link href="/dashboard/forecasting"><Button variant="outline"><TrendingUp className="size-4"/>Forecast</Button></Link><Link href="/dashboard/bi"><Button variant="outline"><BarChart3 className="size-4"/>Visual BI</Button></Link><ModuleResetButton scope="planning" label="Reset planning data" description="Remove saved budgets and planning lines only. Actual finance data remains." onReset={() => { setSelected(null); void load(); }} /></div>
    </div>

    {error ? <Alert variant="destructive"><AlertTriangle className="size-4"/><AlertDescription>{error}</AlertDescription></Alert> : null}

    <Card className="overflow-hidden border-indigo-200 bg-gradient-to-r from-indigo-50/80 via-background to-emerald-50/50 dark:border-indigo-900 dark:from-indigo-950/20 dark:to-emerald-950/10">
      <CardContent className="grid gap-4 p-5 md:grid-cols-4">
        <ContextStat label="Mapped accounts" value={context?.mapped_accounts ?? 0} note="Available for detailed allocation" />
        <ContextStat label="Actual history" value={`${context?.actual_months ?? 0} months`} note={context?.latest_actual_month ? `Through ${String(context.latest_actual_month).slice(0,7)}` : "No mapped actuals yet"} />
        <ContextStat label="Saved plans" value={(context?.native_versions?.length ?? 0)+(context?.imported_versions?.length ?? 0)} note="Can be reused as a starting point" />
        <div className="rounded-2xl border bg-background/80 p-4"><p className="text-sm font-semibold">How it works</p><p className="mt-2 text-xs leading-5 text-muted-foreground">Actuals → management target → monthly seasonality → optional GL allocation → forecast / scenario.</p></div>
      </CardContent>
    </Card>

    <div className="grid gap-5 xl:grid-cols-[350px_1fr]">
      <div className="space-y-5">
        <Card><CardHeader><CardTitle className="flex items-center gap-2"><WandSparkles className="size-5"/>Create a budget</CardTitle><CardDescription>You should rarely need to start from zero.</CardDescription></CardHeader><CardContent className="space-y-3">
          <label className="space-y-1 text-xs"><span>Name</span><Input value={name} onChange={e=>setName(e.target.value)}/></label>
          <div className="grid grid-cols-2 gap-2"><label className="space-y-1 text-xs"><span>Start</span><Input type="date" value={start} onChange={e=>setStart(e.target.value)}/></label><label className="space-y-1 text-xs"><span>End</span><Input type="date" value={end} onChange={e=>setEnd(e.target.value)}/></label></div>
          <label className="space-y-1 text-xs"><span>Start from</span><select className="h-10 w-full rounded-md border bg-background px-3" value={seedMode} onChange={e=>setSeedMode(e.target.value)}><option value="actuals">Mapped actuals (recommended)</option><option value="previous_budget">Previous budget</option><option value="blank">Blank model</option></select></label>
          {seedMode==="previous_budget"?<label className="space-y-1 text-xs"><span>Previous plan</span><select className="h-10 w-full rounded-md border bg-background px-3" value={sourceVersion} onChange={e=>setSourceVersion(e.target.value)}><option value="">Choose a plan</option>{context?.native_versions?.map(v=><option key={v.id} value={`native:${v.id}`}>{v.version_name}</option>)}{context?.imported_versions?.filter(v=>v.plan_type==="budget").map(v=><option key={v.version_name} value={`imported:${v.version_name}`}>{v.version_name} · imported</option>)}</select></label>:null}
          <label className="space-y-1 text-xs"><span>Starting detail</span><select className="h-10 w-full rounded-md border bg-background px-3" value={detailLevel} onChange={e=>setDetailLevel(e.target.value)}><option value="high_level">High level — P&L groups</option><option value="detailed">Detailed — mapped COA</option></select></label>
          {seedMode!=="blank"?<label className="space-y-1 text-xs"><span>Initial growth / change (%)</span><Input type="number" step="0.1" value={growth} onChange={e=>setGrowth(Number(e.target.value))}/></label>:null}
          <Button className="w-full" onClick={()=>void create()} disabled={loading||!name.trim()||(seedMode==="previous_budget"&&!sourceVersion)}>{loading?<Loader2 className="size-4 animate-spin"/>:<Sparkles className="size-4"/>}Build starting budget</Button>
        </CardContent></Card>

        <Card><CardHeader><CardTitle>Versions</CardTitle><CardDescription>Open an existing budget.</CardDescription></CardHeader><CardContent className="space-y-2">{initialLoading?<div className="flex min-h-20 items-center justify-center"><Loader2 className="size-5 animate-spin"/></div>:versions.length===0?<div className="rounded-xl border border-dashed p-4 text-sm text-muted-foreground">No native planning versions yet.</div>:versions.map(v=><button key={v.id} onClick={()=>void open(v.id)} className={`w-full rounded-xl border p-3 text-left transition hover:bg-muted ${selected?.id===v.id?"border-foreground bg-muted/50":""}`}><b>{v.version_name}</b><p className="text-xs text-muted-foreground">{v.status} · {v.plan_type} · {String(v.financial_year_start).slice(0,4)}</p></button>)}</CardContent></Card>
      </div>

      <div className="space-y-5">
        {!selected?<Card><CardContent className="flex min-h-[420px] flex-col items-center justify-center gap-4 text-center"><div className="flex size-14 items-center justify-center rounded-2xl bg-indigo-100 text-indigo-700"><Sparkles className="size-6"/></div><div><p className="font-semibold">Choose or build a budget</p><p className="mt-1 max-w-xl text-sm text-muted-foreground">FinCruiz can seed the next financial period from the last mapped actuals even though those actuals belong to a different year.</p></div></CardContent></Card>:<>
          <Card><CardHeader className="flex-row items-start justify-between gap-4"><div><CardTitle>{selected.version_name}</CardTitle><CardDescription>{selected.lines?.length??0} lines · {String(selected.financial_year_start).slice(0,10)} to {String(selected.financial_year_end).slice(0,10)}</CardDescription></div><div className="flex flex-wrap justify-end gap-2"><Button variant="outline" onClick={undo} disabled={!history.length}><History className="size-4"/>Undo edit</Button><Button variant="outline" onClick={()=>void reseed()} disabled={loading}><RotateCcw className="size-4"/>Restore seed</Button><Button variant="outline" onClick={add}><Plus className="size-4"/>Line</Button><Button onClick={()=>void save()} disabled={loading}>{loading?<Loader2 className="size-4 animate-spin"/>:<Save className="size-4"/>}Save</Button></div></CardHeader>
            <CardContent><div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-4">{["Revenue","Cost of Sales","Payroll","Operating Expenses"].map(g=><div key={g} className="rounded-2xl border bg-muted/20 p-4"><p className="text-xs text-muted-foreground">{g}</p><p className="mt-1 text-xl font-semibold">{formatMoney(annualTotals[g]??0)}</p></div>)}</div></CardContent>
          </Card>

          <Card><CardHeader><div className="flex flex-wrap items-start justify-between gap-3"><div><CardTitle>Management targets</CardTitle><CardDescription>Start with only the numbers management actually decides. FinCruiz can derive the operating envelope, phase it monthly and drill it to the mapped COA.</CardDescription></div><div className="flex rounded-xl border bg-muted/30 p-1"><button onClick={()=>setTargetMode("executive")} className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${targetMode==="executive"?"bg-background shadow-sm":"text-muted-foreground"}`}>Executive targets</button><button onClick={()=>setTargetMode("detailed")} className={`rounded-lg px-3 py-1.5 text-xs font-semibold ${targetMode==="detailed"?"bg-background shadow-sm":"text-muted-foreground"}`}>Detailed targets</button></div></div></CardHeader><CardContent className="space-y-4">
            {targetMode==="executive"?<><div className="grid gap-3 md:grid-cols-3"><label className="space-y-1 text-xs"><span>Revenue target</span><Input type="number" step="0.01" value={revenueTarget} onChange={e=>setRevenueTarget(Number(e.target.value))}/><small className="text-muted-foreground">Annual management revenue target</small></label><label className="space-y-1 text-xs"><span>Gross margin target (%)</span><Input type="number" step="0.1" value={grossMarginTarget} onChange={e=>setGrossMarginTarget(Number(e.target.value))}/><small className="text-muted-foreground">FinCruiz derives Cost of Sales</small></label><label className="space-y-1 text-xs"><span>Net profit target</span><Input type="number" step="0.01" value={netProfitTarget} onChange={e=>setNetProfitTarget(Number(e.target.value))}/><small className="text-muted-foreground">FinCruiz derives the operating-cost envelope</small></label></div><div className="rounded-xl border border-indigo-200 bg-indigo-50/50 p-4 text-sm dark:border-indigo-900 dark:bg-indigo-950/20"><b>What FinCruiz will do:</b> Revenue → GP target → Cost of Sales → allowable operating costs → Payroll/Opex mix → monthly phasing → optional GL allocation. The active workspace scope controls whether this is consolidated or branch-specific.</div></>:<div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">{groups.slice(0,7).map(g=><label key={g} className="space-y-1 text-xs"><span>{g}</span><Input type="number" step="0.01" value={targets[g]??0} onChange={e=>setTargets({...targets,[g]:Number(e.target.value)})}/></label>)}</div>}
            <div className="flex flex-wrap items-end gap-3"><label className="space-y-1 text-xs"><span>Allocate to</span><select className="h-10 rounded-md border bg-background px-3" value={allocationDetail} onChange={e=>setAllocationDetail(e.target.value)}><option value="high_level">Keep high level</option><option value="detailed">Mapped GL / COA</option></select></label><label className="space-y-1 text-xs"><span>Allocation basis</span><select className="h-10 rounded-md border bg-background px-3" value={allocationMethod} onChange={e=>setAllocationMethod(e.target.value)}><option value="historical_actuals">Historical account mix</option><option value="equal">Equal across mapped accounts</option></select></label><label className="space-y-1 text-xs"><span>Monthly spread</span><select className="h-10 rounded-md border bg-background px-3" value={seasonality} onChange={e=>setSeasonality(e.target.value)}><option value="historical">Historical monthly pattern</option><option value="equal">Equal monthly split</option></select></label><Button onClick={()=>void allocate()} disabled={loading || (targetMode==="executive" && revenueTarget<=0)}><WandSparkles className="size-4"/>Build target budget</Button></div></CardContent></Card>

          <Card><CardHeader><CardTitle>Budget lines</CardTitle><CardDescription>Manual changes are flexible. Use Undo before saving, or Restore seed to remove all manual edits and rebuild from the original source.</CardDescription></CardHeader><CardContent><div className="max-h-[620px] overflow-auto rounded-xl border"><table className="w-full min-w-[1000px] text-sm"><thead className="sticky top-0 z-10 bg-background"><tr className="border-b">{["Period","Reporting group","Account / subgroup","Account code","Amount","Notes"].map(h=><th className="p-3 text-left" key={h}>{h}</th>)}</tr></thead><tbody>{(selected.lines??[]).map((l:any,i)=><tr key={`${l.period}-${l.reporting_group}-${l.source_account_code??""}-${i}`} className="border-b"><td className="p-2"><Input type="date" value={String(l.period).slice(0,10)} onChange={e=>patchLine(i,{period:e.target.value})}/></td><td className="p-2"><select className="h-10 rounded-md border bg-background px-2" value={l.reporting_group} onChange={e=>patchLine(i,{reporting_group:e.target.value})}>{groups.map(g=><option key={g}>{g}</option>)}</select></td><td className="p-2"><Input value={l.reporting_subgroup??""} onChange={e=>patchLine(i,{reporting_subgroup:e.target.value||null})}/></td><td className="p-2"><Input value={l.source_account_code??""} onChange={e=>patchLine(i,{source_account_code:e.target.value||null})}/></td><td className="p-2"><Input type="number" step="0.01" value={l.amount} onChange={e=>patchLine(i,{amount:Number(e.target.value),driver_type:"manual"})}/></td><td className="p-2"><Input value={l.notes??""} onChange={e=>patchLine(i,{notes:e.target.value||null})}/></td></tr>)}</tbody></table></div></CardContent></Card>
        </>}
      </div>
    </div>

    <Card><CardContent className="grid gap-3 p-5 md:grid-cols-3"><Journey title="Forecast over time" text="Use actual history or this budget to project forward and compare downside/base/upside visually." href="/dashboard/forecasting"/><Journey title="Run management scenarios" text="Change price, volume, headcount, working-capital days or capex without changing the saved ledger." href="/dashboard/decision-simulator"/><Journey title="Open Visual BI" text="See revenue, profit, margin, branch and working-capital graphs without needing to know which report to run." href="/dashboard/bi"/></CardContent></Card>
  </div>;
}

function ContextStat({label,value,note}:{label:string;value:string|number;note:string}){return <div className="rounded-2xl border bg-background/80 p-4"><p className="text-xs text-muted-foreground">{label}</p><p className="mt-1 text-2xl font-semibold">{value}</p><p className="mt-1 text-xs text-muted-foreground">{note}</p></div>}
function Journey({title,text,href}:{title:string;text:string;href:string}){return <Link href={href} className="group rounded-2xl border p-4 transition hover:-translate-y-0.5 hover:shadow-md"><div className="flex items-center justify-between"><p className="font-semibold">{title}</p><ChevronRight className="size-4 transition group-hover:translate-x-1"/></div><p className="mt-2 text-sm leading-6 text-muted-foreground">{text}</p></Link>}
