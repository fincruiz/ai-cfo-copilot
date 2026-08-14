"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import {
  ArrowRight, Bot, BriefcaseBusiness, CheckCircle2, CircleDollarSign, Eye, FlaskConical,
  Gauge, Info, LayoutGrid, Loader2, Settings2, ShieldCheck, TrendingUp, Upload,
  WalletCards, WandSparkles, X, Sparkles,
} from "lucide-react";

import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { HelpTip } from "@/components/ui/help-tip";
import { getApiErrorMessage } from "@/lib/api";
import { formatDays, formatMoney, formatNumber, formatPercent, formatRatio, toNumber } from "@/lib/finance-format";
import { authService } from "@/services/auth-service";
import { financeService } from "@/services/finance-service";
import { workspaceService, type WorkspaceStatus } from "@/services/workspace-service";
import { analyticsService } from "@/services/analytics-service";
import type { AICFOAnswer, AICFOSignal } from "@/types/analytics";
import type { Company } from "@/types/auth";
import type { BalanceSheet, DataHealth, FinancialAssurance, ProfitAndLoss, Ratio, TrialBalance } from "@/types/finance";

type DashboardView = "owner" | "cfo" | "finance" | "custom";
type WidgetKey = "headline" | "metrics" | "priorities" | "briefing" | "profitability" | "confidence" | "dataHealth";

const widgetMeta: Record<WidgetKey, { label: string; description: string }> = {
  headline: { label: "Executive headline", description: "A plain-English opening view of what matters in the business." },
  metrics: { label: "Business pulse", description: "Revenue, profit, margin and a management-relevant cash/working-capital metric." },
  priorities: { label: "Management priorities", description: "Ranked positive, attention and high-priority signals with evidence and actions." },
  briefing: { label: "AI executive briefing", description: "A management narrative grounded in your company data, with external context when relevant." },
  profitability: { label: "Profitability snapshot", description: "Current P&L totals for a quicker finance review." },
  confidence: { label: "Financial confidence", description: "Structural data reliability checks. More relevant to finance teams than most executives." },
  dataHealth: { label: "Data health", description: "Trial-balance, mapping, validation and duplicate checks." },
};

const presets: Record<Exclude<DashboardView, "custom">, WidgetKey[]> = {
  owner: ["headline", "metrics", "priorities", "briefing"],
  cfo: ["headline", "metrics", "priorities", "briefing", "profitability", "confidence"],
  finance: ["metrics", "profitability", "confidence", "dataHealth", "priorities"],
};

export default function DashboardPage() {
  const router = useRouter();
  const [company, setCompany] = useState<Company | null>(null);
  const [pnl, setPnl] = useState<ProfitAndLoss | null>(null);
  const [balanceSheet, setBalanceSheet] = useState<BalanceSheet | null>(null);
  const [trialBalance, setTrialBalance] = useState<TrialBalance | null>(null);
  const [kpis, setKpis] = useState<Ratio[]>([]);
  const [unmappedCount, setUnmappedCount] = useState(0);
  const [isLoading, setIsLoading] = useState(true);
  const [error, setError] = useState("");
  const [workspace, setWorkspace] = useState<WorkspaceStatus | null>(null);
  const [loadingDemo, setLoadingDemo] = useState(false);
  const [dataHealth, setDataHealth] = useState<DataHealth | null>(null);
  const [executiveBrief, setExecutiveBrief] = useState<AICFOAnswer | null>(null);
  const [signals, setSignals] = useState<AICFOSignal[]>([]);
  const [assurance, setAssurance] = useState<FinancialAssurance | null>(null);
  const [dashboardView, setDashboardView] = useState<DashboardView>("owner");
  const [customWidgets, setCustomWidgets] = useState<WidgetKey[]>(presets.owner);
  const [customizeOpen, setCustomizeOpen] = useState(false);

  useEffect(() => {
    const storedView = window.localStorage.getItem("fincruiz_dashboard_view") as DashboardView | null;
    const storedWidgets = window.localStorage.getItem("fincruiz_dashboard_widgets");
    if (storedView && ["owner", "cfo", "finance", "custom"].includes(storedView)) setDashboardView(storedView);
    if (storedWidgets) {
      try { setCustomWidgets(JSON.parse(storedWidgets)); } catch { /* ignore stale setting */ }
    }

    async function load() {
      if (!authService.hasAccessToken()) { router.replace("/login"); return; }
      try {
        const currentCompany = await authService.getCurrentCompany();
        setCompany(currentCompany);
        const [profitLoss, bs, tb, ratios, suggestions, workspaceStatus, health, assuranceResult] = await Promise.all([
          financeService.getProfitAndLoss(), financeService.getBalanceSheet(), financeService.getTrialBalance(),
          financeService.getKpis(), financeService.getMappingSuggestions(), workspaceService.getStatus(),
          financeService.getDataHealth(), financeService.getFinancialAssurance(),
        ]);
        setPnl(profitLoss); setBalanceSheet(bs); setTrialBalance(tb); setKpis(ratios);
        setUnmappedCount(suggestions.length); setWorkspace(workspaceStatus); setDataHealth(health); setAssurance(assuranceResult);
        if (workspaceStatus.has_financial_data) {
          analyticsService.getExecutiveBrief().then(setExecutiveBrief).catch(() => undefined);
          analyticsService.getProactiveSignals().then((r) => setSignals(r.signals)).catch(() => undefined);
        }
      } catch (loadError) { setError(getApiErrorMessage(loadError)); }
      finally { setIsLoading(false); }
    }
    void load();
  }, [router]);

  async function loadDemoWorkspace() {
    setLoadingDemo(true); setError("");
    try { await workspaceService.loadDemo(false); window.location.reload(); }
    catch (demoError) { setError(getApiErrorMessage(demoError)); setLoadingDemo(false); }
  }

  function chooseView(view: DashboardView) {
    setDashboardView(view); window.localStorage.setItem("fincruiz_dashboard_view", view);
    if (view !== "custom") setCustomizeOpen(false);
  }
  function toggleWidget(widget: WidgetKey) {
    setDashboardView("custom"); window.localStorage.setItem("fincruiz_dashboard_view", "custom");
    setCustomWidgets((current) => {
      const next = current.includes(widget) ? current.filter((x) => x !== widget) : [...current, widget];
      window.localStorage.setItem("fincruiz_dashboard_widgets", JSON.stringify(next));
      return next;
    });
  }

  if (isLoading) return <div className="flex min-h-[60vh] items-center justify-center gap-3 text-muted-foreground"><Loader2 className="size-5 animate-spin"/>Loading your dashboard...</div>;

  const companyName = company?.trading_name ?? company?.legal_name ?? "Your company";
  const currency = company?.currency_code ?? "AUD";
  const hasFinancials = Boolean(trialBalance?.lines.length);
  const mappedReady = unmappedCount === 0 && hasFinancials;
  const visible = new Set<WidgetKey>(dashboardView === "custom" ? customWidgets : presets[dashboardView]);
  const grossMargin = toNumber(pnl?.revenue) ? (toNumber(pnl?.gross_profit) / toNumber(pnl?.revenue)) * 100 : 0;
  const managementKpi = kpis.find((item) => /cash conversion|dso|current ratio|cash/i.test(item.name));
  const managementKpiValue = managementKpi ? formatKpi(managementKpi) : formatMoney(balanceSheet?.current_assets, currency);
  const managementKpiLabel = managementKpi?.name ?? "Current assets";
  const managementKpiHelp = managementKpi?.interpretation || "A management-relevant liquidity or working-capital indicator from the latest data.";
  const greeting = new Date().getHours() < 12 ? "morning" : new Date().getHours() < 18 ? "afternoon" : "evening";

  return <div className="mx-auto max-w-7xl space-y-6">
    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <div><p className="text-xs font-semibold uppercase tracking-[.18em] text-muted-foreground">Dashboard view</p><div className="mt-2 flex flex-wrap gap-2">{([
        ["owner", "Owner / CEO"], ["cfo", "CFO"], ["finance", "Finance team"], ["custom", "Custom"],
      ] as [DashboardView,string][]).map(([key,label]) => <button key={key} type="button" onClick={() => chooseView(key)} className={`rounded-full border px-3 py-1.5 text-xs font-semibold transition ${dashboardView === key ? "border-primary bg-primary text-primary-foreground" : "bg-background hover:bg-muted"}`}>{label}</button>)}</div></div>
      <Button type="button" variant="outline" onClick={() => setCustomizeOpen(true)}><Settings2 className="size-4"/>Customize dashboard</Button>
    </div>

    {visible.has("headline") ? <section className="fincruiz-hero overflow-hidden rounded-[28px] border p-6 sm:p-8">
      <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
        <div className="max-w-3xl"><p className="flex items-center gap-2 text-sm font-semibold text-primary"><Sparkles className="size-4"/>Your business, in plain English</p><h1 className="mt-2 text-3xl font-bold tracking-tight sm:text-4xl">Good {greeting}. Here is what matters for {companyName}.</h1><p className="mt-3 text-muted-foreground">Start with business outcomes. Drill into the finance mechanics only when you need the evidence.</p></div>
        <div className="flex flex-wrap gap-2"><Link href="/dashboard/intelligence" className={buttonVariants()}><BriefcaseBusiness className="size-4"/>What needs attention?</Link><Link href="/dashboard/reports" className={buttonVariants({variant:"outline"})}>See financials<ArrowRight className="size-4"/></Link></div>
      </div>
    </section> : null}

    {error ? <Card className="border-destructive/30"><CardContent className="py-4 text-sm text-destructive">{error}</CardContent></Card> : null}

    {visible.has("metrics") ? <div className="grid gap-4 sm:grid-cols-2 xl:grid-cols-4">
      <MetricCard title="Revenue" help="Total income generated in the current reporting data." description="Current reporting period" value={formatMoney(pnl?.revenue, currency)} icon={TrendingUp}/>
      <MetricCard title="Net profit" help="Profit remaining after operating expenses, finance costs and tax." description="After tax" value={formatMoney(pnl?.net_profit, currency)} icon={CircleDollarSign}/>
      <MetricCard title="Gross margin" help="Gross profit as a percentage of revenue. Useful for understanding pricing, product mix and direct-cost pressure." description="Revenue retained after direct costs" value={formatPercent(grossMargin)} icon={Gauge}/>
      <MetricCard title={managementKpiLabel} help={managementKpiHelp} description={managementKpi?.category || "Liquidity / working capital"} value={managementKpiValue} icon={WalletCards}/>
    </div> : null}

    {workspace?.demo_data_active ? <Card className="border-sky-200 bg-sky-50/50"><CardContent className="flex items-center gap-3 py-4"><FlaskConical className="size-5 text-sky-700"/><div><p className="font-semibold">Demo workspace active</p><p className="text-sm text-muted-foreground">Synthetic company data is active. You can reset it without affecting your profile.</p></div></CardContent></Card> : null}

    {!hasFinancials ? <Card><CardContent className="flex flex-col items-start gap-5 py-8 sm:flex-row sm:items-center sm:justify-between"><div><p className="font-semibold">Choose how you want to start</p><p className="mt-1 text-sm text-muted-foreground">Upload your own data or explore a synthetic business first.</p></div><div className="flex gap-2"><Link href="/dashboard/uploads" className={buttonVariants()}>Upload your data<Upload className="size-4"/></Link><Button variant="outline" onClick={() => void loadDemoWorkspace()} disabled={loadingDemo}>{loadingDemo ? <Loader2 className="size-4 animate-spin"/> : <FlaskConical className="size-4"/>}Explore demo</Button></div></CardContent></Card>
    : !mappedReady ? <Card className="border-amber-200 bg-amber-50/40"><CardContent className="flex flex-col items-start gap-4 py-5 sm:flex-row sm:items-center sm:justify-between"><div className="flex gap-3"><WandSparkles className="mt-0.5 size-5 text-amber-700"/><div><p className="font-semibold">One setup step needs attention</p><p className="mt-1 text-sm text-muted-foreground">Review {unmappedCount} account mapping{unmappedCount === 1 ? "" : "s"} before relying on management reporting.</p></div></div><Link href="/dashboard/mapping" className={buttonVariants()}>Review mappings<ArrowRight className="size-4"/></Link></CardContent></Card>
    : null}

    {visible.has("priorities") && hasFinancials ? <Card className="overflow-hidden"><CardHeader><div className="flex items-start justify-between gap-3"><div><CardTitle>What management should focus on</CardTitle><CardDescription>FinCruiz ranks material movements before asking you to choose a report.</CardDescription></div><HelpTip text="These signals are generated from deterministic finance checks first. AI can then explain them or add external context." side="left"/></div></CardHeader><CardContent className="grid gap-3 lg:grid-cols-3">
      {signals.length ? signals.slice(0,3).map((signal,index) => <div key={`${signal.title}-${index}`} className="rounded-2xl border p-4"><div className="flex items-center justify-between gap-3"><span className={`rounded-full px-2 py-1 text-[11px] font-semibold uppercase ${signal.severity === "high" ? "bg-red-100 text-red-800" : signal.severity === "positive" ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}`}>{signal.severity === "high" ? "Priority" : signal.severity === "positive" ? "Positive" : "Attention"}</span><span className="text-xs text-muted-foreground">{index + 1}</span></div><p className="mt-3 font-semibold leading-6">{signal.title}</p><p className="mt-2 text-sm leading-6 text-muted-foreground">{signal.evidence}</p><div className="mt-4 rounded-xl bg-muted/50 p-3 text-sm"><b>Next action:</b> {signal.action}</div></div>) : <div className="col-span-full rounded-2xl bg-muted/40 p-6 text-sm text-muted-foreground">No material signal is available yet. Add enough monthly data to compare changes over time.</div>}
    </CardContent></Card> : null}

    {visible.has("briefing") && executiveBrief ? <Card className="border-primary/20 bg-primary/[.035]"><CardHeader><div className="flex items-start justify-between gap-3"><div><CardTitle className="flex items-center gap-2"><Bot className="size-5"/>AI Executive Briefing</CardTitle><CardDescription>Company evidence first; external industry/economic context only where relevant.</CardDescription></div><HelpTip text="The AI explanation is not the source of truth for calculations. FinCruiz prepares financial context first, then asks AI to explain it." side="left"/></div></CardHeader><CardContent><div className="whitespace-pre-wrap text-sm leading-7">{executiveBrief.answer}</div><div className="mt-5 flex flex-wrap gap-2"><Link href="/dashboard/intelligence" className={buttonVariants({variant:"outline"})}>Investigate in Intelligence Center<ArrowRight className="size-4"/></Link><Link href="/dashboard/three-way-forecast" className={buttonVariants({variant:"outline"})}>Model a decision<TrendingUp className="size-4"/></Link></div></CardContent></Card> : null}

    {visible.has("profitability") ? <Card><CardHeader><CardTitle>Profitability snapshot</CardTitle><CardDescription>Finance view of the current report totals.</CardDescription></CardHeader><CardContent className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">{[
      ["Revenue", pnl?.revenue], ["Cost of sales", pnl?.cost_of_sales], ["Gross profit", pnl?.gross_profit], ["Operating expenses", pnl?.operating_expenses], ["Net profit", pnl?.net_profit],
    ].map(([label,value]) => <div key={String(label)} className="rounded-2xl border p-4"><p className="text-xs text-muted-foreground">{label}</p><p className="mt-2 text-lg font-semibold tabular-nums">{formatMoney(value as string | number | undefined, currency)}</p></div>)}</CardContent></Card> : null}

    {visible.has("confidence") && hasFinancials && assurance ? <Card><CardHeader><div className="flex items-start justify-between"><div><CardTitle className="flex items-center gap-2"><ShieldCheck className="size-5"/>Financial confidence</CardTitle><CardDescription>Structural reliability of the data feeding reports and AI.</CardDescription></div><HelpTip text="This is a structural data-quality score, not an audit opinion. It checks balance, mapping, validation and similar controls." side="left"/></div></CardHeader><CardContent><div className="flex flex-col gap-5 lg:flex-row lg:items-center"><div className="min-w-40 rounded-2xl bg-slate-950 p-5 text-white"><p className="text-sm text-slate-400">Confidence</p><p className="mt-1 text-4xl font-bold">{formatNumber(assurance.score, 0)}<span className="text-base font-medium text-slate-400"> /100</span></p><p className="mt-2 text-xs text-slate-400">Grade {assurance.grade}</p></div><div className="grid flex-1 gap-3 sm:grid-cols-2 xl:grid-cols-3">{assurance.checks.slice(0,6).map((check) => <div key={check.key} className="rounded-xl border p-3"><div className="flex items-center gap-2"><span className={`size-2 rounded-full ${check.status === "pass" ? "bg-emerald-500" : check.status === "warning" ? "bg-amber-500" : "bg-red-500"}`}/><span className="text-sm font-medium">{check.label}</span></div><p className="mt-2 text-xs leading-5 text-muted-foreground">{check.detail}</p></div>)}</div></div></CardContent></Card> : null}

    {visible.has("dataHealth") ? <Card><CardHeader><div className="flex items-start justify-between"><div><CardTitle>Data health</CardTitle><CardDescription>Technical checks for the finance team.</CardDescription></div><HelpTip text="Executives usually do not need this on their default dashboard. Finance users can keep it visible as a control panel." side="left"/></div></CardHeader><CardContent className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5"><Status label="Trial balance difference" value={formatMoney(dataHealth?.trial_balance_difference ?? trialBalance?.difference, currency, 2)} good={Boolean(dataHealth?.is_trial_balance_balanced)}/><Status label="Balance sheet difference" value={formatMoney(dataHealth?.balance_sheet_difference, currency, 2)} good={Boolean(dataHealth?.is_balance_sheet_balanced)}/><Status label="Mapped accounts" value={`${dataHealth?.mapped_account_count ?? 0}/${dataHealth?.account_count ?? 0}`} good={Boolean(dataHealth?.is_mapping_complete && dataHealth?.account_count)}/><Status label="Invalid transactions" value={String(dataHealth?.invalid_transaction_count ?? 0)} good={(dataHealth?.invalid_transaction_count ?? 0) === 0}/><Status label="Potential duplicates" value={String(dataHealth?.duplicate_candidate_count ?? 0)} good={(dataHealth?.duplicate_candidate_count ?? 0) === 0}/></CardContent></Card> : null}

    <Card className="border-dashed"><CardContent className="flex flex-col gap-4 py-5 sm:flex-row sm:items-center sm:justify-between"><div><p className="font-semibold">Need a feature you cannot see here?</p><p className="mt-1 text-sm text-muted-foreground">Use <b>Explore FinCruiz</b> in the sidebar. Simplified navigation never removes capabilities.</p></div><Button variant="outline" onClick={() => setCustomizeOpen(true)}><LayoutGrid className="size-4"/>Adjust this dashboard</Button></CardContent></Card>

    {customizeOpen ? <div className="fixed inset-0 z-[130] flex items-center justify-center bg-slate-950/45 p-4 backdrop-blur-sm" onMouseDown={(e) => e.target === e.currentTarget && setCustomizeOpen(false)}><div className="w-full max-w-2xl rounded-[26px] border bg-background p-5 shadow-2xl sm:p-6"><div className="flex items-start justify-between gap-4"><div><h2 className="text-xl font-semibold">Customize your dashboard</h2><p className="mt-1 text-sm text-muted-foreground">Role presets change what is emphasized, not what you are allowed to access.</p></div><button type="button" onClick={() => setCustomizeOpen(false)} className="flex size-9 items-center justify-center rounded-xl hover:bg-muted"><X className="size-4"/></button></div><div className="mt-5 grid gap-3">{(Object.keys(widgetMeta) as WidgetKey[]).map((key) => { const checked = visible.has(key); return <label key={key} className="flex cursor-pointer items-start gap-3 rounded-2xl border p-4 hover:bg-muted/40"><input type="checkbox" className="mt-1" checked={checked} onChange={() => toggleWidget(key)}/><div><p className="font-medium">{widgetMeta[key].label}</p><p className="mt-1 text-sm leading-6 text-muted-foreground">{widgetMeta[key].description}</p></div></label>; })}</div><div className="mt-5 flex justify-end"><Button onClick={() => setCustomizeOpen(false)}>Done</Button></div></div></div> : null}
  </div>;
}

function formatKpi(kpi: Ratio): string {
  const unit = (kpi.unit || "").toLowerCase();
  if (unit.includes("day")) return formatDays(kpi.value);
  if (unit.includes("%") || unit.includes("percent")) return formatPercent(kpi.value);
  if (unit.includes("x") || unit.includes("ratio")) return formatRatio(kpi.value);
  return formatNumber(kpi.value, 2);
}

function MetricCard({ title, description, value, icon: Icon, help }: { title: string; description: string; value: string; icon: typeof TrendingUp; help: string }) {
  return <Card className="overflow-hidden"><CardHeader className="pb-2"><div className="flex items-start justify-between gap-3"><div className="flex items-center gap-1.5"><CardTitle className="text-base">{title}</CardTitle><HelpTip text={help} side="top"/></div><Icon className="size-5 text-muted-foreground"/></div><CardDescription>{description}</CardDescription></CardHeader><CardContent><p className="truncate text-2xl font-semibold tabular-nums">{value}</p></CardContent></Card>;
}

function Status({ label, value, good }: { label: string; value: string; good: boolean }) {
  return <div className="rounded-xl border p-3"><div className="flex items-center gap-2"><span className={`size-2 rounded-full ${good ? "bg-emerald-500" : "bg-amber-500"}`}/><span className="text-xs font-medium text-muted-foreground">{label}</span></div><p className="mt-2 font-semibold">{value}</p></div>;
}
