"use client";

import { useEffect, useMemo, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import {
  ArrowRight, Bot, BriefcaseBusiness, CircleDollarSign, FlaskConical,
  Gauge, Loader2, Settings2, ShieldCheck, TrendingUp, Upload,
  WalletCards, WandSparkles, Sparkles,
} from "lucide-react";

import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { HelpTip } from "@/components/ui/help-tip";
import { ViewportModal } from "@/components/ui/viewport-modal";
import { AskFinCruizDashboard } from "@/components/ask-fincruiz-dashboard";
import { DailyBusinessPulse } from "@/components/daily-business-pulse";
import { LaunchReadinessCard } from "@/components/launch-readiness-card";
import { InsightChart } from "@/components/insight-chart";
import { ManagementPerformanceBoard } from "@/components/management-performance-board";
import { getApiErrorMessage } from "@/lib/api";
import { formatDays, formatMoney, formatNumber, formatPercent, formatRatio, toNumber } from "@/lib/finance-format";
import { authService } from "@/services/auth-service";
import { financeService } from "@/services/finance-service";
import { workspaceService, type WorkspaceStatus, type LaunchReadiness } from "@/services/workspace-service";
import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import type { AICFOAnswer, AICFOSignal, AnalyticsOverview } from "@/types/analytics";
import type { Company } from "@/types/auth";
import type { BalanceSheet, DataHealth, FinancialAssurance, ProfitAndLoss, Ratio, TrialBalance } from "@/types/finance";

type DashboardView = "owner" | "cfo" | "finance" | "custom";
type WidgetKey = "headline" | "metrics" | "performance" | "priorities" | "briefing" | "profitability" | "confidence" | "dataHealth";

const widgetMeta: Record<WidgetKey, { label: string; description: string }> = {
  headline: { label: "Executive headline", description: "A plain-English opening view of what matters in the business." },
  metrics: { label: "Business pulse", description: "Revenue, profit, margin and a management-relevant cash/working-capital metric." },
  performance: { label: "Performance trends", description: "Management trends, period comparisons, sparklines and working-capital direction." },
  priorities: { label: "Management priorities", description: "Ranked positive, attention and high-priority signals with evidence and actions." },
  briefing: { label: "AI executive briefing", description: "A management narrative grounded in your company data, with external context when relevant." },
  profitability: { label: "Profitability snapshot", description: "Current P&L totals for a quicker finance review." },
  confidence: { label: "Financial confidence", description: "Structural data reliability checks. More relevant to finance teams than most executives." },
  dataHealth: { label: "Data health", description: "Trial-balance, mapping, validation and duplicate checks." },
};

const presets: Record<Exclude<DashboardView, "custom">, WidgetKey[]> = {
  owner: ["headline", "metrics", "performance", "priorities"],
  cfo: ["headline", "metrics", "performance", "priorities", "briefing", "confidence"],
  finance: ["metrics", "performance", "profitability", "confidence", "dataHealth", "priorities"],
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
  const [launchReadiness, setLaunchReadiness] = useState<LaunchReadiness | null>(null);
  const [loadingDemo, setLoadingDemo] = useState(false);
  const [dataHealth, setDataHealth] = useState<DataHealth | null>(null);
  const [executiveBrief, setExecutiveBrief] = useState<AICFOAnswer | null>(null);
  const [signals, setSignals] = useState<AICFOSignal[]>([]);
  const [assurance, setAssurance] = useState<FinancialAssurance | null>(null);
  const [analyticsOverview, setAnalyticsOverview] = useState<AnalyticsOverview | null>(null);
  const [dashboardView, setDashboardView] = useState<DashboardView>("owner");
  const [customWidgets, setCustomWidgets] = useState<WidgetKey[]>(presets.owner);
  const [customizeOpen, setCustomizeOpen] = useState(false);
  const [pulseOpen, setPulseOpen] = useState(false);

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
        const [profitLoss, bs, tb, ratios, suggestions, workspaceStatus, health, assuranceResult, overviewResult, readinessResult] = await Promise.all([
          financeService.getProfitAndLoss(), financeService.getBalanceSheet(), financeService.getTrialBalance(),
          financeService.getKpis(), financeService.getMappingSuggestions(), workspaceService.getStatus(),
          financeService.getDataHealth(), financeService.getFinancialAssurance(), analyticsService.getOverview(), workspaceService.getLaunchReadiness(),
        ]);
        setPnl(profitLoss); setBalanceSheet(bs); setTrialBalance(tb); setKpis(ratios);
        setUnmappedCount(suggestions.length); setWorkspace(workspaceStatus); setDataHealth(health); setAssurance(assuranceResult); setAnalyticsOverview(overviewResult); setLaunchReadiness(readinessResult);
        if (workspaceStatus.has_financial_data) {
          analyticsService.getExecutiveBrief().then(setExecutiveBrief).catch(() => undefined);
          analyticsService.getProactiveSignals().then((r) => setSignals(r.signals)).catch(() => undefined);
        }
      } catch (loadError) { setError(getApiErrorMessage(loadError)); }
      finally { setIsLoading(false); }
    }
    void load();
  }, [router]);

  useEffect(() => {
    if (isLoading || !trialBalance?.lines?.length) return;
    const today = new Date().toISOString().slice(0, 10);
    const key = `fincruiz_daily_pulse_seen_${today}`;
    if (window.localStorage.getItem(key) === "true") return;
    const timer = window.setTimeout(() => {
      setPulseOpen(true);
      window.localStorage.setItem(key, "true");
      usageService.track("daily_business_pulse_opened", { source: "automatic" });
    }, 650);
    return () => window.clearTimeout(timer);
  }, [isLoading, trialBalance]);

  async function loadDemoWorkspace() {
    setLoadingDemo(true); setError("");
    try { await workspaceService.loadDemo(false); window.location.reload(); }
    catch (demoError) { setError(getApiErrorMessage(demoError)); setLoadingDemo(false); }
  }

  function chooseView(view: DashboardView) {
    setDashboardView(view); window.localStorage.setItem("fincruiz_dashboard_view", view);
    usageService.track("dashboard_role_view_changed", { view });
    if (view !== "custom") setCustomizeOpen(false);
  }
  function toggleWidget(widget: WidgetKey) {
    setDashboardView("custom"); window.localStorage.setItem("fincruiz_dashboard_view", "custom");
    setCustomWidgets((current) => {
      const next = current.includes(widget) ? current.filter((x) => x !== widget) : [...current, widget];
      window.localStorage.setItem("fincruiz_dashboard_widgets", JSON.stringify(next));
      usageService.track("dashboard_customized", { widgets: next });
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

  return <div className="mx-auto max-w-[1440px] space-y-5 animate-content-ready">
    <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
      <div>
        <p className="fincruiz-eyebrow">Executive workspace</p>
        <p className="mt-1 text-sm text-muted-foreground">A focused view of performance, priorities and the decisions that follow.</p>
      </div>
      <div className="flex flex-wrap items-center gap-2">
        <div className="inline-flex rounded-xl border border-border/80 bg-card p-1">
          {([[
            "owner", "Owner / CEO"
          ], ["cfo", "CFO"], ["finance", "Finance"], ["custom", "Custom"]] as [DashboardView,string][]).map(([key,label]) => <button key={key} type="button" onClick={() => chooseView(key)} className={`rounded-lg px-3 py-1.5 text-xs font-semibold transition ${dashboardView === key ? "bg-primary text-primary-foreground shadow-sm" : "text-muted-foreground hover:bg-muted hover:text-foreground"}`}>{label}</button>)}
        </div>
        <Button type="button" variant="outline" onClick={() => setCustomizeOpen(true)}><Settings2 className="size-4"/>Customize</Button>
      </div>
    </div>

    {visible.has("headline") ? <section className="fincruiz-panel overflow-hidden p-6 sm:p-8">
      <div className="grid gap-7 lg:grid-cols-[1fr_auto] lg:items-end">
        <div className="max-w-4xl">
          <div className="flex flex-wrap items-center gap-2">
            <span className="inline-flex items-center gap-2 rounded-full border border-primary/15 bg-primary/[.055] px-3 py-1.5 text-xs font-semibold text-primary"><Sparkles className="size-3.5"/>Management brief</span>
            {assurance ? <span className="inline-flex items-center gap-2 rounded-full border bg-background px-3 py-1.5 text-xs font-semibold text-muted-foreground"><span className={`size-1.5 rounded-full ${assurance.score >= 85 ? "bg-emerald-500" : assurance.score >= 70 ? "bg-amber-500" : "bg-rose-500"}`}/>Financial confidence {formatNumber(assurance.score,0)}/100</span> : null}
          </div>
          <h1 className="mt-5 text-3xl font-semibold tracking-[-.035em] sm:text-[42px] sm:leading-[1.08]">Good {greeting}. Here&apos;s what changed for {companyName}.</h1>
          <p className="mt-3 max-w-2xl text-[15px] leading-7 text-muted-foreground">Start with the management signal. Open the financial evidence only when you need to understand the driver.</p>
        </div>
        <div className="flex flex-wrap gap-2 lg:justify-end">
          <Button type="button" onClick={() => { setPulseOpen(true); usageService.track("daily_business_pulse_opened", { source: "dashboard_button" }); }}><Sparkles className="size-4"/>Daily brief</Button>
          <Link href="/dashboard/intelligence" className={buttonVariants({variant:"outline"})}><BriefcaseBusiness className="size-4"/>Investigate</Link>
          <Link href="/dashboard/reports" className={buttonVariants({variant:"outline"})}>Financials<ArrowRight className="size-4"/></Link>
        </div>
      </div>
    </section> : null}

    {launchReadiness ? <LaunchReadinessCard readiness={launchReadiness} /> : null}
    {error ? <Card className="border-destructive/30"><CardContent className="py-4 text-sm text-destructive">{error}</CardContent></Card> : null}

    {visible.has("metrics") ? <div className="stagger-grid grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
      <MetricCard title="Revenue" help="Total income generated in the current reporting data." description="Current reporting period" value={formatMoney(pnl?.revenue, currency)} icon={TrendingUp}/>
      <MetricCard title="Net profit" help="Profit remaining after operating expenses, finance costs and tax." description="After tax" value={formatMoney(pnl?.net_profit, currency)} icon={CircleDollarSign}/>
      <MetricCard title="Gross margin" help="Gross profit as a percentage of revenue. Useful for understanding pricing, product mix and direct-cost pressure." description="Revenue retained after direct costs" value={formatPercent(grossMargin)} icon={Gauge}/>
      <MetricCard title={managementKpiLabel} help={managementKpiHelp} description={managementKpi?.category || "Liquidity / working capital"} value={managementKpiValue} icon={WalletCards}/>
    </div> : null}

    {workspace?.demo_data_active ? <div className="fincruiz-panel flex items-center gap-3 px-4 py-3"><FlaskConical className="size-4 text-sky-600"/><div><p className="text-sm font-semibold">Synthetic demo workspace</p><p className="text-xs text-muted-foreground">No customer data is being used.</p></div></div> : null}

    {!hasFinancials ? <Card><CardContent className="flex flex-col items-start gap-5 py-8 sm:flex-row sm:items-center sm:justify-between"><div><p className="font-semibold">Connect your first source</p><p className="mt-1 text-sm text-muted-foreground">Upload your ledger or explore the product with synthetic data.</p></div><div className="flex gap-2"><Link href="/dashboard/uploads" className={buttonVariants()}>Upload data<Upload className="size-4"/></Link><Button variant="outline" onClick={() => void loadDemoWorkspace()} disabled={loadingDemo}>{loadingDemo ? <Loader2 className="size-4 animate-spin"/> : <FlaskConical className="size-4"/>}Explore demo</Button></div></CardContent></Card>
    : !mappedReady ? <Card className="border-amber-200/70 bg-amber-50/50 dark:bg-amber-950/10"><CardContent className="flex flex-col items-start gap-4 py-5 sm:flex-row sm:items-center sm:justify-between"><div className="flex gap-3"><WandSparkles className="mt-0.5 size-5 text-amber-700"/><div><p className="font-semibold">One setup step needs attention</p><p className="mt-1 text-sm text-muted-foreground">Review {unmappedCount} account mapping{unmappedCount === 1 ? "" : "s"} before relying on management reporting.</p></div></div><Link href="/dashboard/mapping" className={buttonVariants()}>Review mappings<ArrowRight className="size-4"/></Link></CardContent></Card>
    : null}

    {visible.has("priorities") && hasFinancials ? <Card>
      <CardHeader><div className="flex items-start justify-between gap-3"><div><p className="fincruiz-eyebrow">Management attention</p><CardTitle className="mt-1 text-xl">Three signals worth your time</CardTitle><CardDescription>Ranked from the underlying finance checks before AI adds explanation.</CardDescription></div><HelpTip text="These signals are generated from deterministic finance checks first. AI can then explain them or add external context." side="left"/></div></CardHeader>
      <CardContent className="grid gap-3 lg:grid-cols-3">
        {signals.length ? signals.slice(0,3).map((signal,index) => <div key={`${signal.title}-${index}`} className="fincruiz-priority"><div className="flex items-center justify-between gap-3"><span className={`rounded-full px-2.5 py-1 text-[10px] font-bold uppercase tracking-[.08em] ${signal.severity === "high" ? "bg-rose-100 text-rose-700 dark:bg-rose-950/30 dark:text-rose-300" : signal.severity === "positive" ? "bg-emerald-100 text-emerald-700 dark:bg-emerald-950/30 dark:text-emerald-300" : "bg-amber-100 text-amber-700 dark:bg-amber-950/30 dark:text-amber-300"}`}>{signal.severity === "high" ? "Priority" : signal.severity === "positive" ? "Positive" : "Attention"}</span><span className="text-xs tabular-nums text-muted-foreground">0{index + 1}</span></div><p className="mt-4 font-semibold leading-6 tracking-[-.01em]">{signal.title}</p><p className="mt-2 text-sm leading-6 text-muted-foreground">{signal.evidence}</p><div className="mt-4 border-t pt-3 text-xs leading-5 text-muted-foreground"><b className="text-foreground">Next:</b> {signal.action}</div></div>) : <div className="col-span-full rounded-2xl bg-muted/40 p-6 text-sm text-muted-foreground">No material signal is available yet. Add enough monthly data to compare changes over time.</div>}
      </CardContent>
    </Card> : null}

    <AskFinCruizDashboard />

    {visible.has("performance") && hasFinancials ? <ManagementPerformanceBoard overview={analyticsOverview} currency={currency} /> : null}

    {visible.has("briefing") && executiveBrief ? <Card className="border-primary/15 bg-primary/[.025]"><CardHeader><div className="flex items-start justify-between gap-3"><div><p className="fincruiz-eyebrow">Interpretation</p><CardTitle className="mt-1 flex items-center gap-2 text-xl"><Bot className="size-5"/>AI executive briefing</CardTitle><CardDescription>Company evidence first; external industry/economic context only where relevant.</CardDescription></div><HelpTip text="The AI explanation is not the source of truth for calculations. FinCruiz prepares financial context first, then asks AI to explain it." side="left"/></div></CardHeader><CardContent><div className="max-w-4xl whitespace-pre-wrap text-sm leading-7">{executiveBrief.answer}</div>{executiveBrief.visualization ? <InsightChart visualization={executiveBrief.visualization} /> : null}<div className="mt-5 flex flex-wrap gap-2"><Link href="/dashboard/intelligence" className={buttonVariants({variant:"outline"})}>Open evidence<ArrowRight className="size-4"/></Link><Link href="/dashboard/three-way-forecast" className={buttonVariants({variant:"outline"})}>Model the decision<TrendingUp className="size-4"/></Link></div></CardContent></Card> : null}

    {visible.has("profitability") ? <Card><CardHeader><CardTitle>Profitability detail</CardTitle><CardDescription>Current P&L totals for a faster finance review.</CardDescription></CardHeader><CardContent className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">{[
      ["Revenue", pnl?.revenue], ["Cost of sales", pnl?.cost_of_sales], ["Gross profit", pnl?.gross_profit], ["Operating expenses", pnl?.operating_expenses], ["Net profit", pnl?.net_profit],
    ].map(([label,value]) => <div key={String(label)} className="rounded-xl border border-border/80 bg-muted/20 p-4"><p className="text-xs text-muted-foreground">{label}</p><p className="mt-2 text-lg font-semibold tabular-nums tracking-[-.02em]">{formatMoney(value as string | number | undefined, currency)}</p></div>)}</CardContent></Card> : null}

    {visible.has("confidence") && hasFinancials && assurance ? <Card><CardHeader><div className="flex items-start justify-between"><div><CardTitle className="flex items-center gap-2"><ShieldCheck className="size-5"/>Financial confidence</CardTitle><CardDescription>Structural reliability of the data feeding reports and AI.</CardDescription></div><HelpTip text="This is a structural data-quality score, not an audit opinion. It checks balance, mapping, validation and similar controls." side="left"/></div></CardHeader><CardContent><div className="flex flex-col gap-5 lg:flex-row lg:items-center"><div className="min-w-40 rounded-2xl bg-slate-950 p-5 text-white"><p className="text-sm text-slate-400">Confidence</p><p className="mt-1 text-4xl font-semibold tracking-[-.04em]">{formatNumber(assurance.score, 0)}<span className="text-base font-medium text-slate-400"> /100</span></p><p className="mt-2 text-xs text-slate-400">Grade {assurance.grade}</p></div><div className="grid flex-1 gap-3 sm:grid-cols-2 xl:grid-cols-3">{assurance.checks.slice(0,6).map((check) => <div key={check.key} className="rounded-xl border border-border/80 p-3"><div className="flex items-center gap-2"><span className={`size-2 rounded-full ${check.status === "pass" ? "bg-emerald-500" : check.status === "warning" ? "bg-amber-500" : "bg-red-500"}`}/><span className="text-sm font-medium">{check.label}</span></div><p className="mt-2 text-xs leading-5 text-muted-foreground">{check.detail}</p></div>)}</div></div></CardContent></Card> : null}

    {visible.has("dataHealth") ? <Card><CardHeader><div className="flex items-start justify-between"><div><CardTitle>Data health</CardTitle><CardDescription>Technical checks for finance users.</CardDescription></div><HelpTip text="Executives usually do not need this on their default dashboard. Finance users can keep it visible as a control panel." side="left"/></div></CardHeader><CardContent className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5"><Status label="Trial balance difference" value={formatMoney(dataHealth?.trial_balance_difference ?? trialBalance?.difference, currency, 2)} good={Boolean(dataHealth?.is_trial_balance_balanced)}/><Status label="Balance sheet difference" value={formatMoney(dataHealth?.balance_sheet_difference, currency, 2)} good={Boolean(dataHealth?.is_balance_sheet_balanced)}/><Status label="Mapped accounts" value={`${dataHealth?.mapped_account_count ?? 0}/${dataHealth?.account_count ?? 0}`} good={Boolean(dataHealth?.is_mapping_complete && dataHealth?.account_count)}/><Status label="Invalid transactions" value={String(dataHealth?.invalid_transaction_count ?? 0)} good={(dataHealth?.invalid_transaction_count ?? 0) === 0}/><Status label="Potential duplicates" value={String(dataHealth?.duplicate_candidate_count ?? 0)} good={(dataHealth?.duplicate_candidate_count ?? 0) === 0}/></CardContent></Card> : null}

    <DailyBusinessPulse
      open={pulseOpen}
      onClose={() => setPulseOpen(false)}
      signals={signals}
      revenue={toNumber(pnl?.revenue)}
      netProfit={toNumber(pnl?.net_profit)}
      grossMargin={grossMargin}
      currency={currency}
      companyName={companyName}
    />

    <ViewportModal
      open={customizeOpen}
      onClose={() => setCustomizeOpen(false)}
      title="Customize your dashboard"
      description="Role presets change what is emphasized, not what you are allowed to access."
      footer={<div className="flex justify-end"><Button onClick={() => setCustomizeOpen(false)}>Done</Button></div>}
    >
      <div className="grid gap-3">
        {(Object.keys(widgetMeta) as WidgetKey[]).map((key) => {
          const checked = visible.has(key);
          return <label key={key} className="flex cursor-pointer items-start gap-3 rounded-2xl border p-4 transition hover:bg-muted/40"><input type="checkbox" className="mt-1" checked={checked} onChange={() => toggleWidget(key)}/><div><p className="font-medium">{widgetMeta[key].label}</p><p className="mt-1 text-sm leading-6 text-muted-foreground">{widgetMeta[key].description}</p></div></label>;
        })}
      </div>
    </ViewportModal>
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
  return <div className="fincruiz-kpi group"><div className="flex items-start justify-between gap-3"><div><div className="flex items-center gap-1.5"><p className="text-xs font-semibold text-muted-foreground">{title}</p><HelpTip title={title} text={help} side="bottom"/></div><p className="mt-1 text-[11px] text-muted-foreground/80">{description}</p></div><span className="flex size-8 items-center justify-center rounded-xl bg-primary/[.07] text-primary"><Icon className="size-4"/></span></div><p className="mt-6 truncate text-[28px] font-semibold tabular-nums tracking-[-.045em]">{value}</p></div>;
}

function Status({ label, value, good }: { label: string; value: string; good: boolean }) {
  return <div className="rounded-xl border p-3"><div className="flex items-center gap-2"><span className={`size-2 rounded-full ${good ? "bg-emerald-500" : "bg-amber-500"}`}/><span className="text-xs font-medium text-muted-foreground">{label}</span></div><p className="mt-2 font-semibold">{value}</p></div>;
}
