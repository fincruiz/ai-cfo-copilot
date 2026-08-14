"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { Bot, ArrowRight, Building2, CheckCircle2, CircleDollarSign, FlaskConical, Gauge, Loader2, ShieldCheck, TrendingUp, Upload, WalletCards, WandSparkles } from "lucide-react";

import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getApiErrorMessage } from "@/lib/api";
import { formatMoney, toNumber } from "@/lib/finance-format";
import { authService } from "@/services/auth-service";
import { financeService } from "@/services/finance-service";
import { workspaceService, type WorkspaceStatus } from "@/services/workspace-service";
import { analyticsService } from "@/services/analytics-service";
import type { AICFOAnswer, AICFOSignal } from "@/types/analytics";
import type { Company } from "@/types/auth";
import type { BalanceSheet, DataHealth, FinancialAssurance, ProfitAndLoss, Ratio, TrialBalance } from "@/types/finance";

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

  useEffect(() => {
    async function load() {
      if (!authService.hasAccessToken()) {
        router.replace("/login");
        return;
      }
      try {
        const currentCompany = await authService.getCurrentCompany();
        setCompany(currentCompany);
        const [profitLoss, bs, tb, ratios, suggestions, workspaceStatus, health, assuranceResult] = await Promise.all([
          financeService.getProfitAndLoss(),
          financeService.getBalanceSheet(),
          financeService.getTrialBalance(),
          financeService.getKpis(),
          financeService.getMappingSuggestions(),
          workspaceService.getStatus(),
          financeService.getDataHealth(),
          financeService.getFinancialAssurance(),
        ]);
        setPnl(profitLoss);
        setBalanceSheet(bs);
        setTrialBalance(tb);
        setKpis(ratios);
        setUnmappedCount(suggestions.length);
        setWorkspace(workspaceStatus);
        setDataHealth(health);
        setAssurance(assuranceResult);
        if (workspaceStatus.has_financial_data) {
          analyticsService.getExecutiveBrief().then(setExecutiveBrief).catch(() => undefined);
          analyticsService.getProactiveSignals().then((r) => setSignals(r.signals)).catch(() => undefined);
        }
      } catch (loadError) {
        setError(getApiErrorMessage(loadError));
      } finally {
        setIsLoading(false);
      }
    }
    void load();
  }, [router]);


  async function loadDemoWorkspace() {
    setLoadingDemo(true);
    setError("");
    try {
      await workspaceService.loadDemo(false);
      window.location.reload();
    } catch (demoError) {
      setError(getApiErrorMessage(demoError));
      setLoadingDemo(false);
    }
  }

  if (isLoading) return <div className="flex min-h-[60vh] items-center justify-center gap-3 text-muted-foreground"><Loader2 className="size-5 animate-spin" />Loading your dashboard...</div>;

  const companyName = company?.trading_name ?? company?.legal_name ?? "Your company";
  const currency = company?.currency_code ?? "AUD";
  const cashRatio = kpis.find((item) => item.name.toLowerCase().includes("cash"));
  const hasFinancials = Boolean(trialBalance?.lines.length);
  const mappedReady = unmappedCount === 0 && hasFinancials;

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <section className="fincruiz-hero overflow-hidden rounded-[28px] border p-6 sm:p-8">
        <div className="flex flex-col gap-6 lg:flex-row lg:items-center lg:justify-between">
          <div className="max-w-2xl"><p className="text-sm font-semibold text-primary">Your business, in plain English</p><h1 className="mt-2 text-3xl font-bold tracking-tight sm:text-4xl">Good {new Date().getHours() < 12 ? "morning" : new Date().getHours() < 18 ? "afternoon" : "evening"}. Here is what matters for {companyName}.</h1><p className="mt-3 text-muted-foreground">FinCruiz checks the accounting structure first, then turns it into decisions, risks and next actions.</p></div>
          <div className="flex flex-wrap gap-2"><Link href="/dashboard/uploads" className={buttonVariants()}><Upload className="size-4" />Update my data</Link><Link href="/dashboard/reports" className={buttonVariants({variant:"outline"})}>See the numbers<ArrowRight className="size-4"/></Link></div>
        </div>
      </section>

      {error ? <Card className="border-destructive/30"><CardContent className="py-4 text-sm text-destructive">{error}</CardContent></Card> : null}

      <div className="grid gap-5 md:grid-cols-2 xl:grid-cols-4">
        <MetricCard title="Current company" description={`${company?.country_code ?? ""} · ${currency}`} value={companyName} icon={Building2} />
        <MetricCard title="Revenue" description="Current reporting data" value={formatMoney(pnl?.revenue, currency)} icon={TrendingUp} />
        <MetricCard title="Net profit" description="After tax" value={formatMoney(pnl?.net_profit, currency)} icon={CircleDollarSign} />
        <MetricCard title="Total assets" description={cashRatio ? `${cashRatio.name}: ${cashRatio.value ?? "—"}` : "Balance sheet"} value={formatMoney(balanceSheet?.total_assets, currency)} icon={WalletCards} />
      </div>

      {workspace?.demo_data_active ? (
        <Card className="border-sky-200 bg-sky-50/50"><CardContent className="flex items-center gap-3 py-5"><FlaskConical className="size-5 text-sky-700" /><div><p className="font-semibold">Demo workspace active</p><p className="text-sm text-muted-foreground">You are exploring synthetic FinCruiz data. Reset it anytime from Settings before uploading real company data.</p></div></CardContent></Card>
      ) : null}

      {!hasFinancials ? (
        <Card><CardContent className="flex flex-col items-start gap-5 py-8 sm:flex-row sm:items-center sm:justify-between"><div><p className="font-semibold">Choose how you want to start</p><p className="mt-1 text-sm text-muted-foreground">Upload your own general ledger, or explore a synthetic 12-month demo first. No real company data is needed for the demo.</p></div><div className="flex flex-wrap gap-2"><Link href="/dashboard/uploads" className={buttonVariants()}>Upload your data<Upload className="size-4" /></Link><Button variant="outline" onClick={() => void loadDemoWorkspace()} disabled={loadingDemo}>{loadingDemo ? <Loader2 className="size-4 animate-spin" /> : <FlaskConical className="size-4" />}Explore demo</Button></div></CardContent></Card>
      ) : !mappedReady ? (
        <Card className="border-amber-200 bg-amber-50/40"><CardContent className="flex flex-col items-start gap-4 py-6 sm:flex-row sm:items-center sm:justify-between"><div className="flex gap-3"><WandSparkles className="mt-0.5 size-5 text-amber-700" /><div><p className="font-semibold">Review {unmappedCount} account mapping{unmappedCount === 1 ? "" : "s"}</p><p className="mt-1 text-sm text-muted-foreground">Your trial balance is ready. Approve mappings to populate P&L, balance sheet and KPIs.</p></div></div><Link href="/dashboard/mapping" className={buttonVariants()}>Review mappings<ArrowRight className="size-4" /></Link></CardContent></Card>
      ) : (
        <Card className="border-emerald-200 bg-emerald-50/40"><CardContent className="flex items-center gap-3 py-5"><CheckCircle2 className="size-5 text-emerald-700" /><div><p className="font-semibold">Finance data is ready</p><p className="text-sm text-muted-foreground">Ledger is balanced and all accounts are mapped.</p></div></CardContent></Card>
      )}


      {hasFinancials ? (
        <Card className="overflow-hidden border-primary/15">
          <CardContent className="p-0">
            <div className="grid md:grid-cols-[220px_1fr]">
              <div className="bg-primary px-6 py-6 text-primary-foreground">
                <p className="text-sm font-medium opacity-80">Workspace readiness</p>
                <p className="mt-2 text-3xl font-semibold">{mappedReady ? "100%" : "75%"}</p>
                <p className="mt-2 text-sm opacity-80">Upload → map → analyse → plan</p>
              </div>
              <div className="grid gap-3 p-5 sm:grid-cols-4">
                {[
                  ["1", "Ledger loaded", true],
                  ["2", "Accounts mapped", mappedReady],
                  ["3", "Insights active", true],
                  ["4", "Forecast ready", mappedReady],
                ].map(([step,label,done]) => <div key={String(step)} className="rounded-xl border p-3"><div className="flex items-center gap-2"><span className={`flex size-6 items-center justify-center rounded-full text-xs font-semibold ${done ? "bg-emerald-100 text-emerald-800" : "bg-muted text-muted-foreground"}`}>{step}</span><span className="text-sm font-medium">{label}</span></div></div>)}
              </div>
            </div>
          </CardContent>
        </Card>
      ) : null}


      {hasFinancials && assurance ? (
        <Card className="overflow-hidden border-primary/15">
          <CardContent className="p-0">
            <div className="grid lg:grid-cols-[220px_1fr]">
              <div className="bg-slate-950 p-6 text-white dark:bg-black/30"><div className="flex items-center gap-2 text-sm text-slate-300"><ShieldCheck className="size-4"/>Financial confidence</div><div className="mt-3 flex items-end gap-2"><span className="text-5xl font-bold">{assurance.score}</span><span className="pb-1 text-sm text-slate-400">/100 · grade {assurance.grade}</span></div><p className="mt-3 text-xs leading-5 text-slate-400">Structural reliability of the data loaded into FinCruiz.</p></div>
              <div className="p-5"><div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-3">{assurance.checks.slice(0,6).map((check)=><div key={check.key} className="rounded-xl border p-3"><div className="flex items-center gap-2"><span className={`size-2 rounded-full ${check.status === "pass" ? "bg-emerald-500" : check.status === "warning" ? "bg-amber-500" : "bg-red-500"}`}/><span className="text-sm font-medium">{check.label}</span></div><p className="mt-2 text-xs leading-5 text-muted-foreground">{check.detail}</p></div>)}</div></div>
            </div>
          </CardContent>
        </Card>
      ) : null}

      {signals.length ? (
        <Card>
          <CardHeader><CardTitle className="flex items-center gap-2"><WandSparkles className="size-5"/>Proactive management signals</CardTitle><CardDescription>Deterministic checks from your latest monthly finance data. These are calculated before the AI narrative is generated.</CardDescription></CardHeader>
          <CardContent className="grid gap-3 lg:grid-cols-3">
            {signals.slice(0,3).map((signal,index)=><div key={`${signal.title}-${index}`} className="rounded-2xl border p-4"><div className="flex items-center justify-between gap-3"><p className="font-semibold">{signal.title}</p><span className={`rounded-full px-2 py-1 text-[11px] font-semibold ${signal.severity === "high" ? "bg-red-100 text-red-800" : signal.severity === "positive" ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}`}>{signal.severity}</span></div><p className="mt-3 text-sm text-muted-foreground">{signal.evidence}</p><p className="mt-3 text-sm"><b>Management action:</b> {signal.action}</p></div>)}
          </CardContent>
        </Card>
      ) : null}

      {executiveBrief ? (
        <Card className="border-primary/20 bg-primary/5">
          <CardHeader><CardTitle className="flex items-center gap-2"><Bot className="size-5"/>AI Executive Briefing</CardTitle><CardDescription>Proactive CFO view combining your loaded finance data with relevant external context when available.</CardDescription></CardHeader>
          <CardContent><div className="whitespace-pre-wrap text-sm leading-7">{executiveBrief.answer}</div>{executiveBrief.sources?.length ? <div className="mt-4 text-xs text-muted-foreground">External sources used: {executiveBrief.sources.length}</div>:null}</CardContent>
        </Card>
      ) : null}

      <div className="grid gap-6 xl:grid-cols-[1.4fr_1fr]">
        <Card>
          <CardHeader><CardTitle>Profitability snapshot</CardTitle><CardDescription>Current report totals</CardDescription></CardHeader>
          <CardContent className="space-y-4">
            {[
              ["Revenue", pnl?.revenue],
              ["Cost of sales", pnl?.cost_of_sales],
              ["Gross profit", pnl?.gross_profit],
              ["Operating expenses", pnl?.operating_expenses],
              ["Net profit", pnl?.net_profit],
            ].map(([label, value]) => <div key={String(label)} className="flex items-center justify-between border-b pb-3 last:border-0"><span className="text-sm text-muted-foreground">{label}</span><span className="font-semibold tabular-nums">{formatMoney(value as string | number | undefined, currency)}</span></div>)}
          </CardContent>
        </Card>
        <Card>
          <CardHeader><CardTitle>Data health</CardTitle><CardDescription>{dataHealth?.overall_status === "healthy" ? "All core finance checks passed" : "Ledger and mapping status"}</CardDescription></CardHeader>
          <CardContent className="space-y-4">
            <Status label="Trial balance difference" value={formatMoney(dataHealth?.trial_balance_difference ?? trialBalance?.difference, currency)} good={Boolean(dataHealth?.is_trial_balance_balanced)} />
            <Status label="Balance sheet difference" value={formatMoney(dataHealth?.balance_sheet_difference, currency)} good={Boolean(dataHealth?.is_balance_sheet_balanced)} />
            <Status label="Mapped accounts" value={`${dataHealth?.mapped_account_count ?? 0}/${dataHealth?.account_count ?? 0}`} good={Boolean(dataHealth?.is_mapping_complete && dataHealth?.account_count)} />
            <Status label="Invalid transactions" value={String(dataHealth?.invalid_transaction_count ?? 0)} good={(dataHealth?.invalid_transaction_count ?? 0) === 0} />
            <Status label="Potential duplicates" value={String(dataHealth?.duplicate_candidate_count ?? 0)} good={(dataHealth?.duplicate_candidate_count ?? 0) === 0} />
            <Link href="/dashboard/reports" className={buttonVariants({ variant: "outline", className: "w-full" })}>Open financial reports<ArrowRight className="size-4" /></Link>
          </CardContent>
        </Card>
      </div>
    </div>
  );
}

function MetricCard({ title, description, value, icon: Icon }: { title: string; description: string; value: string; icon: typeof Building2 }) {
  return <Card><CardHeader><div className="flex items-center justify-between"><CardTitle className="text-base">{title}</CardTitle><Icon className="size-5 text-muted-foreground" /></div><CardDescription>{description}</CardDescription></CardHeader><CardContent><p className="truncate text-2xl font-semibold">{value}</p></CardContent></Card>;
}

function Status({ label, value, good }: { label: string; value: string; good: boolean }) {
  return <div className="flex items-center justify-between gap-3"><span className="text-sm text-muted-foreground">{label}</span><span className={`rounded-full px-2 py-1 text-xs font-medium ${good ? "bg-emerald-100 text-emerald-800" : "bg-amber-100 text-amber-800"}`}>{value}</span></div>;
}
