"use client";

import { useEffect, useState } from "react";
import Link from "next/link";
import { useRouter } from "next/navigation";
import { ArrowRight, Building2, CheckCircle2, CircleDollarSign, Loader2, TrendingUp, Upload, WalletCards, WandSparkles } from "lucide-react";

import { Button, buttonVariants } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getApiErrorMessage } from "@/lib/api";
import { formatMoney, toNumber } from "@/lib/finance-format";
import { authService } from "@/services/auth-service";
import { financeService } from "@/services/finance-service";
import type { Company } from "@/types/auth";
import type { BalanceSheet, ProfitAndLoss, Ratio, TrialBalance } from "@/types/finance";

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

  useEffect(() => {
    async function load() {
      if (!authService.hasAccessToken()) {
        router.replace("/login");
        return;
      }
      try {
        const currentCompany = await authService.getCurrentCompany();
        setCompany(currentCompany);
        const [profitLoss, bs, tb, ratios, suggestions] = await Promise.all([
          financeService.getProfitAndLoss(),
          financeService.getBalanceSheet(),
          financeService.getTrialBalance(),
          financeService.getKpis(),
          financeService.getMappingSuggestions(),
        ]);
        setPnl(profitLoss);
        setBalanceSheet(bs);
        setTrialBalance(tb);
        setKpis(ratios);
        setUnmappedCount(suggestions.length);
      } catch (loadError) {
        setError(getApiErrorMessage(loadError));
      } finally {
        setIsLoading(false);
      }
    }
    void load();
  }, [router]);

  if (isLoading) return <div className="flex min-h-[60vh] items-center justify-center gap-3 text-muted-foreground"><Loader2 className="size-5 animate-spin" />Loading your dashboard...</div>;

  const companyName = company?.trading_name ?? company?.legal_name ?? "Your company";
  const currency = company?.currency_code ?? "AUD";
  const cashRatio = kpis.find((item) => item.name.toLowerCase().includes("cash"));
  const hasFinancials = Boolean(trialBalance?.lines.length);
  const mappedReady = unmappedCount === 0 && hasFinancials;

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div><p className="text-sm font-medium text-muted-foreground">Dashboard</p><h1 className="mt-1 text-3xl font-semibold tracking-tight">Welcome back</h1><p className="mt-2 text-muted-foreground">Financial overview for <span className="font-medium text-foreground">{companyName}</span>.</p></div>
        <Link href="/dashboard/uploads" className={buttonVariants()}><Upload className="size-4" />Upload general ledger</Link>
      </div>

      {error ? <Card className="border-destructive/30"><CardContent className="py-4 text-sm text-destructive">{error}</CardContent></Card> : null}

      <div className="grid gap-5 md:grid-cols-2 xl:grid-cols-4">
        <MetricCard title="Current company" description={`${company?.country_code ?? ""} · ${currency}`} value={companyName} icon={Building2} />
        <MetricCard title="Revenue" description="Current reporting data" value={formatMoney(pnl?.revenue, currency)} icon={TrendingUp} />
        <MetricCard title="Net profit" description="After tax" value={formatMoney(pnl?.net_profit, currency)} icon={CircleDollarSign} />
        <MetricCard title="Total assets" description={cashRatio ? `${cashRatio.name}: ${cashRatio.value ?? "—"}` : "Balance sheet"} value={formatMoney(balanceSheet?.total_assets, currency)} icon={WalletCards} />
      </div>

      {!hasFinancials ? (
        <Card><CardContent className="flex flex-col items-start gap-4 py-8 sm:flex-row sm:items-center sm:justify-between"><div><p className="font-semibold">Start with a general ledger upload</p><p className="mt-1 text-sm text-muted-foreground">Upload a CSV to validate transactions and generate your trial balance.</p></div><Link href="/dashboard/uploads" className={buttonVariants()}>Upload now<ArrowRight className="size-4" /></Link></CardContent></Card>
      ) : !mappedReady ? (
        <Card className="border-amber-200 bg-amber-50/40"><CardContent className="flex flex-col items-start gap-4 py-6 sm:flex-row sm:items-center sm:justify-between"><div className="flex gap-3"><WandSparkles className="mt-0.5 size-5 text-amber-700" /><div><p className="font-semibold">Review {unmappedCount} account mapping{unmappedCount === 1 ? "" : "s"}</p><p className="mt-1 text-sm text-muted-foreground">Your trial balance is ready. Approve mappings to populate P&L, balance sheet and KPIs.</p></div></div><Link href="/dashboard/mapping" className={buttonVariants()}>Review mappings<ArrowRight className="size-4" /></Link></CardContent></Card>
      ) : (
        <Card className="border-emerald-200 bg-emerald-50/40"><CardContent className="flex items-center gap-3 py-5"><CheckCircle2 className="size-5 text-emerald-700" /><div><p className="font-semibold">Finance data is ready</p><p className="text-sm text-muted-foreground">Ledger is balanced and all accounts are mapped.</p></div></CardContent></Card>
      )}

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
          <CardHeader><CardTitle>Data health</CardTitle><CardDescription>Ledger and mapping status</CardDescription></CardHeader>
          <CardContent className="space-y-4">
            <Status label="Trial balance difference" value={formatMoney(trialBalance?.difference, currency)} good={Math.abs(toNumber(trialBalance?.difference)) < 0.01} />
            <Status label="Ledger accounts" value={String(trialBalance?.lines.length ?? 0)} good={Boolean(trialBalance?.lines.length)} />
            <Status label="Unmapped accounts" value={String(unmappedCount)} good={unmappedCount === 0} />
            <Status label="KPI indicators" value={String(kpis.length)} good={kpis.length > 0} />
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
