"use client";

import { useEffect, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  BarChart3,
  Loader2,
  RefreshCw,
  Users,
  WalletCards,
} from "lucide-react";
import Link from "next/link";

import { Alert, AlertDescription } from "@/components/ui/alert";
import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { formatMoney, formatPercent, toNumber } from "@/lib/finance-format";
import { getApiErrorMessage } from "@/lib/api";
import { analyticsService } from "@/services/analytics-service";
import type { AnalyticsOverview, WorkingCapitalSummary } from "@/types/analytics";

export default function AnalyticsPage() {
  const [data, setData] = useState<AnalyticsOverview | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setLoading(true);
    setError("");
    try {
      setData(await analyticsService.getOverview());
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => {
    void load();
  }, []);

  if (loading) {
    return (
      <div className="flex min-h-[500px] items-center justify-center gap-3 text-muted-foreground">
        <Loader2 className="size-5 animate-spin" />
        Building business analytics...
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl space-y-7">
      <div className="flex flex-col gap-4 sm:flex-row sm:items-end sm:justify-between">
        <div className="animate-rise">
          <p className="text-sm font-medium text-muted-foreground">Reporting & analytics</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">Business Analytics</h1>
          <p className="mt-2 max-w-3xl text-muted-foreground">
            Trends, branches, customer collections and vendor payment exposure in one view.
          </p>
        </div>
        <Button variant="outline" onClick={() => void load()}>
          <RefreshCw className="size-4" />
          Refresh
        </Button>
      </div>

      {error ? (
        <Alert variant="destructive">
          <AlertDescription>{error}</AlertDescription>
        </Alert>
      ) : null}

      <div className="grid gap-4 md:grid-cols-2 xl:grid-cols-4">
        <Metric
          label="Latest revenue"
          value={data?.monthly_actuals?.at(-1)?.revenue}
          note="Latest available month"
        />
        <Metric
          label="Latest net profit"
          value={data?.monthly_actuals?.at(-1)?.net_profit}
          note="Latest available month"
        />
        <Metric
          label="Accounts receivable"
          value={data?.ar_summary?.total_outstanding}
          note={data?.ar_summary ? `${formatPercent(data.ar_summary.overdue_percent)} overdue` : "Upload AR ageing"}
        />
        <Metric
          label="Accounts payable"
          value={data?.ap_summary?.total_outstanding}
          note={data?.ap_summary ? `${formatPercent(data.ap_summary.overdue_percent)} overdue` : "Upload AP ageing"}
        />
      </div>

      <div className="grid gap-5 xl:grid-cols-[1.1fr_0.9fr]">
        <Card className="overflow-hidden">
          <CardHeader>
            <CardTitle className="flex items-center gap-2">
              <BarChart3 className="size-5" />
              Monthly performance
            </CardTitle>
            <CardDescription>Revenue and net profit across saved monthly actuals.</CardDescription>
          </CardHeader>
          <CardContent>
            {!data?.monthly_actuals?.length ? (
              <Empty text="Upload and map a GL to activate monthly analytics." />
            ) : (
              <div className="space-y-4">
                {data.monthly_actuals.slice(-12).map((row, index, values) => {
                  const maxRevenue = Math.max(...values.map((item) => Math.abs(toNumber(item.revenue))), 1);
                  const revenueWidth = Math.max(4, Math.abs(toNumber(row.revenue)) / maxRevenue * 100);
                  const profitWidth = Math.max(2, Math.min(100, Math.abs(toNumber(row.net_profit)) / maxRevenue * 100));

                  return (
                    <div key={String(row.month)} className="grid grid-cols-[92px_1fr_110px] items-center gap-3">
                      <p className="text-sm text-muted-foreground">{String(row.month).slice(0, 7)}</p>
                      <div className="space-y-1.5">
                        <div className="h-3 rounded-full bg-muted">
                          <div className="h-3 rounded-full bg-indigo-500 transition-all" style={{ width: `${revenueWidth}%` }} />
                        </div>
                        <div className="h-2 rounded-full bg-muted">
                          <div className="h-2 rounded-full bg-emerald-500 transition-all" style={{ width: `${profitWidth}%` }} />
                        </div>
                      </div>
                      <p className="text-right text-sm font-semibold">{formatMoney(row.revenue)}</p>
                    </div>
                  );
                })}
              </div>
            )}
          </CardContent>
        </Card>

        <Card>
          <CardHeader>
            <CardTitle>Management insights</CardTitle>
            <CardDescription>Rule-based insights grounded in current uploads.</CardDescription>
          </CardHeader>
          <CardContent className="space-y-3">
            {(data?.insights ?? []).map((insight, index) => (
              <div key={`${insight}-${index}`} className="flex gap-3 rounded-2xl border bg-muted/20 p-4">
                <AlertTriangle className="mt-0.5 size-5 shrink-0 text-amber-600" />
                <p className="text-sm leading-6">{insight}</p>
              </div>
            ))}
          </CardContent>
        </Card>
      </div>

      <div className="grid gap-5 lg:grid-cols-2">
        <WorkingCapitalCard title="Customer receivables" icon={Users} summary={data?.ar_summary ?? null} />
        <WorkingCapitalCard title="Vendor payables" icon={WalletCards} summary={data?.ap_summary ?? null} />
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Customer profitability — data requirement</CardTitle>
          <CardDescription>
            AR ageing measures exposure and collection behaviour, but it does not contain customer revenue or direct costs.
          </CardDescription>
        </CardHeader>
        <CardContent className="flex flex-col gap-4 rounded-b-xl bg-indigo-50/50 p-6 dark:bg-indigo-950/20 sm:flex-row sm:items-center sm:justify-between">
          <p className="max-w-3xl text-sm leading-6 text-muted-foreground">
            To calculate customer profitability, the next import needs customer-level invoice revenue,
            credit notes and attributable cost or gross-margin data. Vendor spend analysis similarly
            needs a purchase invoice history, not only the current AP balance.
          </p>
          <Link href="/dashboard/import-center">
            <Button>
              Open Import Centre
              <ArrowRight className="size-4" />
            </Button>
          </Link>
        </CardContent>
      </Card>
    </div>
  );
}

function Metric({
  label,
  value,
  note,
}: {
  label: string;
  value?: string | number;
  note: string;
}) {
  return (
    <Card className="animate-card-in">
      <CardHeader className="pb-2">
        <CardDescription>{label}</CardDescription>
        <CardTitle className="text-2xl">{value == null ? "—" : formatMoney(value)}</CardTitle>
        <p className="text-xs text-muted-foreground">{note}</p>
      </CardHeader>
    </Card>
  );
}

function WorkingCapitalCard({
  title,
  icon: Icon,
  summary,
}: {
  title: string;
  icon: typeof Users;
  summary: WorkingCapitalSummary | null;
}) {
  return (
    <Card>
      <CardHeader>
        <CardTitle className="flex items-center gap-2">
          <Icon className="size-5" />
          {title}
        </CardTitle>
      </CardHeader>
      <CardContent>
        {!summary ? (
          <Empty text="Upload the relevant ageing file in the Import Centre." />
        ) : (
          <div className="space-y-5">
            <div className="grid grid-cols-3 gap-3">
              <SmallMetric label="Outstanding" value={formatMoney(summary.total_outstanding)} />
              <SmallMetric label="Overdue" value={formatMoney(summary.overdue_amount)} />
              <SmallMetric label="Overdue %" value={formatPercent(summary.overdue_percent)} />
            </div>

            <div>
              <p className="mb-3 text-sm font-semibold">Ageing profile</p>
              <div className="space-y-2">
                {summary.buckets.map((bucket) => (
                  <div key={bucket.bucket} className="flex items-center justify-between rounded-xl bg-muted/30 px-4 py-3 text-sm">
                    <span>{bucket.bucket}</span>
                    <span className="font-semibold">{formatMoney(bucket.amount)}</span>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <p className="mb-3 text-sm font-semibold">Largest exposures</p>
              <div className="space-y-2">
                {summary.top_parties.slice(0, 5).map((party) => (
                  <div key={party.party_name} className="flex items-center justify-between border-b py-2 text-sm last:border-0">
                    <span>{party.party_name}</span>
                    <span className="font-semibold">{formatMoney(party.outstanding_amount)}</span>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}
      </CardContent>
    </Card>
  );
}

function SmallMetric({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded-xl border bg-muted/20 p-3">
      <p className="text-xs text-muted-foreground">{label}</p>
      <p className="mt-1 font-semibold">{value}</p>
    </div>
  );
}

function Empty({ text }: { text: string }) {
  return (
    <div className="flex min-h-40 items-center justify-center rounded-2xl border border-dashed text-center text-sm text-muted-foreground">
      {text}
    </div>
  );
}
