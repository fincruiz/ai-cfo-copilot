"use client";

import { useMemo, useState } from "react";
import { ArrowDownRight, ArrowUpRight, Minus, TrendingUp } from "lucide-react";

import { HelpTip } from "@/components/ui/help-tip";
import { formatMoney, formatPercent, toNumber } from "@/lib/finance-format";
import type { AnalyticsOverview } from "@/types/analytics";

type Period = 3 | 6 | 12 | "all";
type MonthlyRow = Record<string, string | number>;

function pctChange(current: number, prior: number): number | null {
  if (Math.abs(prior) < 0.000001) return null;
  return ((current - prior) / Math.abs(prior)) * 100;
}

function monthLabel(value: unknown): string {
  if (!value) return "—";
  const raw = String(value);
  const parsed = new Date(raw);
  if (!Number.isNaN(parsed.getTime())) return parsed.toLocaleDateString(undefined, { month: "short", year: "2-digit" });
  return raw.slice(0, 10);
}

export function ManagementPerformanceBoard({ overview, currency }: { overview: AnalyticsOverview | null; currency: string }) {
  const [period, setPeriod] = useState<Period>(12);
  const monthly = useMemo(() => {
    const rows = overview?.monthly_actuals ?? [];
    if (period === "all") return rows;
    return rows.slice(-period);
  }, [overview, period]);

  if (!monthly.length) return null;

  const latest = monthly[monthly.length - 1] as MonthlyRow;
  const previous = monthly.length > 1 ? (monthly[monthly.length - 2] as MonthlyRow) : null;
  const latestRevenue = toNumber(latest.revenue);
  const latestProfit = toNumber(latest.net_profit);
  const latestGrossProfit = toNumber(latest.gross_profit);
  const latestOpex = toNumber(latest.operating_expenses);
  const grossMargin = latestRevenue ? (latestGrossProfit / latestRevenue) * 100 : 0;
  const opexRatio = latestRevenue ? (latestOpex / latestRevenue) * 100 : 0;
  const overdueAr = toNumber(overview?.ar_summary?.overdue_amount);
  const overduePct = toNumber(overview?.ar_summary?.overdue_percent);

  const previousRevenue = previous ? toNumber(previous.revenue) : 0;
  const previousProfit = previous ? toNumber(previous.net_profit) : 0;
  const revenueChange = previous ? pctChange(latestRevenue, previousRevenue) : null;
  const profitChange = previous ? pctChange(latestProfit, previousProfit) : null;

  return (
    <section className="overflow-hidden rounded-[28px] border bg-background shadow-sm">
      <div className="flex flex-col gap-4 border-b bg-gradient-to-r from-slate-50 via-background to-indigo-50/60 p-5 dark:from-slate-950 dark:to-indigo-950/20 sm:flex-row sm:items-center sm:justify-between sm:p-6">
        <div>
          <div className="flex items-center gap-2">
            <TrendingUp className="size-5 text-indigo-600" />
            <h2 className="text-xl font-semibold">Business performance</h2>
            <HelpTip title="Business performance" text="A management view of recent revenue, profit, margin and working-capital movements. Use the detailed finance pages when you need transaction or report evidence." side="bottom" />
          </div>
          <p className="mt-1 text-sm text-muted-foreground">See direction first, then drill into the accounting detail only when needed.</p>
        </div>
        <div className="flex rounded-full border bg-background p-1 shadow-sm">
          {([3, 6, 12, "all"] as Period[]).map((item) => (
            <button key={String(item)} type="button" onClick={() => setPeriod(item)} className={`rounded-full px-3 py-1.5 text-xs font-semibold transition ${period === item ? "bg-slate-950 text-white dark:bg-white dark:text-slate-950" : "text-muted-foreground hover:bg-muted"}`}>
              {item === "all" ? "All" : `${item}M`}
            </button>
          ))}
        </div>
      </div>

      <div className="grid gap-px bg-border lg:grid-cols-[1.5fr_.5fr]">
        <div className="bg-background p-5 sm:p-6">
          <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
            <TrendTile label="Latest revenue" value={formatMoney(latestRevenue, currency)} change={revenueChange} series={monthly.map((row) => toNumber(row.revenue))} />
            <TrendTile label="Latest net profit" value={formatMoney(latestProfit, currency)} change={profitChange} series={monthly.map((row) => toNumber(row.net_profit))} />
            <TrendTile label="Gross margin" value={formatPercent(grossMargin)} change={null} series={monthly.map((row) => { const rev = toNumber(row.revenue); return rev ? (toNumber(row.gross_profit) / rev) * 100 : 0; })} />
            <TrendTile label="Overdue receivables" value={overview?.ar_summary ? formatMoney(overdueAr, currency) : "Not loaded"} secondary={overview?.ar_summary ? `${formatPercent(overduePct)} of AR` : "Add AR ageing for this view"} change={null} series={[]} />
          </div>

          <div className="mt-5 rounded-2xl border bg-muted/15 p-4">
            <div className="mb-4 flex flex-wrap items-center justify-between gap-3">
              <div><p className="text-sm font-semibold">Revenue and profit trajectory</p><p className="text-xs text-muted-foreground">Recent reporting periods · hover points for exact values</p></div>
              <div className="flex gap-3 text-xs text-muted-foreground"><span className="flex items-center gap-1.5"><span className="size-2 rounded-full bg-indigo-600"/>Revenue</span><span className="flex items-center gap-1.5"><span className="size-2 rounded-full bg-emerald-500"/>Net profit</span></div>
            </div>
            <PerformanceTrend rows={monthly} currency={currency} />
          </div>
        </div>

        <aside className="bg-background p-5 sm:p-6">
          <p className="text-xs font-semibold uppercase tracking-[.16em] text-muted-foreground">Latest period</p>
          <p className="mt-1 text-lg font-semibold">{monthLabel(latest.month)}</p>
          <div className="mt-5 space-y-3">
            <Driver label="Gross margin" value={formatPercent(grossMargin)} detail="Revenue retained after direct costs" tone={grossMargin >= 30 ? "good" : grossMargin >= 15 ? "watch" : "risk"} />
            <Driver label="Operating expense ratio" value={formatPercent(opexRatio)} detail="Operating expenses as a share of revenue" tone={opexRatio <= 35 ? "good" : opexRatio <= 50 ? "watch" : "risk"} />
            {overview?.ar_summary ? <Driver label="Overdue AR" value={formatPercent(overduePct)} detail={`${formatMoney(overdueAr, currency)} currently overdue`} tone={overduePct <= 20 ? "good" : overduePct <= 40 ? "watch" : "risk"} /> : null}
          </div>
          {overview?.insights?.length ? <div className="mt-5 rounded-2xl bg-slate-950 p-4 text-white"><p className="text-xs font-semibold uppercase tracking-[.14em] text-slate-400">FinCruiz observation</p><p className="mt-2 text-sm leading-6">{overview.insights[0]}</p></div> : null}
        </aside>
      </div>
    </section>
  );
}

function TrendTile({ label, value, change, series, secondary }: { label: string; value: string; change: number | null; series: number[]; secondary?: string }) {
  const positive = change !== null && change > 0;
  const negative = change !== null && change < 0;
  return <div className="rounded-2xl border bg-background p-4 transition hover:-translate-y-0.5 hover:shadow-md">
    <p className="text-xs font-medium text-muted-foreground">{label}</p>
    <div className="mt-2 flex items-end justify-between gap-3"><p className="text-xl font-semibold tabular-nums">{value}</p>{series.length > 1 ? <Sparkline values={series} /> : null}</div>
    <div className="mt-2 min-h-5 text-xs">{change !== null ? <span className={`inline-flex items-center gap-1 font-semibold ${positive ? "text-emerald-600" : negative ? "text-red-600" : "text-muted-foreground"}`}>{positive ? <ArrowUpRight className="size-3.5"/> : negative ? <ArrowDownRight className="size-3.5"/> : <Minus className="size-3.5"/>}{Math.abs(change).toFixed(2)}% vs prior period</span> : secondary ? <span className="text-muted-foreground">{secondary}</span> : <span className="text-muted-foreground">Current period</span>}</div>
  </div>;
}

function Sparkline({ values }: { values: number[] }) {
  const clean = values.map(Number).filter(Number.isFinite);
  if (clean.length < 2) return null;
  const min = Math.min(...clean), max = Math.max(...clean), range = Math.max(max - min, 1);
  const points = clean.map((value, index) => `${(index / (clean.length - 1)) * 76 + 2},${28 - ((value - min) / range) * 24}`).join(" ");
  return <svg viewBox="0 0 80 32" className="h-8 w-20 overflow-visible" aria-hidden="true"><polyline points={points} fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round" className="text-indigo-500"/></svg>;
}

function PerformanceTrend({ rows, currency }: { rows: MonthlyRow[]; currency: string }) {
  const revenue = rows.map((row) => toNumber(row.revenue));
  const profit = rows.map((row) => toNumber(row.net_profit));
  const all = [...revenue, ...profit];
  const min = Math.min(...all, 0), max = Math.max(...all, 1), range = Math.max(max - min, 1);
  const width = 760, height = 250, left = 28, right = 18, top = 18, bottom = 34;
  const x = (index: number) => left + (index * (width - left - right)) / Math.max(rows.length - 1, 1);
  const y = (value: number) => height - bottom - ((value - min) / range) * (height - top - bottom);
  const line = (values: number[]) => values.map((value, index) => `${x(index)},${y(value)}`).join(" ");
  const display = (value: number) => new Intl.NumberFormat(undefined, { style: "currency", currency, notation: "compact", maximumFractionDigits: 2 }).format(value);
  return <svg viewBox={`0 0 ${width} ${height}`} className="h-64 w-full overflow-visible">
    {[0,1,2,3].map((index) => <line key={index} x1={left} x2={width-right} y1={top + index * (height-top-bottom)/3} y2={top + index * (height-top-bottom)/3} stroke="currentColor" opacity=".07" />)}
    <polyline points={line(revenue)} fill="none" stroke="#4f46e5" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
    <polyline points={line(profit)} fill="none" stroke="#10b981" strokeWidth="3" strokeLinecap="round" strokeLinejoin="round" />
    {rows.map((row, index) => <g key={`${row.month}-${index}`}>
      <circle cx={x(index)} cy={y(revenue[index])} r="3.5" fill="#4f46e5"><title>{monthLabel(row.month)} · Revenue: {display(revenue[index])}</title></circle>
      <circle cx={x(index)} cy={y(profit[index])} r="3.5" fill="#10b981"><title>{monthLabel(row.month)} · Net profit: {display(profit[index])}</title></circle>
      {(index === 0 || index === rows.length - 1 || index % Math.max(Math.ceil(rows.length / 6), 1) === 0) ? <text x={x(index)} y={height-8} textAnchor="middle" fontSize="10" fill="currentColor" opacity=".55">{monthLabel(row.month).slice(0, 8)}</text> : null}
    </g>)}
  </svg>;
}

function Driver({ label, value, detail, tone }: { label: string; value: string; detail: string; tone: "good" | "watch" | "risk" }) {
  const dot = tone === "good" ? "bg-emerald-500" : tone === "watch" ? "bg-amber-500" : "bg-red-500";
  return <div className="rounded-2xl border p-3.5"><div className="flex items-center justify-between gap-3"><span className="flex items-center gap-2 text-sm font-medium"><span className={`size-2 rounded-full ${dot}`}/>{label}</span><b className="tabular-nums">{value}</b></div><p className="mt-2 text-xs leading-5 text-muted-foreground">{detail}</p></div>;
}
