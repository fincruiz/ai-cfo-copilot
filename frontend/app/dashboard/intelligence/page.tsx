"use client";

import { useEffect, useMemo, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  BrainCircuit,
  CheckCircle2,
  ChevronRight,
  CircleDollarSign,
  Database,
  Lightbulb,
  Plus,
  RefreshCw,
  Send,
  ShieldCheck,
  Sparkles,
  TrendingDown,
  TrendingUp,
} from "lucide-react";

import { Button } from "@/components/ui/button";
import { getApiErrorMessage } from "@/lib/api";
import { intelligenceService } from "@/services/intelligence-service";
import type {
  BrainOverview,
  IntelligenceMetric,
  IntelligencePriority,
  MonthlyTrendPoint,
} from "@/types/integrations";

type AskResult = {
  answer: string;
  sources?: Array<{ title: string; url: string }>;
  action?: { label: string; route: string } | null;
};

function compactMoney(value: number, currency: string) {
  try {
    return new Intl.NumberFormat("en", {
      style: "currency",
      currency,
      notation: Math.abs(value) >= 1_000_000 ? "compact" : "standard",
      maximumFractionDigits: Math.abs(value) >= 1_000_000 ? 1 : 0,
    }).format(value);
  } catch {
    return `${currency} ${Math.round(value).toLocaleString()}`;
  }
}

function formatMetric(metric: IntelligenceMetric, currency: string) {
  if (metric.value === null || metric.value === undefined) return "—";
  if (metric.format === "currency") return compactMoney(metric.value, currency);
  if (metric.format === "percent") return `${metric.value.toFixed(1)}%`;
  if (metric.format === "score") return `${Math.round(metric.value)}/100`;
  return metric.value.toLocaleString();
}

function changeText(metric: IntelligenceMetric) {
  if (metric.change === null || metric.change === undefined) return null;
  if (metric.change_unit === "points") return `${metric.change >= 0 ? "+" : ""}${metric.change.toFixed(1)} pts`;
  if (metric.change_unit === "of_ar") return `${metric.change.toFixed(1)}% of AR`;
  return `${metric.change >= 0 ? "+" : ""}${metric.change.toFixed(1)}%`;
}

function priorityStyle(level: string) {
  if (level === "critical") {
    return {
      shell: "border-red-200 bg-red-50/70",
      badge: "bg-red-100 text-red-700",
      icon: <AlertTriangle className="size-4 text-red-600" />,
      label: "Critical",
    };
  }
  if (level === "attention") {
    return {
      shell: "border-amber-200 bg-amber-50/70",
      badge: "bg-amber-100 text-amber-700",
      icon: <AlertTriangle className="size-4 text-amber-600" />,
      label: "Attention",
    };
  }
  if (level === "positive") {
    return {
      shell: "border-emerald-200 bg-emerald-50/70",
      badge: "bg-emerald-100 text-emerald-700",
      icon: <CheckCircle2 className="size-4 text-emerald-600" />,
      label: "Positive",
    };
  }
  return {
    shell: "border-slate-200 bg-slate-50/70",
    badge: "bg-slate-100 text-slate-700",
    icon: <Lightbulb className="size-4 text-slate-600" />,
    label: "Monitor",
  };
}

function MiniTrendChart({ data }: { data: MonthlyTrendPoint[] }) {
  if (!data.length) {
    return (
      <div className="flex h-52 items-center justify-center rounded-2xl bg-muted/30 text-sm text-muted-foreground">
        Add at least two reporting periods to see the management trend.
      </div>
    );
  }

  const width = 760;
  const height = 220;
  const pad = 18;
  const values = data.flatMap((x) => [x.revenue, x.net_profit]);
  const min = Math.min(0, ...values);
  const max = Math.max(1, ...values);
  const range = max - min || 1;
  const x = (i: number) => pad + (i * (width - pad * 2)) / Math.max(data.length - 1, 1);
  const y = (v: number) => height - pad - ((v - min) / range) * (height - pad * 2);
  const revenue = data.map((p, i) => `${x(i)},${y(p.revenue)}`).join(" ");
  const profit = data.map((p, i) => `${x(i)},${y(p.net_profit)}`).join(" ");

  return (
    <div>
      <div className="mb-4 flex items-center gap-5 text-xs text-muted-foreground">
        <span className="flex items-center gap-2"><span className="size-2.5 rounded-full bg-slate-900" />Revenue</span>
        <span className="flex items-center gap-2"><span className="size-2.5 rounded-full bg-emerald-500" />Net profit</span>
      </div>
      <svg viewBox={`0 0 ${width} ${height}`} className="h-52 w-full overflow-visible" role="img" aria-label="Revenue and profit trend">
        <line x1={pad} x2={width - pad} y1={height - pad} y2={height - pad} stroke="currentColor" opacity="0.08" />
        <polyline points={revenue} fill="none" stroke="currentColor" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round" />
        <polyline points={profit} fill="none" stroke="#10b981" strokeWidth="4" strokeLinecap="round" strokeLinejoin="round" />
        {data.map((p, i) => (
          <g key={`${p.month}-${i}`}>
            <circle cx={x(i)} cy={y(p.revenue)} r="4" fill="currentColor" />
            <circle cx={x(i)} cy={y(p.net_profit)} r="4" fill="#10b981" />
          </g>
        ))}
      </svg>
      <div className="mt-2 grid grid-cols-4 gap-2 text-[11px] text-muted-foreground md:grid-cols-6">
        {data.map((p, i) => (
          <span key={p.month} className={i < data.length - 6 ? "hidden md:block" : "block"}>
            {new Date(p.month).toLocaleDateString(undefined, { month: "short", year: "2-digit" })}
          </span>
        ))}
      </div>
    </div>
  );
}

export default function IntelligencePage() {
  const [data, setData] = useState<BrainOverview | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [title, setTitle] = useState("");
  const [content, setContent] = useState("");
  const [saving, setSaving] = useState(false);
  const [question, setQuestion] = useState("");
  const [asking, setAsking] = useState(false);
  const [answer, setAnswer] = useState<AskResult | null>(null);

  const load = async () => {
    setLoading(true);
    setError("");
    try {
      setData(await intelligenceService.overview());
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void load();
  }, []);

  const remember = async () => {
    if (!title.trim() || !content.trim()) return;
    setSaving(true);
    setError("");
    try {
      await intelligenceService.addMemory({ title, content });
      setTitle("");
      setContent("");
      await load();
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setSaving(false);
    }
  };

  const ask = async (prompt?: string) => {
    const q = (prompt ?? question).trim();
    if (!q) return;
    setQuestion(q);
    setAsking(true);
    setAnswer(null);
    setError("");
    try {
      const result = await intelligenceService.ask(q, true);
      setAnswer(result);
    } catch (e) {
      setError(getApiErrorMessage(e));
    } finally {
      setAsking(false);
    }
  };

  const currency = data?.company.currency || "AUD";
  const connected = data?.connections.filter((x) => x.status === "connected").length ?? 0;
  const totalRecords = data?.source_counts.reduce((sum, x) => sum + Number(x.count || 0), 0) ?? 0;
  const assurance = data?.financial_assurance;

  const topPriorities = useMemo(() => data?.priorities?.slice(0, 5) ?? [], [data]);

  if (loading) {
    return (
      <div className="mx-auto max-w-7xl space-y-6 p-6 lg:p-10">
        <div className="h-40 animate-pulse rounded-[2rem] bg-muted/50" />
        <div className="grid gap-4 md:grid-cols-3"><div className="h-32 animate-pulse rounded-3xl bg-muted/40" /><div className="h-32 animate-pulse rounded-3xl bg-muted/40" /><div className="h-32 animate-pulse rounded-3xl bg-muted/40" /></div>
        <div className="h-80 animate-pulse rounded-3xl bg-muted/40" />
      </div>
    );
  }

  return (
    <div className="mx-auto max-w-7xl space-y-7 p-5 lg:p-10">
      <section className="relative overflow-hidden rounded-[2rem] border bg-gradient-to-br from-slate-950 via-slate-900 to-slate-800 p-7 text-white shadow-xl lg:p-10">
        <div className="pointer-events-none absolute -right-20 -top-28 size-72 rounded-full bg-emerald-400/15 blur-3xl" />
        <div className="pointer-events-none absolute -bottom-32 left-1/3 size-72 rounded-full bg-sky-400/10 blur-3xl" />
        <div className="relative grid gap-7 lg:grid-cols-[1fr_auto] lg:items-end">
          <div>
            <div className="mb-4 inline-flex items-center gap-2 rounded-full border border-white/15 bg-white/10 px-3 py-1.5 text-xs font-medium text-white/85">
              <BrainCircuit className="size-3.5" /> FinCruiz Organizational Brain
            </div>
            <p className="text-sm text-white/60">{data?.company.name || "Your business"} · Management intelligence</p>
            <h1 className="mt-2 max-w-4xl text-3xl font-semibold leading-tight tracking-tight lg:text-5xl">
              {data?.executive_summary.headline}
            </h1>
            <p className="mt-4 max-w-3xl text-sm leading-6 text-white/70 lg:text-base">
              {data?.executive_summary.narrative}
            </p>
          </div>
          <div className="flex flex-wrap gap-2">
            <Button variant="secondary" onClick={() => void ask("What should management focus on this month?")}>
              <Sparkles className="mr-2 size-4" /> Explain this
            </Button>
            <Button variant="outline" className="border-white/20 bg-white/5 text-white hover:bg-white/10 hover:text-white" onClick={() => void load()}>
              <RefreshCw className="mr-2 size-4" /> Refresh
            </Button>
          </div>
        </div>
      </section>

      {error ? (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm text-red-700">
          {error}
        </div>
      ) : null}

      <section>
        <div className="mb-3 flex items-end justify-between gap-4">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Business pulse</p>
            <h2 className="mt-1 text-xl font-semibold">The numbers management should see first</h2>
          </div>
          <p className="hidden text-xs text-muted-foreground md:block">Compared with the previous reporting month where available</p>
        </div>
        <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-5">
          {data?.financial_snapshot.map((metric) => {
            const change = changeText(metric);
            const good = metric.key === "overdue_ar" ? false : Number(metric.change || 0) >= 0;
            return (
              <article key={metric.key} className="rounded-3xl border bg-background p-5 shadow-sm">
                <p className="text-sm text-muted-foreground">{metric.label}</p>
                <p className="mt-2 text-2xl font-semibold tracking-tight">{formatMetric(metric, currency)}</p>
                <div className="mt-3 min-h-9">
                  {change ? (
                    <div className={`inline-flex items-center gap-1 rounded-full px-2 py-1 text-xs font-medium ${good ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"}`}>
                      {good ? <TrendingUp className="size-3" /> : <TrendingDown className="size-3" />}
                      {change}
                    </div>
                  ) : null}
                </div>
                <p className="mt-2 text-xs leading-5 text-muted-foreground">{metric.context}</p>
              </article>
            );
          })}
        </div>
      </section>

      <div className="grid gap-6 xl:grid-cols-[1.25fr_.75fr]">
        <section className="rounded-[2rem] border bg-background p-6 shadow-sm lg:p-7">
          <div className="flex items-center justify-between gap-4">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Trend</p>
              <h2 className="mt-1 text-xl font-semibold">Is the business moving in the right direction?</h2>
            </div>
            <CircleDollarSign className="size-5 text-muted-foreground" />
          </div>
          <div className="mt-6"><MiniTrendChart data={data?.monthly_trends ?? []} /></div>
        </section>

        <section className="rounded-[2rem] border bg-background p-6 shadow-sm lg:p-7">
          <div className="flex items-center justify-between gap-4">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Trust layer</p>
              <h2 className="mt-1 text-xl font-semibold">Can management rely on the data?</h2>
            </div>
            <ShieldCheck className="size-5 text-muted-foreground" />
          </div>
          <div className="mt-6 flex items-end gap-3">
            <p className="text-5xl font-semibold tracking-tight">{assurance?.score ?? 0}</p>
            <div className="pb-1"><p className="text-sm font-medium">/100 · Grade {assurance?.grade ?? "—"}</p><p className="text-xs text-muted-foreground">Financial confidence</p></div>
          </div>
          <div className="mt-5 h-2 overflow-hidden rounded-full bg-muted">
            <div className="h-full rounded-full bg-emerald-500 transition-all" style={{ width: `${Math.max(0, Math.min(100, Number(assurance?.score || 0)))}%` }} />
          </div>
          <div className="mt-5 space-y-2">
            {(assurance?.checks ?? []).slice(0, 5).map((check: any) => (
              <div key={check.key} className="flex items-start gap-2 text-sm">
                {check.status === "pass" ? <CheckCircle2 className="mt-0.5 size-4 text-emerald-600" /> : <AlertTriangle className="mt-0.5 size-4 text-amber-600" />}
                <div><p className="font-medium">{check.label}</p><p className="text-xs text-muted-foreground">{check.detail}</p></div>
              </div>
            ))}
          </div>
        </section>
      </div>

      <section className="rounded-[2rem] border bg-background p-6 shadow-sm lg:p-7">
        <div className="flex flex-col gap-2 md:flex-row md:items-end md:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Management priorities</p>
            <h2 className="mt-1 text-2xl font-semibold">What should management focus on?</h2>
            <p className="mt-1 text-sm text-muted-foreground">FinCruiz ranks material signals, shows the evidence, and suggests the next management action.</p>
          </div>
          <div className="flex gap-2 text-xs">
            <span className="rounded-full bg-red-50 px-2.5 py-1 text-red-700">{data?.executive_summary.critical_count ?? 0} critical</span>
            <span className="rounded-full bg-amber-50 px-2.5 py-1 text-amber-700">{data?.executive_summary.attention_count ?? 0} attention</span>
          </div>
        </div>

        <div className="mt-6 grid gap-3 lg:grid-cols-2">
          {topPriorities.length ? topPriorities.map((priority: IntelligencePriority, index) => {
            const style = priorityStyle(priority.level);
            return (
              <article key={`${priority.title}-${index}`} className={`rounded-2xl border p-5 ${style.shell}`}>
                <div className="flex items-start justify-between gap-3">
                  <div className="flex items-start gap-2.5">
                    <div className="mt-0.5">{style.icon}</div>
                    <div>
                      <div className="flex flex-wrap items-center gap-2">
                        <span className={`rounded-full px-2 py-1 text-[11px] font-semibold uppercase tracking-wide ${style.badge}`}>{style.label}</span>
                        <span className="text-[11px] text-muted-foreground">{priority.source}</span>
                      </div>
                      <h3 className="mt-2 font-semibold">{priority.title}</h3>
                    </div>
                  </div>
                  <span className="text-xs font-semibold text-muted-foreground">#{index + 1}</span>
                </div>
                <div className="mt-4 rounded-xl bg-white/65 p-3 text-sm leading-6">
                  <p><span className="font-semibold">Evidence:</span> {priority.evidence}</p>
                  <p className="mt-2"><span className="font-semibold">Management action:</span> {priority.action}</p>
                </div>
                <button className="mt-4 inline-flex items-center gap-1 text-sm font-semibold" onClick={() => void ask(`Investigate this management signal: ${priority.title}. Evidence: ${priority.evidence}. What should management do next?`)}>
                  Investigate with FinCruiz <ChevronRight className="size-4" />
                </button>
              </article>
            );
          }) : (
            <div className="col-span-full rounded-2xl bg-muted/30 p-8 text-center text-sm text-muted-foreground">
              No management priorities are active yet. Load at least two reporting periods and AR/AP data for deeper signals.
            </div>
          )}
        </div>
      </section>

      <section className="overflow-hidden rounded-[2rem] border bg-slate-950 text-white shadow-xl">
        <div className="grid lg:grid-cols-[.8fr_1.2fr]">
          <div className="border-b border-white/10 p-6 lg:border-b-0 lg:border-r lg:p-8">
            <div className="inline-flex size-11 items-center justify-center rounded-2xl bg-white/10"><Sparkles className="size-5" /></div>
            <h2 className="mt-5 text-2xl font-semibold">Ask FinCruiz</h2>
            <p className="mt-2 text-sm leading-6 text-white/65">Ask in plain English. FinCruiz uses the company data loaded into the workspace and can add current external context when relevant.</p>
            <div className="mt-5 space-y-2">
              {(data?.suggested_questions ?? []).map((q) => (
                <button key={q} onClick={() => void ask(q)} className="flex w-full items-center justify-between rounded-xl border border-white/10 bg-white/5 px-3 py-2.5 text-left text-sm text-white/80 transition hover:bg-white/10">
                  <span>{q}</span><ArrowRight className="size-3.5" />
                </button>
              ))}
            </div>
          </div>
          <div className="p-6 lg:p-8">
            <div className="flex gap-2">
              <input
                value={question}
                onChange={(e) => setQuestion(e.target.value)}
                onKeyDown={(e) => { if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); void ask(); } }}
                placeholder="e.g. Why is profit falling even though revenue is growing?"
                className="min-w-0 flex-1 rounded-xl border border-white/15 bg-white/10 px-4 py-3 text-sm text-white outline-none placeholder:text-white/35 focus:border-white/30"
              />
              <Button variant="secondary" disabled={asking || !question.trim()} onClick={() => void ask()}>
                {asking ? <RefreshCw className="size-4 animate-spin" /> : <Send className="size-4" />}
              </Button>
            </div>
            <div className="mt-5 min-h-52 rounded-2xl border border-white/10 bg-white/[0.04] p-5">
              {asking ? (
                <div className="flex h-40 items-center justify-center gap-2 text-sm text-white/60"><RefreshCw className="size-4 animate-spin" /> FinCruiz is connecting the evidence…</div>
              ) : answer ? (
                <div>
                  <p className="whitespace-pre-wrap text-sm leading-7 text-white/85">{answer.answer}</p>
                  {answer.action ? (
                    <a href={answer.action.route} className="mt-5 inline-flex items-center gap-2 rounded-xl bg-white px-3 py-2 text-sm font-semibold text-slate-950">{answer.action.label}<ArrowRight className="size-4" /></a>
                  ) : null}
                  {answer.sources?.length ? (
                    <div className="mt-5 border-t border-white/10 pt-4"><p className="text-xs font-semibold uppercase tracking-wider text-white/45">External sources</p><div className="mt-2 flex flex-wrap gap-2">{answer.sources.slice(0, 5).map((s) => <a key={s.url} href={s.url} target="_blank" rel="noreferrer" className="rounded-full border border-white/10 px-2.5 py-1 text-xs text-white/65 hover:bg-white/10">{s.title}</a>)}</div></div>
                  ) : null}
                </div>
              ) : (
                <div className="flex h-40 flex-col items-center justify-center text-center"><BrainCircuit className="size-8 text-white/25" /><p className="mt-3 text-sm font-medium text-white/70">Ask about performance, cash, risk, customers, forecasting or external conditions.</p><p className="mt-1 text-xs text-white/40">Company-specific numbers are grounded in the data available to FinCruiz.</p></div>
              )}
            </div>
          </div>
        </div>
      </section>

      <div className="grid gap-6 xl:grid-cols-[1fr_1fr]">
        <section className="rounded-[2rem] border bg-background p-6 shadow-sm">
          <div className="flex items-center justify-between gap-3"><div><p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Source health</p><h2 className="mt-1 text-xl font-semibold">What FinCruiz can currently see</h2></div><Database className="size-5 text-muted-foreground" /></div>
          <div className="mt-5 grid gap-3 sm:grid-cols-3">
            <div className="rounded-2xl bg-muted/30 p-4"><p className="text-2xl font-semibold">{connected}</p><p className="text-xs text-muted-foreground">Connected systems</p></div>
            <div className="rounded-2xl bg-muted/30 p-4"><p className="text-2xl font-semibold">{totalRecords.toLocaleString()}</p><p className="text-xs text-muted-foreground">Synced source records</p></div>
            <div className="rounded-2xl bg-muted/30 p-4"><p className="text-2xl font-semibold">{data?.memories.length ?? 0}</p><p className="text-xs text-muted-foreground">Management memories</p></div>
          </div>
          <div className="mt-5 space-y-2">
            {data?.source_freshness.length ? data.source_freshness.map((source) => (
              <div key={source.provider} className="flex items-center justify-between gap-3 rounded-xl border p-3 text-sm">
                <div><p className="font-medium capitalize">{source.provider} · {source.name}</p><p className="text-xs text-muted-foreground">{source.last_synced_at ? `Last synced ${new Date(source.last_synced_at).toLocaleString()}` : "Connected · no completed sync yet"}</p></div>
                <span className={`rounded-full px-2 py-1 text-xs ${source.last_sync_status === "success" ? "bg-emerald-50 text-emerald-700" : "bg-muted text-muted-foreground"}`}>{source.last_sync_status || source.status}</span>
              </div>
            )) : <p className="rounded-xl bg-muted/30 p-4 text-sm text-muted-foreground">No ERP source is synchronized yet. Finance uploads can still power the Intelligence Center.</p>}
          </div>
        </section>

        <section className="rounded-[2rem] border bg-background p-6 shadow-sm">
          <div className="flex items-center justify-between gap-3"><div><p className="text-xs font-semibold uppercase tracking-[0.18em] text-muted-foreground">Organizational memory</p><h2 className="mt-1 text-xl font-semibold">Teach FinCruiz what management cares about</h2></div><Lightbulb className="size-5 text-muted-foreground" /></div>
          <p className="mt-2 text-sm text-muted-foreground">Targets, thresholds and decisions give context that accounting data cannot explain by itself.</p>
          <div className="mt-4 grid gap-2 sm:grid-cols-[.8fr_1.2fr]"><input className="rounded-xl border bg-background px-3 py-2 text-sm" placeholder="e.g. Minimum cash threshold" value={title} onChange={(e) => setTitle(e.target.value)} /><input className="rounded-xl border bg-background px-3 py-2 text-sm" placeholder="e.g. Keep cash above A$1m" value={content} onChange={(e) => setContent(e.target.value)} /></div>
          <Button className="mt-3" size="sm" onClick={() => void remember()} disabled={saving || !title.trim() || !content.trim()}>{saving ? <RefreshCw className="mr-2 size-4 animate-spin" /> : <Plus className="mr-2 size-4" />}Remember this</Button>
          <div className="mt-5 space-y-2">{data?.memories.slice(0, 4).map((m) => <div className="rounded-xl bg-muted/35 p-3" key={m.id}><p className="text-sm font-medium">{m.title}</p><p className="mt-1 text-xs leading-5 text-muted-foreground">{m.content}</p></div>)}</div>
        </section>
      </div>

      <section className="rounded-3xl border bg-muted/20 p-5">
        <p className="text-xs leading-5 text-muted-foreground">
          FinCruiz Financial Confidence verifies structural consistency and reconciliation of the data available to the platform. Management insights are decision support, not audit, tax, legal or investment advice. External context is separately identified when used.
        </p>
      </section>
    </div>
  );
}
