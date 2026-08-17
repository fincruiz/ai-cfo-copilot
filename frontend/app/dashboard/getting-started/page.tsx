"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { ArrowRight, BarChart3, Building2, Check, CheckCircle2, Circle, Database, Loader2, RefreshCw, ShieldCheck, Sparkles, WandSparkles } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { getApiErrorMessage } from "@/lib/api";
import { workspaceService, type CommercialOnboardingSummary } from "@/services/workspace-service";

function fmt(value: number | null | undefined) {
  return typeof value === "number" ? value.toLocaleString() : "—";
}

export default function GettingStartedPage() {
  const [summary, setSummary] = useState<CommercialOnboardingSummary | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  async function load() {
    setLoading(true);
    setError("");
    try {
      setSummary(await workspaceService.getCommercialOnboardingSummary());
    } catch (loadError) {
      setError(getApiErrorMessage(loadError));
    } finally {
      setLoading(false);
    }
  }

  useEffect(() => { void load(); }, []);

  function ask(question: string) {
    window.dispatchEvent(new CustomEvent("fincruiz:open-ai", { detail: { question } }));
  }

  if (loading && !summary) {
    return <div className="flex min-h-[60vh] items-center justify-center gap-2 text-muted-foreground"><Loader2 className="size-5 animate-spin" />FinCruiz is reviewing your workspace…</div>;
  }

  return (
    <div className="mx-auto max-w-6xl space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <p className="text-sm font-medium text-primary">Guided setup · workspace build</p>
          <h1 className="mt-1 text-3xl font-semibold tracking-tight">{summary?.ready_for_intelligence ? "Your FinCruiz workspace is ready" : "FinCruiz is building your management workspace"}</h1>
          <p className="mt-2 max-w-3xl text-muted-foreground">We verify the active ledger, business structure and mappings before producing management intelligence. Nothing here is inferred from a blank model.</p>
        </div>
        <Button variant="outline" onClick={() => void load()} disabled={loading}><RefreshCw className={`size-4 ${loading ? "animate-spin" : ""}`} />Refresh review</Button>
      </div>

      {error ? <div className="rounded-xl border border-destructive/30 bg-destructive/5 p-4 text-sm text-destructive">{error}</div> : null}

      {summary ? <>
        <Card className="overflow-hidden border-primary/20">
          <CardHeader className="bg-gradient-to-r from-primary/[.08] via-background to-sky-500/[.05]">
            <div className="flex flex-wrap items-center justify-between gap-3">
              <div><CardTitle>{summary.completed_steps} of {summary.total_steps} workspace checks complete</CardTitle><CardDescription>FinCruiz will only show the first briefing after the finance structure is ready.</CardDescription></div>
              <div className="rounded-full border bg-background px-4 py-1.5 text-sm font-bold">{summary.progress_percent}%</div>
            </div>
            <div className="mt-4 h-2 overflow-hidden rounded-full bg-muted"><div className="h-full rounded-full bg-primary transition-all" style={{ width: `${summary.progress_percent}%` }} /></div>
          </CardHeader>
          <CardContent className="grid gap-2 p-5 md:grid-cols-5">
            {summary.steps.map((step) => <div key={step.key} className={`rounded-xl border p-3 ${step.complete ? "bg-emerald-50/60" : "bg-background"}`}><div className="flex items-center gap-2 text-sm font-medium">{step.complete ? <CheckCircle2 className="size-4 text-emerald-600" /> : <Circle className="size-4 text-muted-foreground" />}{step.label}</div></div>)}
          </CardContent>
        </Card>

        <div className="grid gap-3 sm:grid-cols-3 lg:grid-cols-6">
          {[
            [Database, "Transactions", fmt(summary.transaction_count)],
            [BarChart3, "History", summary.months_history ? `${summary.months_history} months` : "—"],
            [WandSparkles, "Accounts", fmt(summary.account_count)],
            [Check, "Mapped", fmt(summary.mapping_count)],
            [Building2, "Branches", fmt(summary.branch_count)],
            [ShieldCheck, "Confidence", summary.financial_confidence_score != null ? `${summary.financial_confidence_score}/100 · ${summary.financial_confidence_grade ?? ""}` : "Pending"],
          ].map(([Icon, label, value]: any) => <Card key={label}><CardContent className="p-4"><Icon className="size-4 text-primary" /><p className="mt-3 text-xs text-muted-foreground">{label}</p><p className="mt-1 font-semibold">{value}</p></CardContent></Card>)}
        </div>

        {!summary.ready_for_intelligence ? (
          <Card className="border-amber-200 bg-amber-50/40">
            <CardHeader>
              <CardTitle>One clear next step</CardTitle>
              <CardDescription>
                {summary.stage === "branch_review_required"
                  ? `${summary.pending_branch_count} detected branch${summary.pending_branch_count === 1 ? "" : "es"} need confirmation before branch-level intelligence is trusted.`
                  : summary.stage === "mapping_required"
                    ? `${summary.unmapped_account_count} account${summary.unmapped_account_count === 1 ? "" : "s"} still need mapping before FinCruiz can calculate reliable management reports.`
                    : summary.stage === "briefing_pending"
                      ? (summary.briefing_error ?? "The finance structure is ready; retry the first briefing.")
                      : "Load a General Ledger or complete a connected-source sync to start the financial build."}
              </CardDescription>
            </CardHeader>
            <CardContent>
              <Link href={summary.next_path} className="inline-flex items-center gap-2 rounded-xl bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground">{summary.next_label}<ArrowRight className="size-4" /></Link>
              {summary.period_start && summary.period_end ? <p className="mt-4 text-xs text-muted-foreground">Active data coverage: {new Date(summary.period_start).toLocaleDateString()} – {new Date(summary.period_end).toLocaleDateString()}</p> : null}
            </CardContent>
          </Card>
        ) : null}

        {summary.ready_for_intelligence && summary.briefing ? <>
          <Card className="overflow-hidden border-violet-200 bg-gradient-to-br from-violet-500/[.08] via-background to-cyan-500/[.06]">
            <CardHeader>
              <div className="flex items-center gap-2 text-sm font-semibold text-violet-700"><Sparkles className="size-4" />Your first management briefing</div>
              <CardTitle className="mt-2 text-2xl">{summary.briefing.executive_summary?.headline ?? "FinCruiz has analysed the current finance data"}</CardTitle>
              <CardDescription className="max-w-3xl text-sm leading-6">{summary.briefing.executive_summary?.narrative}</CardDescription>
            </CardHeader>
            <CardContent className="grid gap-3 lg:grid-cols-3">
              {(summary.briefing.priorities ?? []).slice(0, 3).map((priority, index) => (
                <div key={`${priority.title}-${index}`} className="rounded-2xl border bg-background/85 p-4">
                  <span className={`rounded-full px-2 py-1 text-[11px] font-semibold ${priority.level === "critical" ? "bg-red-100 text-red-700" : priority.level === "positive" ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-800"}`}>{priority.level}</span>
                  <p className="mt-3 font-semibold">{priority.title}</p>
                  {priority.evidence ? <p className="mt-2 text-xs leading-5 text-muted-foreground">Evidence · {priority.evidence}</p> : null}
                  {priority.action ? <p className="mt-2 text-xs leading-5"><b>Next:</b> {priority.action}</p> : null}
                </div>
              ))}
            </CardContent>
          </Card>

          <Card>
            <CardHeader><CardTitle>Ask your first management question</CardTitle><CardDescription>These questions open Ask FinCruiz with the current company evidence already available.</CardDescription></CardHeader>
            <CardContent className="flex flex-wrap gap-2">
              {(summary.briefing.suggested_questions ?? []).slice(0, 4).map((question) => <button type="button" key={question} onClick={() => ask(question)} className="rounded-full border bg-background px-4 py-2 text-sm transition hover:border-primary hover:bg-primary/5">{question}</button>)}
            </CardContent>
          </Card>

          <div className="flex flex-wrap justify-end gap-3">
            <Link href="/dashboard/intelligence" className="inline-flex items-center gap-2 rounded-xl border px-4 py-2 text-sm font-semibold">Open Intelligence Center</Link>
            <Link href="/dashboard" className="inline-flex items-center gap-2 rounded-xl bg-primary px-4 py-2 text-sm font-semibold text-primary-foreground">Open management dashboard<ArrowRight className="size-4" /></Link>
          </div>
        </> : null}
      </> : null}
    </div>
  );
}
