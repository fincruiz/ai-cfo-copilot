"use client";

import { useEffect, useMemo, useState } from "react";
import { CheckCircle2, Eye, Loader2, RefreshCcw, Search } from "lucide-react";

import { Button } from "@/components/ui/button";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Input } from "@/components/ui/input";
import { betaFeedbackService, type BetaFeedbackItem } from "@/services/beta-feedback-service";
import { marketingService, type MarketingFunnel } from "@/services/marketing-service";
import { usageService } from "@/services/usage-service";

export default function BetaFeedbackPage() {
  const [items, setItems] = useState<BetaFeedbackItem[]>([]);
  const [summary, setSummary] = useState<any>(null);
  const [funnel, setFunnel] = useState<any>(null);
  const [marketing, setMarketing] = useState<MarketingFunnel | null>(null);
  const [loading, setLoading] = useState(true);
  const [errors, setErrors] = useState<Record<string, string>>({});
  const [query, setQuery] = useState("");

  async function load() {
    setLoading(true);
    setErrors({});
    const results = await Promise.allSettled([
      betaFeedbackService.list(),
      betaFeedbackService.summary(),
      usageService.funnel(30),
      marketingService.funnel(30),
    ]);

    const next: Record<string, string> = {};
    if (results[0].status === "fulfilled") setItems(results[0].value);
    else next.feedback = "Feedback list could not be loaded.";

    if (results[1].status === "fulfilled") setSummary(results[1].value);
    else next.summary = "Feedback summary could not be loaded.";

    if (results[2].status === "fulfilled") setFunnel(results[2].value);
    else next.telemetry = "Product usage telemetry is temporarily unavailable.";

    if (results[3].status === "fulfilled") setMarketing(results[3].value);
    else next.marketing = "Commercial conversion telemetry is not available yet.";

    setErrors(next);
    setLoading(false);
  }

  useEffect(() => {
    void load();
  }, []);

  const visible = useMemo(
    () =>
      items.filter(
        (item) =>
          !query.trim() ||
          `${item.title} ${item.description} ${item.category} ${item.status}`
            .toLowerCase()
            .includes(query.toLowerCase()),
      ),
    [items, query],
  );

  async function update(item: BetaFeedbackItem, status: string) {
    await betaFeedbackService.update(item.id, status, item.resolution_notes || undefined);
    await load();
  }

  async function viewAttachment(id: string) {
    const blob = await betaFeedbackService.attachment(id);
    const url = URL.createObjectURL(blob);
    window.open(url, "_blank", "noopener,noreferrer");
    window.setTimeout(() => URL.revokeObjectURL(url), 60_000);
  }

  return (
    <div className="mx-auto max-w-7xl space-y-6">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <div>
          <p className="text-sm text-muted-foreground">Controlled beta</p>
          <h1 className="text-3xl font-semibold">Testing, product & commercial feedback</h1>
          <p className="mt-2 max-w-3xl text-sm text-muted-foreground">
            Prioritise P0/P1 issues, watch authenticated product adoption and see whether prospects are moving through the public homepage and guided demo. Finance values, contact details and AI question text are excluded from this telemetry.
          </p>
        </div>
        <Button variant="outline" onClick={() => void load()} disabled={loading}>
          <RefreshCcw className={`size-4 ${loading ? "animate-spin" : ""}`} />Refresh
        </Button>
      </div>

      {Object.values(errors).length ? (
        <div className="grid gap-2">
          {Object.entries(errors).map(([key, value]) => (
            <p key={key} className="rounded-xl border border-amber-200 bg-amber-50 p-3 text-sm text-amber-900">{value}</p>
          ))}
        </div>
      ) : null}

      <div className="grid gap-4 sm:grid-cols-2 lg:grid-cols-5">
        {[
          ["Open", summary?.open ?? 0],
          ["P0 open", summary?.p0_open ?? 0],
          ["P1 open", summary?.p1_open ?? 0],
          ["Fixed", summary?.fixed ?? 0],
          ["Total", summary?.total ?? 0],
        ].map(([label, value]) => (
          <Card key={String(label)}><CardHeader className="pb-2"><CardDescription>{label}</CardDescription><CardTitle>{value}</CardTitle></CardHeader></Card>
        ))}
      </div>

      <div className="grid gap-4 xl:grid-cols-3">
        <Card>
          <CardHeader><CardTitle>30-day beta usage</CardTitle><CardDescription>Authenticated product events.</CardDescription></CardHeader>
          <CardContent className="grid gap-3 sm:grid-cols-2">
            {[
              ["Active users", funnel?.active_users ?? 0],
              ["Page views", funnel?.page_views ?? 0],
              ["AI questions", funnel?.ai_questions ?? 0],
              ["Upload starts", funnel?.upload_starts ?? 0],
              ["Upload completions", funnel?.upload_completions ?? 0],
              ["Front-end errors", funnel?.frontend_errors ?? 0],
            ].map(([label, value]) => <Metric key={String(label)} label={String(label)} value={value} />)}
          </CardContent>
        </Card>

        <Card>
          <CardHeader><CardTitle>Homepage conversion</CardTitle><CardDescription>Anonymous public-site session counts.</CardDescription></CardHeader>
          <CardContent className="grid gap-3 sm:grid-cols-2">
            {[
              ["Visitors", marketing?.visitors ?? 0],
              ["Hero demo", marketing?.hero_demo ?? 0],
              ["Hero signup", marketing?.hero_signup ?? 0],
              ["Homepage AI questions", marketing?.ai_questions ?? 0],
              ["AI → signup", marketing?.ai_signup ?? 0],
              ["Pricing clicks", marketing?.pricing ?? 0],
            ].map(([label, value]) => <Metric key={String(label)} label={String(label)} value={value} />)}
          </CardContent>
        </Card>

        <Card className="border-indigo-200/70">
          <CardHeader><CardTitle>Guided demo engagement</CardTitle><CardDescription>Useful for sales conversations and public demo validation.</CardDescription></CardHeader>
          <CardContent className="grid gap-3 sm:grid-cols-2 xl:grid-cols-1">
            <Metric label="Demo views" value={marketing?.demo_views ?? 0} />
            <Metric label="Questions asked" value={marketing?.demo_questions ?? 0} />
            <Metric label="Demo → workspace" value={marketing?.demo_signup ?? 0} />
          </CardContent>
        </Card>
      </div>

      <div className="flex items-center gap-2 rounded-xl border bg-background px-3">
        <Search className="size-4 text-muted-foreground" />
        <Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Search feedback…" className="border-0 shadow-none focus-visible:ring-0" />
      </div>

      <Card>
        <CardHeader>
          <CardTitle>Reported issues & ideas</CardTitle>
          <CardDescription>Use the floating Feedback button anywhere in the dashboard. P0 = security/wrong data/core blocker · P1 = must fix before paid launch · P2 = improvement.</CardDescription>
        </CardHeader>
        <CardContent className="space-y-3">
          {loading ? (
            <div className="flex items-center gap-2 p-5 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" />Loading reports…</div>
          ) : visible.length ? (
            visible.map((item) => (
              <div key={item.id} className="rounded-2xl border p-4">
                <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
                  <div className="min-w-0">
                    <div className="flex flex-wrap items-center gap-2">
                      <span className={`rounded-full px-2.5 py-1 text-[10px] font-bold uppercase ${item.severity === "p0" ? "bg-red-100 text-red-700" : item.severity === "p1" ? "bg-amber-100 text-amber-700" : "bg-slate-100 text-slate-700"}`}>{item.severity}</span>
                      <span className="rounded-full border px-2.5 py-1 text-[10px] font-bold uppercase">{item.status}</span>
                      <span className="text-xs text-muted-foreground">{item.category.replaceAll("_", " ")}</span>
                    </div>
                    <p className="mt-2 font-semibold">{item.title}</p>
                    <p className="mt-1 text-sm leading-6 text-muted-foreground">{item.description}</p>
                    <p className="mt-2 text-xs text-muted-foreground">{item.reporter_name || "Tester"} · {item.user_role || "member"} · {item.path} · {new Date(item.created_at).toLocaleString()}</p>
                    {item.has_attachment ? <button type="button" className="mt-2 inline-flex items-center gap-1 text-xs font-semibold text-primary" onClick={() => void viewAttachment(item.id)}><Eye className="size-3.5" />View authenticated screenshot</button> : null}
                  </div>
                  <div className="flex shrink-0 flex-wrap gap-2">
                    <Button size="sm" variant={item.status === "reviewing" ? "default" : "outline"} onClick={() => void update(item, "reviewing")}>Reviewing</Button>
                    <Button size="sm" variant={item.status === "fixed" ? "default" : "outline"} onClick={() => void update(item, "fixed")}><CheckCircle2 className="size-3.5" />Fixed</Button>
                    <Button size="sm" variant="outline" onClick={() => void update(item, "closed")}>Close</Button>
                  </div>
                </div>
              </div>
            ))
          ) : (
            <p className="p-5 text-sm text-muted-foreground">No feedback has been reported yet.</p>
          )}
        </CardContent>
      </Card>
    </div>
  );
}

function Metric({ label, value }: { label: string; value: unknown }) {
  return (
    <div className="rounded-xl border p-3">
      <p className="text-xs text-muted-foreground">{label}</p>
      <p className="mt-1 text-xl font-semibold">{String(value ?? 0)}</p>
    </div>
  );
}
