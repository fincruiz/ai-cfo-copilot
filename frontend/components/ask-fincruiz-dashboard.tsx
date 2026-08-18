"use client";

import { FormEvent, useState } from "react";
import { useRouter } from "next/navigation";
import { ArrowRight, Bot, ExternalLink, FileCheck2, Gauge, Globe2, Lightbulb, Loader2, Navigation, Send, Sparkles, Target } from "lucide-react";

import { analyticsService } from "@/services/analytics-service";
import { usageService } from "@/services/usage-service";
import { InsightChart } from "@/components/insight-chart";
import type { AICFOAnswer, AICFOConversationTurn } from "@/types/analytics";

const promptGroups = [
  {label:"Understand performance",items:["Why did profit change this month?","Which branch is hurting margins?","What are my fastest-growing expenses?"]},
  {label:"Improve cash",items:["What is putting pressure on cash?","Which customers are driving working-capital pressure?","What happens if debtor days improve by 10 days?"]},
  {label:"Plan ahead",items:["Forecast the next 12 months.","Can we afford to hire 3 people?","What happens if revenue grows 15%?"]},
  {label:"Challenge the business",items:["What should management focus on today?","What are the three biggest financial risks?","What should the board discuss next?"]},
];

export function AskFinCruizDashboard() {
  const router = useRouter();
  const [question, setQuestion] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState<AICFOAnswer | null>(null);
  const [liveContext, setLiveContext] = useState(true);
  const [conversation, setConversation] = useState<AICFOConversationTurn[]>([]);

  async function ask(value: string) {
    const cleaned = value.trim();
    if (!cleaned || loading) return;
    setQuestion(cleaned);
    setLoading(true);
    setResult(null);
    usageService.track("ai_question_submitted", { source: "dashboard_ask_fincruiz" });
    try {
      const response = await analyticsService.askAiCfo(cleaned, liveContext, conversation);
      setResult(response);
      setConversation((current):AICFOConversationTurn[] => [...current, { role: "user" as const, content: cleaned }, { role: "assistant" as const, content: response.answer }].slice(-8));
    } catch {
      setResult({
        answer: "I could not retrieve the company context just now. Check that the workspace data and backend are available, then try again.",
        mode: "error",
        suggested_questions: [],
        sources: [],
        external_context_used: false,
      });
    } finally {
      setLoading(false);
    }
  }

  async function submit(event: FormEvent) {
    event.preventDefault();
    await ask(question);
  }

  return (
    <section className="overflow-hidden rounded-[28px] border border-indigo-200/70 bg-gradient-to-br from-indigo-50 via-background to-sky-50 shadow-sm dark:border-indigo-500/20 dark:from-indigo-950/30 dark:via-background dark:to-sky-950/20">
      <div className="grid gap-0 lg:grid-cols-[.72fr_1.28fr]">
        <div className="border-b border-indigo-100/80 p-6 dark:border-white/10 lg:border-b-0 lg:border-r">
          <div className="flex items-center gap-3">
            <div className="flex size-12 items-center justify-center rounded-2xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-lg"><Bot className="size-5" /></div>
            <div><p className="text-xs font-semibold uppercase tracking-[.18em] text-indigo-700 dark:text-indigo-300">Conversational BI</p><h2 className="text-xl font-semibold">Ask FinCruiz</h2></div>
          </div>
          <p className="mt-4 text-sm leading-6 text-muted-foreground">Ask the management question in plain English. FinCruiz selects the relevant finance evidence, visualizes it where useful, and routes decisions into the appropriate model.</p>
          <div className="mt-5 space-y-4">
            {promptGroups.map((group)=><div key={group.label}><p className="mb-2 text-[10px] font-bold uppercase tracking-[.14em] text-muted-foreground">{group.label}</p><div className="space-y-2">{group.items.map((prompt) => <button key={prompt} type="button" onClick={() => void ask(prompt)} className="flex w-full items-center justify-between gap-3 rounded-2xl border border-indigo-200 bg-white/80 px-3.5 py-2.5 text-left text-xs font-medium text-slate-700 transition hover:-translate-y-0.5 hover:border-indigo-400 hover:shadow-sm dark:border-white/10 dark:bg-white/5 dark:text-slate-200"><span>{prompt}</span><ArrowRight className="size-3.5 shrink-0"/></button>)}</div></div>)}
          </div>
        </div>

        <div className="p-6">
          <form onSubmit={submit}>
            <label className="block text-sm font-semibold">What do you want to understand or decide?</label>
            <div className="mt-3 flex gap-2 rounded-2xl border bg-background/90 p-2 shadow-sm focus-within:ring-2 focus-within:ring-indigo-500/20">
              <input value={question} onChange={(event) => setQuestion(event.target.value)} placeholder="e.g. Why is profit up but cash down?" className="min-w-0 flex-1 bg-transparent px-3 py-2 text-sm outline-none placeholder:text-muted-foreground" />
              <button type="submit" disabled={loading || !question.trim()} className="flex size-11 shrink-0 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-600 to-sky-500 text-white shadow-sm transition hover:scale-[1.03] disabled:opacity-40" aria-label="Ask FinCruiz">{loading ? <Loader2 className="size-4 animate-spin" /> : <Send className="size-4" />}</button>
            </div>
          </form>

          <label className="mt-3 flex cursor-pointer items-center gap-2 text-xs text-muted-foreground"><input type="checkbox" checked={liveContext} onChange={(event) => setLiveContext(event.target.checked)} /><Globe2 className="size-3.5" /> Use live industry and economic context when relevant</label>

          <div className="mt-5 min-h-32 rounded-2xl border border-dashed bg-background/60 p-4">
            {loading ? <div className="flex items-center gap-2 text-sm text-muted-foreground"><Loader2 className="size-4 animate-spin" /> Reviewing company evidence and choosing the right analysis...</div> : result ? (
              <div className="animate-step-in">
                <div className="flex flex-wrap items-center justify-between gap-2"><div className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[.14em] text-indigo-700 dark:text-indigo-300"><Sparkles className="size-3.5" /> FinCruiz response</div><span className="rounded-full border px-2.5 py-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">{result.mode?.replaceAll("_", " ") || "analysis"}</span></div>
                {result.interpreted_question && result.interpreted_question !== question ? <p className="mt-3 text-[11px] text-muted-foreground">Following the earlier question, FinCruiz interpreted this as: <span className="font-medium text-foreground">{result.interpreted_question}</span></p> : null}
                <div className="mt-3 rounded-2xl border border-indigo-200/70 bg-gradient-to-br from-white to-indigo-50/60 p-4 dark:border-indigo-500/20 dark:from-white/[.04] dark:to-indigo-950/20">
                  <p className="flex items-center gap-2 text-[11px] font-semibold uppercase tracking-[.14em] text-indigo-700 dark:text-indigo-300"><Lightbulb className="size-3.5"/>Management answer</p>
                  <p className="mt-2 whitespace-pre-wrap text-[15px] font-medium leading-7">{result.answer}</p>
                </div>
                {result.visualization ? <div className="mt-4"><div className="mb-2 flex items-center gap-2 text-xs font-semibold text-muted-foreground"><Target className="size-3.5"/>Evidence visual</div><InsightChart visualization={result.visualization} /></div> : null}

                {(result.evidence?.length || result.confidence) ? <div className="mt-4 rounded-2xl border bg-white/70 p-4 dark:bg-white/[.03]">
                  <div className="flex flex-wrap items-center justify-between gap-2"><p className="flex items-center gap-2 text-xs font-semibold uppercase tracking-[.12em] text-muted-foreground"><FileCheck2 className="size-3.5"/>Evidence behind this answer</p>{result.confidence ? <span className={`inline-flex items-center gap-1.5 rounded-full px-2.5 py-1 text-[10px] font-semibold uppercase ${result.confidence === "high" ? "bg-emerald-100 text-emerald-700" : result.confidence === "low" ? "bg-rose-100 text-rose-700" : "bg-amber-100 text-amber-700"}`}><Gauge className="size-3"/>{result.confidence} confidence</span> : null}</div>
                  {result.evidence?.length ? <div className="mt-3 grid gap-2 sm:grid-cols-2">{result.evidence.map((item, index) => <div key={`${item.label}-${index}`} className="rounded-xl border bg-background px-3 py-2.5"><div className="flex items-start justify-between gap-3"><span className="text-xs text-muted-foreground">{item.label}</span><span className="text-sm font-semibold tabular-nums">{item.value}</span></div><p className="mt-1 text-[10px] text-muted-foreground">{item.source}{item.period ? ` · ${item.period}` : ""}</p></div>)}</div> : null}
                  {result.confidence_reason ? <p className="mt-3 text-[11px] leading-5 text-muted-foreground">{result.confidence_reason}</p> : null}
                </div> : null}

                <div className="mt-4 flex flex-wrap gap-2">
                  {result.decision_handoff ? <button type="button" onClick={() => { usageService.track("ai_decision_handoff_opened", { source: "dashboard_ask_fincruiz", feature: result.decision_handoff!.title }); router.push(result.decision_handoff!.route); }} className="inline-flex items-center gap-2 rounded-xl bg-gradient-to-r from-violet-600 to-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:brightness-110"><Sparkles className="size-4"/> Model this decision <ArrowRight className="size-4"/></button> : result.action ? <button type="button" onClick={() => { usageService.track("ai_recommended_tool_opened", { source: "dashboard_ask_fincruiz", feature: result.action!.label }); router.push(result.action!.route); }} className="inline-flex items-center gap-2 rounded-xl bg-indigo-600 px-4 py-2 text-sm font-semibold text-white shadow-sm transition hover:bg-indigo-500"><Navigation className="size-4" /> {result.action.label} <ArrowRight className="size-4" /></button> : null}
                  {result.external_context_used ? <span className="inline-flex items-center gap-1.5 rounded-xl border bg-background px-3 py-2 text-xs text-muted-foreground"><Globe2 className="size-3.5"/>External context used</span> : null}
                </div>

                {result.suggested_questions?.length ? <div className="mt-5 border-t pt-4"><p className="text-xs font-semibold uppercase tracking-[.12em] text-muted-foreground">Continue the analysis</p><div className="mt-2 flex flex-wrap gap-2">{result.suggested_questions.slice(0,4).map((suggestion) => <button key={suggestion} type="button" onClick={() => void ask(suggestion)} className="rounded-full border bg-background px-3 py-1.5 text-xs transition hover:border-indigo-400 hover:bg-indigo-50 dark:hover:bg-indigo-950/20">{suggestion}</button>)}</div></div> : null}

                {result.sources?.length ? <div className="mt-5 border-t pt-4"><p className="text-xs font-semibold uppercase tracking-[.12em] text-muted-foreground">External sources</p><div className="mt-2 grid gap-2 sm:grid-cols-2">{result.sources.slice(0,4).map((source) => <a key={source.url} href={source.url} target="_blank" rel="noreferrer" className="flex items-start gap-2 rounded-xl border bg-background p-3 text-xs hover:border-indigo-300"><ExternalLink className="mt-0.5 size-3.5 shrink-0 text-indigo-500"/><span className="line-clamp-2">{source.title}</span></a>)}</div></div> : null}
              </div>
            ) : <p className="text-sm leading-6 text-muted-foreground">Your answer will appear here. Financial values come from FinCruiz's prepared finance context; the text of your question is not sent to product-usage analytics.</p>}
          </div>
        </div>
      </div>
    </section>
  );
}
